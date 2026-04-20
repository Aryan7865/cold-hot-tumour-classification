"""
03_immune_scoring.py
--------------------
OWNER: P2 (Immune phenotyper)

Purpose
    Convert each tumour's gene-expression profile into an 'immune score' that
    reflects the abundance of tumour-infiltrating lymphocytes, then assign
    every patient a  Hot  /  Cold  /  Intermediate  label.

Method   (pure-Python, no R)
    1.  ssGSEA  via gseapy  using 28 immune gene-set signatures
        (Bindea et al. 2013, Charoentong et al. 2017).
    2.  ESTIMATE stromal + immune score (ssGSEA of the 141-gene ESTIMATE set).
    3.  Combine   ->   Hot (top tertile), Cold (bottom tertile),
                        Intermediate (middle).

Inputs   : data/processed/<cohort>_expr_log.tsv
Outputs  : outputs/tables/immune_scores_<cohort>.csv
           outputs/tables/labels_<cohort>.csv          <- the 'ground truth'
           outputs/figures/immune/*.png
"""

from __future__ import annotations
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

import config
from utils import get_logger, save_fig, save_table, set_plot_style

log = get_logger("immune")
set_plot_style()

# ------------------------------------------------------------------
# IMMUNE GENE SETS
#  If you don't have an internet-downloaded GMT, this dict gives a minimal
#  working fallback (Bindea + ESTIMATE, abbreviated).  Replace with full GMT
#  when possible for publication-grade scores.
# ------------------------------------------------------------------
BINDEA_FALLBACK: dict[str, list[str]] = {
    # --- ESTIMATE immune (abridged) ---
    "ESTIMATE_Immune": [
        "CD247", "CD2", "CD3D", "CD3E", "CD3G", "CD4", "CD8A", "CD8B",
        "CD19", "CD79A", "CD79B", "FCER1G", "FCGR1A", "GZMA", "GZMB",
        "GZMH", "GZMK", "GZMM", "HLA-A", "HLA-B", "HLA-C", "HLA-DMA",
        "HLA-DMB", "HLA-DOA", "HLA-DOB", "HLA-DPA1", "HLA-DPB1",
        "HLA-DQA1", "HLA-DQB1", "HLA-DRA", "HLA-DRB1", "IFNG", "IL2RA",
        "IL2RB", "IL7R", "LCK", "LCP2", "LYZ", "NKG7", "PRF1", "PTPRC",
        "PYHIN1", "SELL", "TNF", "TRAC", "TRBC1", "TRBC2",
    ],
    "ESTIMATE_Stromal": [
        "ACTA2", "ADAM12", "ANGPTL1", "BGN", "CALD1", "COL1A1", "COL1A2",
        "COL3A1", "COL4A1", "COL5A1", "COL6A1", "COL6A2", "COL6A3",
        "DCN", "FAP", "FBN1", "FN1", "LUM", "MMP2", "POSTN", "SPARC",
        "TGFB1", "THBS1", "THBS2", "VCAN", "VIM",
    ],
    # --- Bindea cell populations (abridged) ---
    "CD8_Tcells":      ["CD8A", "CD8B", "GZMK", "PRF1", "GZMA"],
    "CytotoxicCells":  ["GZMA", "GZMB", "GZMH", "GZMK", "GZMM", "KLRB1",
                        "KLRD1", "KLRK1", "NKG7", "PRF1"],
    "Tregs":           ["FOXP3", "IL2RA", "CTLA4", "IKZF2", "TNFRSF18"],
    "Th1":             ["IFNG", "TBX21", "STAT4", "IL12RB2"],
    "Th2":             ["IL4", "IL5", "IL13", "GATA3", "STAT6"],
    "Macrophages":     ["CD68", "CD163", "CD14", "FCGR1A", "ITGAM", "MSR1"],
    "NK_cells":        ["NCAM1", "NKG7", "KLRD1", "KLRF1", "NCR1"],
    "Bcells":          ["CD19", "CD79A", "CD79B", "MS4A1", "BANK1"],
    "DCs":             ["ITGAX", "CD1A", "CD1C", "CD83", "CCR7"],
    "Neutrophils":     ["CSF3R", "FCGR3B", "CXCR2", "ELANE", "MPO"],
}


def load_immune_gmt() -> dict[str, list[str]]:
    """Attempt to read a real ssGSEA GMT; fall back to the builtin dict."""
    gmt = config.IMMUNE_GMT_FALLBACK
    if not gmt.exists():
        log.warning("No GMT file found at %s — using built-in fallback.", gmt)
        return BINDEA_FALLBACK
    gene_sets = {}
    with open(gmt) as fh:
        for line in fh:
            tok = line.strip().split("\t")
            if len(tok) >= 3:
                gene_sets[tok[0]] = tok[2:]
    log.info("Loaded %d gene sets from %s", len(gene_sets), gmt.name)
    return gene_sets


# ------------------------------------------------------------------
# ssGSEA via gseapy
# ------------------------------------------------------------------
def run_ssgsea(expr: pd.DataFrame,
               gene_sets: dict[str, list[str]]) -> pd.DataFrame:
    """Return a (sample x gene-set) enrichment matrix via gseapy.ssgsea."""
    try:
        import gseapy as gp
    except ImportError as e:
        raise RuntimeError("Install gseapy:  pip install gseapy") from e

    # gseapy expects genes as index (HGNC symbol), samples as columns.
    log.info("Running ssGSEA on %s samples x %s genes, %d gene sets",
             expr.shape[1], expr.shape[0], len(gene_sets))

    ss = gp.ssgsea(
        data=expr,
        gene_sets=gene_sets,
        sample_norm_method="rank",
        outdir=None,
        no_plot=True,
        min_size=2,
        max_size=5000,
        threads=4,
    )
    # res2d -> rows = (Term, Name),  'ES' column; pivot to sample x term
    res = ss.res2d.pivot_table(index="Name", columns="Term", values="NES")
    res.index.name = "sample"
    return res


# ------------------------------------------------------------------
# LABELLING
# ------------------------------------------------------------------
def assign_hot_cold(scores: pd.DataFrame,
                    score_col: str = "ESTIMATE_Immune",
                    hot_q: float = config.HOT_QUANTILE,
                    cold_q: float = config.COLD_QUANTILE) -> pd.Series:
    """Tertile split on the specified immune score column."""
    if score_col not in scores.columns:
        # Fall back: mean over all immune cell-type columns
        log.warning("'%s' not in scores — averaging all non-Stromal cols.",
                    score_col)
        use = [c for c in scores.columns if "Stromal" not in c]
        s = scores[use].mean(axis=1)
    else:
        s = scores[score_col]

    lo, hi = s.quantile([cold_q, hot_q])
    labels = pd.Series("Intermediate", index=s.index, name="label")
    labels[s >= hi] = "Hot"
    labels[s <= lo] = "Cold"
    log.info("Label distribution:\n%s", labels.value_counts().to_string())
    return labels


# ------------------------------------------------------------------
# PLOTS
# ------------------------------------------------------------------
def plot_score_distribution(scores: pd.DataFrame,
                            labels: pd.Series,
                            cohort: str) -> None:
    df = scores.join(labels)
    immune_cols = [c for c in scores.columns if c.startswith("ESTIMATE")]
    for col in immune_cols:
        fig, ax = plt.subplots(figsize=(8, 5))
        sns.histplot(data=df, x=col, hue="label", multiple="stack",
                     palette=config.PALETTE_LABEL, ax=ax, bins=40)
        ax.set_title(f"{col} distribution — {cohort}")
        save_fig(fig, f"{cohort}_{col}_hist", subfolder="immune")

    fig, ax = plt.subplots(figsize=(10, 6))
    sns.heatmap(scores.T, cmap="RdBu_r", center=0,
                xticklabels=False, ax=ax)
    ax.set_title(f"ssGSEA enrichment heatmap — {cohort}")
    save_fig(fig, f"{cohort}_ssgsea_heatmap", subfolder="immune")


# ------------------------------------------------------------------
# ENTRY POINT
# ------------------------------------------------------------------
def score_cohort(cohort: str) -> pd.DataFrame:
    expr_path = config.PROCESSED_DIR / f"{cohort}_expr_log.tsv"
    if not expr_path.exists():
        log.error("Missing %s — run 02_preprocessing.py first.", expr_path)
        return pd.DataFrame()

    expr = pd.read_csv(expr_path, sep="\t", index_col=0)
    gene_sets = load_immune_gmt()
    scores = run_ssgsea(expr, gene_sets)

    labels = assign_hot_cold(scores)
    save_table(scores, f"immune_scores_{cohort}")
    save_table(labels.to_frame(), f"labels_{cohort}")

    plot_score_distribution(scores, labels, cohort)
    return scores


def main() -> None:
    log.info("=== IMMUNE SCORING ===")
    cohorts = config.TCGA_PROJECTS + [config.VALIDATION_COHORT]
    for c in cohorts:
        if (config.PROCESSED_DIR / f"{c}_expr_log.tsv").exists():
            score_cohort(c)


if __name__ == "__main__":
    main()
