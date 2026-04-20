"""
02_preprocessing.py
-------------------
OWNER: P1 (Data Engineer)

Purpose
    Turn raw expression files (TCGA counts / GEO intensities) into a single
    analysis-ready matrix:

        (1) load raw counts for each cohort
        (2) filter lowly-expressed genes
        (3) log2(x + 1) transform
        (4) harmonise gene IDs across cohorts
        (5) (optional) take the top-K most variable genes for downstream ML
        (6) check distribution normality    (QQ-plots, Shapiro / KS tests)
        (7) save processed matrices to  data/processed/

Inputs   : data/raw/<cohort>/expression_counts.tsv   (from step 01)
Outputs  : data/processed/<cohort>_expr_log.tsv
           data/processed/<cohort>_expr_topvar.tsv
           outputs/figures/preprocessing/*.png
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

import config
from utils import get_logger, save_fig, set_plot_style

log = get_logger("preprocess")
set_plot_style()


# ------------------------------------------------------------------
# LOADERS
# ------------------------------------------------------------------
def _load_raw(cohort: str) -> pd.DataFrame:
    """Return a genes x samples DataFrame for a cohort."""
    raw = config.RAW_DIR / cohort
    for candidate in ("expression_counts.tsv", "expression.tsv"):
        path = raw / candidate
        if path.exists():
            log.info("Loading %s/%s", cohort, candidate)
            return pd.read_csv(path, sep="\t", index_col=0)
    raise FileNotFoundError(f"No expression file found for {cohort}")


# ------------------------------------------------------------------
# CORE STEPS
# ------------------------------------------------------------------
def filter_low_expression(expr: pd.DataFrame,
                          min_cpm: float = config.MIN_EXPR_CPM,
                          min_pct: float = config.MIN_PCT_SAMPLES) -> pd.DataFrame:
    """Keep genes with >= min_cpm in at least min_pct samples."""
    cpm = expr.div(expr.sum(axis=0), axis=1) * 1e6
    keep = (cpm >= min_cpm).mean(axis=1) >= min_pct
    log.info("Filter: %d / %d genes kept", int(keep.sum()), len(keep))
    return expr.loc[keep]


def log_transform(expr: pd.DataFrame,
                  pseudocount: float = config.LOG_PSEUDOCOUNT) -> pd.DataFrame:
    """log2(x + pseudocount).  Idempotent for already-log-transformed microarray."""
    if (expr.values < 0).any():
        log.warning("Negative values present -> assuming already log-space. Skipping.")
        return expr
    if expr.max().max() < 50:
        log.warning("Values small (max %.2f) -> assuming already log-space. Skipping.",
                    expr.max().max())
        return expr
    return np.log2(expr + pseudocount)


def top_variance(expr: pd.DataFrame,
                 k: int = config.VARIANCE_TOP_K) -> pd.DataFrame:
    """Return top-k most-variable genes (row-wise variance)."""
    var = expr.var(axis=1).sort_values(ascending=False)
    return expr.loc[var.head(k).index]


# ------------------------------------------------------------------
# DISTRIBUTION CHECKS (course concept: Standard Normal Distribution)
# ------------------------------------------------------------------
def distribution_checks(expr: pd.DataFrame, cohort: str,
                        n_genes: int = 5) -> pd.DataFrame:
    """Run Shapiro-Wilk on a sample of genes + plot QQ + histogram."""
    rng = np.random.default_rng(config.RANDOM_STATE)
    sample_idx = rng.choice(expr.index, size=min(n_genes, len(expr)),
                            replace=False)

    rows = []
    fig, axes = plt.subplots(2, n_genes, figsize=(4 * n_genes, 7))
    if n_genes == 1:
        axes = axes[:, None]

    for j, g in enumerate(sample_idx):
        values = expr.loc[g].dropna().values
        if len(values) < 3:
            continue
        sh_stat, sh_p = stats.shapiro(values[:5000])    # shapiro has n-cap
        ks_stat, ks_p = stats.kstest(
            (values - values.mean()) / (values.std(ddof=0) + 1e-9), "norm"
        )
        rows.append({"gene": g,
                     "shapiro_stat": sh_stat, "shapiro_p": sh_p,
                     "ks_stat": ks_stat,     "ks_p": ks_p,
                     "n": len(values)})

        axes[0, j].hist(values, bins=40, color="steelblue", edgecolor="white")
        axes[0, j].set_title(f"{g}")
        axes[0, j].set_xlabel("log2 expr")
        stats.probplot(values, dist="norm", plot=axes[1, j])
        axes[1, j].set_title("QQ plot")

    fig.suptitle(f"Distribution checks — {cohort}", fontsize=16)
    save_fig(fig, f"dist_check_{cohort}", subfolder="preprocessing")

    return pd.DataFrame(rows)


# ------------------------------------------------------------------
# HARMONISATION (across cohorts)
# ------------------------------------------------------------------
def harmonise_genes(matrices: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Keep only genes present in every cohort (intersection)."""
    common = set.intersection(*[set(m.index) for m in matrices.values()])
    log.info("Intersection of %d cohorts -> %d common genes",
             len(matrices), len(common))
    common = sorted(common)
    return {name: m.loc[common] for name, m in matrices.items()}


# ------------------------------------------------------------------
# ENTRY POINT
# ------------------------------------------------------------------
def process_cohort(cohort: str) -> pd.DataFrame:
    expr = _load_raw(cohort)
    expr = filter_low_expression(expr)
    expr = log_transform(expr)

    # Save full log-expression
    out = config.PROCESSED_DIR / f"{cohort}_expr_log.tsv"
    expr.to_csv(out, sep="\t")
    log.info("Saved -> %s (%s genes x %s samples)",
             out.name, *expr.shape)

    # Distribution sanity check
    dist_tbl = distribution_checks(expr, cohort)
    dist_tbl.to_csv(config.TABLES_DIR / f"distribution_{cohort}.csv", index=False)

    # Top-variance subset
    top = top_variance(expr)
    top.to_csv(config.PROCESSED_DIR / f"{cohort}_expr_topvar.tsv", sep="\t")

    return expr


def main() -> None:
    log.info("=== PREPROCESSING ===")
    cohorts = config.TCGA_PROJECTS + [config.GEO_VALIDATION, config.GEO_SURVIVAL]
    matrices = {c: process_cohort(c) for c in cohorts
                if (config.RAW_DIR / c).exists()}

    if len(matrices) >= 2:
        harmonised = harmonise_genes(matrices)
        for name, m in harmonised.items():
            m.to_csv(config.PROCESSED_DIR / f"{name}_expr_harmonised.tsv",
                     sep="\t")


if __name__ == "__main__":
    main()
