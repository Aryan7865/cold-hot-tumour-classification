"""
06_hypothesis_testing.py
------------------------
OWNER: P4 (Statistician)

Purpose
    Rigorously quantify which genes / clinical variables differ between
    Hot / Cold / Intermediate tumours.

Course concepts exercised
    - Student's t-test  (Hot vs Cold, per gene)
    - ANOVA             (three groups, per gene)
    - F-test            (compare variances between groups)
    - Kruskal-Wallis    (non-parametric ANOVA)
    - Chi-squared       (categorical clinical vars vs label)
    - Fisher exact      (small contingency)
    - Multiple-test correction  :  Bonferroni + Benjamini-Hochberg

Outputs
    outputs/tables/dge_<cohort>.csv           per-gene stats
    outputs/tables/dge_sig_<cohort>.csv       FDR < 0.05, |effect| large
    outputs/tables/chi2_clinical_<cohort>.csv
    outputs/figures/hypothesis/*.png
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from statsmodels.stats.multitest import multipletests

import config
from utils import get_logger, save_fig, save_table, set_plot_style

log = get_logger("hypothesis")
set_plot_style()


# ------------------------------------------------------------------
# LOAD
# ------------------------------------------------------------------
def load_expr_and_labels(cohort: str) -> tuple[pd.DataFrame, pd.Series]:
    expr = pd.read_csv(config.PROCESSED_DIR / f"{cohort}_expr_log.tsv",
                       sep="\t", index_col=0)
    labels = pd.read_csv(config.TABLES_DIR / f"labels_{cohort}.csv",
                         index_col=0)["label"]
    common = expr.columns.intersection(labels.index)
    return expr[common], labels.loc[common]


# ------------------------------------------------------------------
# PER-GENE PARAMETRIC + NON-PARAMETRIC TESTS
# ------------------------------------------------------------------
def gene_level_tests(expr: pd.DataFrame, labels: pd.Series) -> pd.DataFrame:
    """Run t-test (Hot vs Cold), ANOVA, Kruskal, F-test of variance for each gene."""
    hot  = expr.loc[:, labels == "Hot"].values
    cold = expr.loc[:, labels == "Cold"].values
    inter= expr.loc[:, labels == "Intermediate"].values

    log.info("Testing %d genes  |  Hot=%d  Cold=%d  Inter=%d",
             expr.shape[0], hot.shape[1], cold.shape[1], inter.shape[1])

    # Vectorised t-test (Welch) — Hot vs Cold
    t_stat, t_p = stats.ttest_ind(hot, cold, axis=1, equal_var=False,
                                  nan_policy="omit")

    # Vectorised one-way ANOVA across 3 groups
    f_stat, anova_p = stats.f_oneway(hot, cold, inter, axis=1)

    # Non-parametric Kruskal-Wallis (per gene loop — scipy isn't vectorised)
    kw_stat = np.empty(expr.shape[0])
    kw_p    = np.empty(expr.shape[0])
    for i in range(expr.shape[0]):
        kw_stat[i], kw_p[i] = stats.kruskal(hot[i], cold[i], inter[i])

    # F-test of variance (Hot vs Cold)
    var_hot, var_cold = hot.var(axis=1, ddof=1), cold.var(axis=1, ddof=1)
    f_var = var_hot / np.where(var_cold > 0, var_cold, np.nan)
    df1, df2 = hot.shape[1] - 1, cold.shape[1] - 1
    f_var_p = 2 * np.minimum(
        stats.f.cdf(f_var, df1, df2),
        stats.f.sf (f_var, df1, df2),
    )

    mean_hot  = hot.mean(axis=1)
    mean_cold = cold.mean(axis=1)
    logfc = mean_hot - mean_cold      # already in log-space

    df = pd.DataFrame({
        "gene": expr.index,
        "mean_Hot":  mean_hot,
        "mean_Cold": mean_cold,
        "mean_Inter": inter.mean(axis=1),
        "log2FC_HotVsCold": logfc,
        "t_stat": t_stat, "t_p": t_p,
        "anova_F": f_stat, "anova_p": anova_p,
        "kw_stat": kw_stat, "kw_p": kw_p,
        "Fvar": f_var, "Fvar_p": f_var_p,
    }).set_index("gene")

    # Multiple-testing correction (BH + Bonferroni) on ANOVA and t
    for col in ("t_p", "anova_p", "kw_p"):
        df[col + "_bh"]   = multipletests(df[col].fillna(1), method="fdr_bh")[1]
        df[col + "_bonf"] = multipletests(df[col].fillna(1), method="bonferroni")[1]

    return df


# ------------------------------------------------------------------
# VOLCANO + HEATMAP
# ------------------------------------------------------------------
def volcano_plot(df: pd.DataFrame, cohort: str,
                 fc_thr: float = 1.0, p_thr: float = 0.05) -> None:
    d = df.copy()
    d["-log10p"] = -np.log10(d["t_p_bh"].replace(0, 1e-300))
    d["sig"] = ((d["t_p_bh"] < p_thr) & (d["log2FC_HotVsCold"].abs() > fc_thr))

    fig, ax = plt.subplots(figsize=(8, 6))
    ax.scatter(d["log2FC_HotVsCold"], d["-log10p"],
               c=np.where(d["sig"], "red", "lightgrey"),
               s=8, alpha=0.6)
    ax.axhline(-np.log10(p_thr), ls="--", color="black", lw=0.5)
    ax.axvline( fc_thr, ls="--", color="black", lw=0.5)
    ax.axvline(-fc_thr, ls="--", color="black", lw=0.5)
    ax.set_xlabel("log2 FC  (Hot − Cold)")
    ax.set_ylabel("-log10 adjusted p (BH)")
    ax.set_title(f"Volcano — Hot vs Cold, {cohort}")
    save_fig(fig, f"volcano_{cohort}", subfolder="hypothesis")


def top_gene_heatmap(expr: pd.DataFrame, labels: pd.Series,
                     df: pd.DataFrame, cohort: str, top_n: int = 40) -> None:
    top = df.sort_values("t_p_bh").head(top_n).index
    sub = expr.loc[top]
    # z-score per gene
    sub = sub.sub(sub.mean(axis=1), axis=0).div(sub.std(axis=1) + 1e-9, axis=0)
    # order samples by label
    order = labels.sort_values().index
    sub = sub[order]

    col_colors = labels.loc[order].map(config.PALETTE_LABEL)
    g = sns.clustermap(sub, row_cluster=True, col_cluster=False,
                       col_colors=col_colors, cmap="RdBu_r", center=0,
                       xticklabels=False, figsize=(12, 8))
    g.fig.suptitle(f"Top-{top_n} DE genes — {cohort}", y=1.02)
    out = config.FIGURES_DIR / "hypothesis" / f"heatmap_top{top_n}_{cohort}.png"
    out.parent.mkdir(parents=True, exist_ok=True)
    g.savefig(out, bbox_inches="tight", dpi=config.FIG_DPI)
    plt.close(g.fig)
    log.info("Saved heatmap -> %s", out.name)


# ------------------------------------------------------------------
# CHI-SQUARED for clinical variables
# ------------------------------------------------------------------
def chi_square_clinical(cohort: str, labels: pd.Series) -> pd.DataFrame:
    clin_path = config.RAW_DIR / cohort / "clinical.tsv"
    if not clin_path.exists():
        log.info("No clinical data for %s -> skipping chi-squared.", cohort)
        return pd.DataFrame()
    clin = pd.read_csv(clin_path, sep="\t")
    id_col = next((c for c in clin.columns if "submitter" in c.lower()), None)
    if id_col is None:
        log.warning("No submitter_id column; skipping.")
        return pd.DataFrame()
    clin = clin.set_index(id_col)
    common = clin.index.intersection(labels.index)
    clin = clin.loc[common]
    lab  = labels.loc[common]

    rows = []
    for var in clin.columns:
        if clin[var].dtype == object and clin[var].nunique() < 20:
            ct = pd.crosstab(lab, clin[var].fillna("NA"))
            if ct.shape[0] < 2 or ct.shape[1] < 2:
                continue
            chi2, p, dof, _ = stats.chi2_contingency(ct)
            rows.append({"variable": var, "chi2": chi2, "dof": dof, "p": p,
                         "n_categories": ct.shape[1]})
            if var.lower() in {"msi_status", "ajcc_pathologic_stage"}:
                fig, ax = plt.subplots(figsize=(7, 4))
                (ct.div(ct.sum(axis=1), axis=0) * 100).plot(kind="bar", stacked=True,
                                                            ax=ax, colormap="tab20")
                ax.set_title(f"{var} vs Hot/Cold (χ²={chi2:.1f}, p={p:.2g})")
                ax.set_ylabel("%")
                save_fig(fig, f"{cohort}_{var}_chi2", subfolder="hypothesis")

    tbl = pd.DataFrame(rows)
    if len(tbl):
        tbl["p_bh"] = multipletests(tbl["p"], method="fdr_bh")[1]
    save_table(tbl, f"chi2_clinical_{cohort}")
    return tbl


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def process_cohort(cohort: str) -> None:
    log.info("=== %s ===", cohort)
    expr, labels = load_expr_and_labels(cohort)
    df = gene_level_tests(expr, labels)
    save_table(df, f"dge_{cohort}")

    sig = df[(df["t_p_bh"] < config.ALPHA) &
             (df["log2FC_HotVsCold"].abs() > 1.0)]
    log.info("%d genes significant after BH + |logFC|>1", len(sig))
    save_table(sig, f"dge_sig_{cohort}")

    volcano_plot(df, cohort)
    if len(sig):
        top_gene_heatmap(expr, labels, df, cohort)

    chi_square_clinical(cohort, labels)


def main() -> None:
    log.info("=== HYPOTHESIS TESTING ===")
    cohorts = config.TCGA_PROJECTS + [config.VALIDATION_COHORT]
    for c in cohorts:
        if (config.TABLES_DIR / f"labels_{c}.csv").exists():
            process_cohort(c)


if __name__ == "__main__":
    main()
