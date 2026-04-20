"""
08_survival_analysis.py
-----------------------
OWNER: P6 (Survival + Validation)

Purpose
    Link the immune phenotype we derived to clinical outcome:
        - Kaplan-Meier curves for Hot vs Cold vs Intermediate
        - Log-rank test
        - Cox proportional-hazards regression (multivariate)
        - PLS regression  (immune score  ->  survival time)

Course concepts exercised
    - Regression analysis (Cox, PLS)
    - Hypothesis testing (log-rank)

Inputs   : outputs/tables/labels_<cohort>.csv
           outputs/tables/immune_scores_<cohort>.csv
           data/raw/<cohort>/clinical.tsv (TCGA) or phenotype.tsv (GEO)

Outputs  : outputs/figures/survival/*.png
           outputs/tables/cox_<cohort>.csv
           outputs/tables/logrank_<cohort>.csv
           outputs/tables/pls_<cohort>.csv
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from sklearn.cross_decomposition import PLSRegression
from sklearn.preprocessing import StandardScaler

import config
from utils import get_logger, save_fig, save_table, set_plot_style

log = get_logger("survival")
set_plot_style()


# ------------------------------------------------------------------
# CLINICAL LOADER
# ------------------------------------------------------------------
def _load_clinical(cohort: str) -> pd.DataFrame:
    """Return a DataFrame indexed by sample id with columns
       'time' (days) and 'event' (1 = death)."""
    tcga = config.RAW_DIR / cohort / "clinical.tsv"
    geo  = config.RAW_DIR / cohort / "phenotype.tsv"

    if tcga.exists():
        df = pd.read_csv(tcga, sep="\t")
        id_col = next((c for c in df.columns if "submitter" in c.lower()), None)
        df = df.set_index(id_col)
        # Consolidate survival fields
        time = df.get("demographic.days_to_death")
        followup = df.get("diagnoses.days_to_last_follow_up")
        vital = df.get("demographic.vital_status", "").fillna("").astype(str)
        df["time"]  = time.fillna(followup)
        df["event"] = vital.str.lower().eq("dead").astype(int)
        return df[["time", "event"]].dropna()

    if geo.exists():
        pheno = pd.read_csv(geo, sep="\t", index_col=0)
        log.warning("Survival parsing for GEO %s is cohort-specific; "
                    "edit this function manually.", cohort)
        return pheno  # team-member can adapt per cohort
    return pd.DataFrame()


# ------------------------------------------------------------------
# KAPLAN-MEIER + LOG-RANK
# ------------------------------------------------------------------
def kaplan_meier(cohort: str) -> None:
    try:
        from lifelines import KaplanMeierFitter
        from lifelines.statistics import multivariate_logrank_test
    except ImportError as e:
        raise RuntimeError("pip install lifelines") from e

    clin = _load_clinical(cohort)
    if clin.empty or "time" not in clin or "event" not in clin:
        log.warning("No usable survival data for %s", cohort)
        return
    labels = pd.read_csv(config.TABLES_DIR / f"labels_{cohort}.csv",
                         index_col=0)["label"]
    common = clin.index.intersection(labels.index)
    clin = clin.loc[common].copy()
    clin["label"] = labels.loc[common]
    clin = clin.dropna(subset=["time", "event", "label"])
    clin["time"]  = pd.to_numeric(clin["time"],  errors="coerce")
    clin["event"] = pd.to_numeric(clin["event"], errors="coerce")
    clin = clin.dropna()

    if len(clin) < 10:
        log.warning("Only %d samples with survival -> skipping KM.", len(clin))
        return

    fig, ax = plt.subplots(figsize=(8, 5))
    kmf = KaplanMeierFitter()
    for grp, sub in clin.groupby("label"):
        kmf.fit(sub["time"], sub["event"], label=f"{grp} (n={len(sub)})")
        kmf.plot_survival_function(ax=ax, ci_show=True,
                                    color=config.PALETTE_LABEL.get(grp, "grey"))
    res = multivariate_logrank_test(clin["time"], clin["label"], clin["event"])
    ax.set_title(f"KM — {cohort}  (log-rank p={res.p_value:.2g})")
    ax.set_xlabel("Days"); ax.set_ylabel("Survival probability")
    save_fig(fig, f"{cohort}_kaplan_meier", subfolder="survival")

    save_table(pd.DataFrame({"test_statistic": [res.test_statistic],
                             "p_value": [res.p_value]}),
               f"logrank_{cohort}")


# ------------------------------------------------------------------
# COX  (multivariate)
# ------------------------------------------------------------------
def cox_regression(cohort: str) -> None:
    try:
        from lifelines import CoxPHFitter
    except ImportError as e:
        raise RuntimeError("pip install lifelines") from e

    clin  = _load_clinical(cohort)
    if clin.empty:
        return
    scores = pd.read_csv(config.TABLES_DIR / f"immune_scores_{cohort}.csv",
                         index_col=0)

    common = clin.index.intersection(scores.index)
    df = clin.loc[common].join(scores.loc[common])
    df = df.apply(pd.to_numeric, errors="coerce").dropna()
    if len(df) < 20:
        log.warning("Insufficient Cox data (n=%d) for %s", len(df), cohort)
        return

    # Keep up to 5 most-variance covariates to avoid singular design
    covariates = (df.drop(columns=["time", "event"])
                    .var().sort_values(ascending=False).head(5).index.tolist())
    cph = CoxPHFitter(penalizer=0.1)
    cph.fit(df[covariates + ["time", "event"]],
            duration_col="time", event_col="event")
    log.info("Cox summary:\n%s", cph.summary)
    save_table(cph.summary, f"cox_{cohort}")


# ------------------------------------------------------------------
# PLS
# ------------------------------------------------------------------
def pls_regression(cohort: str) -> None:
    clin = _load_clinical(cohort)
    if clin.empty:
        return
    expr = pd.read_csv(config.PROCESSED_DIR / f"{cohort}_expr_topvar.tsv",
                       sep="\t", index_col=0)

    common = clin.index.intersection(expr.columns)
    y = pd.to_numeric(clin.loc[common, "time"], errors="coerce")
    X = expr[common].T
    mask = y.notna()
    X, y = X[mask], y[mask]
    if len(y) < 20:
        log.warning("PLS skipped -> too few samples (%d)", len(y))
        return

    Xs = StandardScaler().fit_transform(X.values)
    pls = PLSRegression(n_components=3)
    pls.fit(Xs, y.values)
    score = pls.score(Xs, y.values)
    log.info("PLS R² on training (survival vs genes) = %.3f", score)

    fig, ax = plt.subplots(figsize=(6, 5))
    ax.scatter(y.values, pls.predict(Xs).ravel(),
               alpha=0.7, color="teal")
    ax.set_xlabel("Observed days")
    ax.set_ylabel("Predicted days (PLS)")
    ax.set_title(f"PLS regression — {cohort}  (R²={score:.2f})")
    save_fig(fig, f"{cohort}_pls_survival", subfolder="survival")

    save_table(pd.DataFrame({"component": range(1, 4),
                             "x_scores_var": pls.x_scores_.var(axis=0),
                             "y_scores_var": pls.y_scores_.var(axis=0)}),
               f"pls_{cohort}")


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def main() -> None:
    log.info("=== SURVIVAL ANALYSIS ===")
    cohorts = config.TCGA_PROJECTS + [config.VALIDATION_COHORT]
    for c in cohorts:
        if (config.TABLES_DIR / f"labels_{c}.csv").exists():
            kaplan_meier(c)
            cox_regression(c)
            pls_regression(c)


if __name__ == "__main__":
    main()
