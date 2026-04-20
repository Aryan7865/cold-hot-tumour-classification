"""
09_validation_external.py
-------------------------
OWNER: P6 (Survival + Validation)

Purpose
    Test whether the classifier trained on TCGA (COAD+READ) generalises to an
    entirely different technology / cohort  (GSE39582 microarray).
    This is the single biggest differentiator vs the previous-year project —
    independent external validation is what reviewers look for.

Pipeline
    1. Load TCGA model (from step 07) and its gene-feature set.
    2. Load GSE39582 log-expression matrix.
    3. Intersect features (gene symbols).  Zero-impute missing.
    4. Predict Hot / Cold on GSE39582.
    5. Compare predictions to GSE39582's own ssGSEA-based labels
       (computed by 03_immune_scoring.py when run on that cohort).
    6. Report accuracy / F1, plus confusion heatmap.

Outputs
    outputs/tables/external_validation_metrics.csv
    outputs/figures/validation/*.png
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import (classification_report, confusion_matrix,
                             balanced_accuracy_score, f1_score)

import config
from utils import get_logger, save_fig, save_table, load_pickle, set_plot_style

log = get_logger("validation")
set_plot_style()


# ------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------
def load_matching_matrix(train_cohort: str,
                         test_cohort: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    train = pd.read_csv(config.PROCESSED_DIR / f"{train_cohort}_expr_topvar.tsv",
                        sep="\t", index_col=0)
    test  = pd.read_csv(config.PROCESSED_DIR / f"{test_cohort}_expr_log.tsv",
                        sep="\t", index_col=0)

    # Intersect on gene symbols; for microarray, probe IDs may differ, so run
    # preprocessing ensuring gene symbols are used.
    common = train.index.intersection(test.index)
    log.info("Common genes between %s and %s: %d",
             train_cohort, test_cohort, len(common))

    test_aligned = test.loc[common]
    # Add zero-rows for any training features missing in test
    missing = train.index.difference(common)
    if len(missing):
        log.warning("%d training features missing in test -> imputing 0",
                    len(missing))
        zeros = pd.DataFrame(0, index=missing, columns=test_aligned.columns)
        test_aligned = pd.concat([test_aligned, zeros]).loc[train.index]
    else:
        test_aligned = test_aligned.reindex(train.index)

    return train, test_aligned


def best_model_name(train_cohort: str) -> str:
    """Read step-07 cv results and return the top model name."""
    cv = pd.read_csv(config.TABLES_DIR / f"cv_results_{train_cohort}.csv")
    return cv.sort_values("balanced_accuracy_mean",
                          ascending=False).iloc[0]["model"]


# ------------------------------------------------------------------
# VALIDATION
# ------------------------------------------------------------------
def validate(train_cohort: str = "TCGA-COAD",
             test_cohort:  str = config.GEO_VALIDATION) -> pd.DataFrame:
    log.info("Validating %s -> %s", train_cohort, test_cohort)

    best = best_model_name(train_cohort)
    log.info("Best model on %s: %s", train_cohort, best)
    pipe = load_pickle(f"{best}_{train_cohort}")

    _, X_test = load_matching_matrix(train_cohort, test_cohort)

    # Model expects samples x genes in the training feature order
    y_pred = pipe.predict(X_test.T.values)

    # Compare to GSE39582's own ssGSEA-based labels if available
    lab_path = config.TABLES_DIR / f"labels_{test_cohort}.csv"
    if not lab_path.exists():
        log.warning("No ssGSEA labels for %s (run step 03 on it).", test_cohort)
        pd.DataFrame({"sample": X_test.columns, "pred": y_pred}) \
          .to_csv(config.TABLES_DIR /
                  f"external_predictions_{test_cohort}.csv", index=False)
        return pd.DataFrame()

    labels = pd.read_csv(lab_path, index_col=0)["label"]
    common = X_test.columns.intersection(labels.index)
    y_true = labels.loc[common]
    y_pred = pd.Series(y_pred, index=X_test.columns).loc[common]

    acc  = (y_true == y_pred).mean()
    bac  = balanced_accuracy_score(y_true, y_pred)
    f1m  = f1_score(y_true, y_pred, average="macro")
    log.info("acc=%.3f  balanced-acc=%.3f  f1=%.3f", acc, bac, f1m)

    # Confusion matrix
    classes = sorted(y_true.unique())
    cm = confusion_matrix(y_true, y_pred, labels=classes)
    fig, ax = plt.subplots(figsize=(5, 4))
    sns.heatmap(cm, annot=True, fmt="d", cmap="Purples",
                xticklabels=classes, yticklabels=classes, ax=ax)
    ax.set_xlabel("Predicted"); ax.set_ylabel("True (ssGSEA)")
    ax.set_title(f"External validation — {test_cohort}\n"
                 f"model trained on {train_cohort}")
    save_fig(fig, f"validation_confmat_{test_cohort}",
             subfolder="validation")

    report = classification_report(y_true, y_pred, output_dict=True)
    tbl = pd.DataFrame(report).T
    tbl["model"]        = best
    tbl["train_cohort"] = train_cohort
    tbl["test_cohort"]  = test_cohort
    save_table(tbl, f"external_validation_{test_cohort}")

    # Summary row
    summary = pd.DataFrame([{
        "train": train_cohort,  "test": test_cohort,
        "model": best, "accuracy": acc,
        "balanced_accuracy": bac, "f1_macro": f1m,
        "n": len(y_true),
    }])
    save_table(summary, f"external_validation_summary_{test_cohort}")
    return summary


def main() -> None:
    log.info("=== EXTERNAL VALIDATION ===")
    for train in config.TCGA_PROJECTS:
        if (config.MODELS_DIR).glob(f"*_{train}.pkl"):
            try:
                validate(train_cohort=train)
            except FileNotFoundError as e:
                log.warning("Skipping %s -> %s", train, e)


if __name__ == "__main__":
    main()
