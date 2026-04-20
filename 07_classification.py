"""
07_classification.py
--------------------
OWNER: P5 (Machine Learning)

Purpose
    Build classifiers that predict Hot / Cold / (Intermediate) from a gene
    expression profile, using every classifier taught in BT3041:
        - kNN         (nearest-neighbour)
        - SVM         (linear + RBF)
        - Logistic    (multinomial)
        - Random Forest   (bonus, tree baseline)
        - XGBoost         (bonus)

Pipeline
        StandardScaler  ->  Feature selection  ->  Model  ->  5-fold CV
        Report: accuracy, balanced accuracy, macro-F1, ROC-AUC (OvR),
        confusion matrix, per-class precision / recall.

Feature selection strategies (course concept: feature selection)
    1. Variance filter (top-K variable genes) — already done in step 02
    2. ANOVA-based filter (SelectKBest with f_classif)
    3. Signature panel (top-DE genes from step 06)

Outputs
    outputs/tables/cv_results_<cohort>.csv
    outputs/models/<model>_<cohort>.pkl
    outputs/figures/classification/*.png
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.feature_selection import SelectKBest, f_classif
from sklearn.linear_model import LogisticRegression
from sklearn.neighbors import KNeighborsClassifier
from sklearn.svm import SVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import StandardScaler, label_binarize
from sklearn.pipeline import Pipeline
from sklearn.model_selection import StratifiedKFold, cross_val_score, cross_val_predict
from sklearn.metrics import (classification_report, confusion_matrix,
                             balanced_accuracy_score, f1_score, roc_auc_score,
                             roc_curve)

import config
from utils import (get_logger, save_fig, save_table, save_pickle,
                   set_plot_style)

log = get_logger("classification")
set_plot_style()


LABEL_ORDER = ["Cold", "Intermediate", "Hot"]


# ------------------------------------------------------------------
# LOAD
# ------------------------------------------------------------------
def load_xy(cohort: str, exclude_intermediate: bool = False
            ) -> tuple[pd.DataFrame, pd.Series]:
    expr = pd.read_csv(config.PROCESSED_DIR / f"{cohort}_expr_topvar.tsv",
                       sep="\t", index_col=0)
    labels = pd.read_csv(config.TABLES_DIR / f"labels_{cohort}.csv",
                         index_col=0)["label"]
    common = expr.columns.intersection(labels.index)
    X = expr[common].T        # samples x genes
    y = labels.loc[common]
    if exclude_intermediate:
        keep = y != "Intermediate"
        X, y = X[keep], y[keep]
    return X, y


# ------------------------------------------------------------------
# MODEL ZOO
# ------------------------------------------------------------------
def build_models(n_features: int) -> dict[str, Pipeline]:
    k = min(500, n_features)
    models = {
        "kNN-5": Pipeline([
            ("sc",  StandardScaler()),
            ("fs",  SelectKBest(f_classif, k=k)),
            ("clf", KNeighborsClassifier(n_neighbors=5)),
        ]),
        "SVM-linear": Pipeline([
            ("sc",  StandardScaler()),
            ("fs",  SelectKBest(f_classif, k=k)),
            ("clf", SVC(kernel="linear", C=1.0, probability=True,
                        random_state=config.RANDOM_STATE)),
        ]),
        "SVM-RBF": Pipeline([
            ("sc",  StandardScaler()),
            ("fs",  SelectKBest(f_classif, k=k)),
            ("clf", SVC(kernel="rbf", C=1.0, gamma="scale", probability=True,
                        random_state=config.RANDOM_STATE)),
        ]),
        "Logistic": Pipeline([
            ("sc",  StandardScaler()),
            ("fs",  SelectKBest(f_classif, k=k)),
            ("clf", LogisticRegression(max_iter=2000, C=1.0,
                                       multi_class="auto",
                                       random_state=config.RANDOM_STATE)),
        ]),
        "RandomForest": Pipeline([
            ("sc",  StandardScaler(with_mean=False)),
            ("clf", RandomForestClassifier(n_estimators=500,
                                           random_state=config.RANDOM_STATE,
                                           n_jobs=-1)),
        ]),
    }
    try:
        from xgboost import XGBClassifier
        models["XGBoost"] = Pipeline([
            ("sc",  StandardScaler(with_mean=False)),
            ("clf", XGBClassifier(n_estimators=500,
                                  random_state=config.RANDOM_STATE,
                                  eval_metric="mlogloss",
                                  use_label_encoder=False,
                                  n_jobs=-1)),
        ])
    except ImportError:
        log.info("xgboost not installed; skipping.")
    return models


# ------------------------------------------------------------------
# EVALUATION
# ------------------------------------------------------------------
def cv_score(models: dict[str, Pipeline], X: pd.DataFrame, y: pd.Series
             ) -> pd.DataFrame:
    cv = StratifiedKFold(n_splits=config.N_SPLITS_CV, shuffle=True,
                         random_state=config.RANDOM_STATE)
    rows = []
    for name, pipe in models.items():
        log.info("CV -> %s", name)
        acc = cross_val_score(pipe, X, y, cv=cv, scoring="accuracy", n_jobs=-1)
        bac = cross_val_score(pipe, X, y, cv=cv, scoring="balanced_accuracy",
                              n_jobs=-1)
        f1  = cross_val_score(pipe, X, y, cv=cv, scoring="f1_macro", n_jobs=-1)
        rows.append({"model": name,
                     "accuracy_mean":           acc.mean(),
                     "accuracy_std":            acc.std(),
                     "balanced_accuracy_mean":  bac.mean(),
                     "f1_macro_mean":           f1.mean()})
    return pd.DataFrame(rows).sort_values("balanced_accuracy_mean",
                                          ascending=False)


def confusion_and_roc(pipe: Pipeline, X: pd.DataFrame, y: pd.Series,
                      name: str, cohort: str) -> None:
    cv = StratifiedKFold(n_splits=config.N_SPLITS_CV, shuffle=True,
                         random_state=config.RANDOM_STATE)
    y_pred  = cross_val_predict(pipe, X, y, cv=cv, n_jobs=-1)
    classes = sorted(y.unique())
    cm = confusion_matrix(y, y_pred, labels=classes)

    fig, ax = plt.subplots(figsize=(5, 4))
    sns.heatmap(cm, annot=True, fmt="d", cmap="Blues",
                xticklabels=classes, yticklabels=classes, ax=ax)
    ax.set_xlabel("Predicted"); ax.set_ylabel("True")
    ax.set_title(f"{name} — confusion ({cohort})")
    save_fig(fig, f"{cohort}_{name}_confmat", subfolder="classification")

    try:
        y_proba = cross_val_predict(pipe, X, y, cv=cv, method="predict_proba",
                                    n_jobs=-1)
        Y = label_binarize(y, classes=classes)
        fig, ax = plt.subplots(figsize=(6, 5))
        for i, c in enumerate(classes):
            fpr, tpr, _ = roc_curve(Y[:, i], y_proba[:, i])
            auc = roc_auc_score(Y[:, i], y_proba[:, i])
            ax.plot(fpr, tpr, label=f"{c}  (AUC={auc:.2f})")
        ax.plot([0, 1], [0, 1], ls="--", color="grey")
        ax.set_xlabel("FPR"); ax.set_ylabel("TPR")
        ax.set_title(f"{name} — ROC OvR ({cohort})")
        ax.legend(frameon=False)
        save_fig(fig, f"{cohort}_{name}_roc", subfolder="classification")
    except Exception as e:          # noqa: BLE001
        log.warning("ROC skipped for %s: %s", name, e)


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def process_cohort(cohort: str) -> None:
    log.info("=== %s ===", cohort)
    X, y = load_xy(cohort)
    log.info("X %s  |  y distribution: %s", X.shape, y.value_counts().to_dict())

    models = build_models(n_features=X.shape[1])
    results = cv_score(models, X, y)
    save_table(results, f"cv_results_{cohort}")

    # Fit + save best model, confusion matrix / ROC for top 3
    top_models = results.head(3)["model"].tolist()
    for name in top_models:
        pipe = models[name].fit(X, y)
        save_pickle(pipe, f"{name}_{cohort}")
        confusion_and_roc(models[name], X, y, name, cohort)

    # Plot a bar chart summary
    fig, ax = plt.subplots(figsize=(7, 4))
    sns.barplot(data=results, y="model",
                x="balanced_accuracy_mean", palette="viridis", ax=ax)
    ax.set_xlim(0, 1)
    ax.set_xlabel("Balanced accuracy (5-fold CV)")
    ax.set_title(f"Model comparison — {cohort}")
    save_fig(fig, f"{cohort}_model_comparison", subfolder="classification")


def main() -> None:
    log.info("=== CLASSIFICATION ===")
    cohorts = config.TCGA_PROJECTS + [config.VALIDATION_COHORT]
    for c in cohorts:
        if (config.TABLES_DIR / f"labels_{c}.csv").exists():
            process_cohort(c)


if __name__ == "__main__":
    main()
