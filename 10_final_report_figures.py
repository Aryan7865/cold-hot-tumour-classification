"""
10_final_report_figures.py
--------------------------
OWNER: whole team  (run at the end)

Purpose
    Assemble a clean set of publication-ready composite figures + summary
    tables for the final term-paper PDF.  Designed to be run AFTER every
    other module has produced its outputs.

Produces
    outputs/figures/report/Fig1_pipeline.png          (schematic — optional, left
                                                       as placeholder)
    outputs/figures/report/Fig2_embedding_panel.png   (PCA + UMAP + t-SNE in one)
    outputs/figures/report/Fig3_survival_panel.png    (KM + Cox forest)
    outputs/figures/report/Fig4_model_comparison.png  (CV + external validation)
    outputs/tables/final_summary.csv
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import seaborn as sns

import config
from utils import get_logger, set_plot_style

log = get_logger("report")
set_plot_style()

REPORT_DIR = config.FIGURES_DIR / "report"
REPORT_DIR.mkdir(parents=True, exist_ok=True)


def _load_image(path: Path):
    if path.exists():
        return mpimg.imread(path)
    return None


def composite_embeddings(cohort: str = "TCGA-COAD") -> None:
    panels = ["PCA", "tsne", "UMAP"]
    fig, axes = plt.subplots(1, 3, figsize=(15, 5))
    for ax, method in zip(axes, panels):
        img = _load_image(config.FIGURES_DIR / "dim_reduction" /
                          f"{cohort}_{method.lower()}.png")
        if img is not None:
            ax.imshow(img)
            ax.axis("off")
            ax.set_title(method)
    fig.suptitle(f"Hot / Cold tumours separate in reduced space — {cohort}",
                 fontsize=14)
    fig.savefig(REPORT_DIR / "Fig2_embedding_panel.png",
                bbox_inches="tight", dpi=config.FIG_DPI)
    plt.close(fig)


def composite_survival(cohort: str = "TCGA-COAD") -> None:
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    km  = _load_image(config.FIGURES_DIR / "survival" /
                      f"{cohort}_kaplan_meier.png")
    pls = _load_image(config.FIGURES_DIR / "survival" /
                      f"{cohort}_pls_survival.png")
    for ax, img, title in zip(axes, (km, pls),
                               ("Kaplan-Meier", "PLS")):
        if img is not None:
            ax.imshow(img); ax.axis("off"); ax.set_title(title)
    fig.suptitle(f"Survival analysis — {cohort}", fontsize=14)
    fig.savefig(REPORT_DIR / "Fig3_survival_panel.png",
                bbox_inches="tight", dpi=config.FIG_DPI)
    plt.close(fig)


def composite_models() -> None:
    cvs, ext = [], []
    for c in config.TCGA_PROJECTS + [config.VALIDATION_COHORT]:
        cv_path = config.TABLES_DIR / f"cv_results_{c}.csv"
        if cv_path.exists():
            d = pd.read_csv(cv_path); d["cohort"] = c; cvs.append(d)
    cv = pd.concat(cvs) if cvs else pd.DataFrame()

    ext_path = (config.TABLES_DIR /
                f"external_validation_summary_{config.VALIDATION_COHORT}.csv")
    if ext_path.exists():
        ext = pd.read_csv(ext_path)
    else:
        ext = pd.DataFrame()

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    if len(cv):
        sns.barplot(data=cv, x="balanced_accuracy_mean", y="model",
                    hue="cohort", ax=axes[0])
        axes[0].set_xlim(0, 1)
        axes[0].set_title("CV balanced accuracy")
    if len(ext):
        sns.barplot(data=ext, x="balanced_accuracy", y="train",
                    ax=axes[1], color="coral")
        axes[1].set_xlim(0, 1)
        axes[1].set_title(f"External validation on {config.VALIDATION_COHORT}")
    fig.suptitle("Classifier performance", fontsize=14)
    fig.savefig(REPORT_DIR / "Fig4_model_comparison.png",
                bbox_inches="tight", dpi=config.FIG_DPI)
    plt.close(fig)


def final_summary_table() -> None:
    rows = []
    for c in config.TCGA_PROJECTS + [config.VALIDATION_COHORT]:
        lab = config.TABLES_DIR / f"labels_{c}.csv"
        if not lab.exists():
            continue
        df = pd.read_csv(lab)
        rows.append({
            "cohort": c,
            "n_samples": len(df),
            "Hot":  (df["label"] == "Hot").sum(),
            "Cold": (df["label"] == "Cold").sum(),
            "Intermediate": (df["label"] == "Intermediate").sum(),
        })
    pd.DataFrame(rows).to_csv(config.TABLES_DIR / "final_summary.csv", index=False)


def main() -> None:
    log.info("=== FINAL REPORT FIGURES ===")
    composite_embeddings()
    composite_survival()
    composite_models()
    final_summary_table()
    log.info("All report artefacts written to %s", REPORT_DIR)


if __name__ == "__main__":
    main()
