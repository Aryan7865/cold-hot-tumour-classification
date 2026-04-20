"""
04_dimensionality_reduction.py
------------------------------
OWNER: P3 (Exploratory Analyst)

Purpose
    Visualise the structure of the expression matrix using EVERY dimensionality
    reduction technique from BT3041:
        - PCA           (sklearn)
        - ICA           (FastICA)
        - MDS           (metric + non-metric)
        - t-SNE         (sklearn)
        - UMAP          (umap-learn)

    Overlay each embedding with the Hot / Cold / Intermediate label to show
    that immune phenotype partitions the transcriptome.

Inputs   : data/processed/<cohort>_expr_topvar.tsv
           outputs/tables/labels_<cohort>.csv
Outputs  : outputs/figures/dim_reduction/<cohort>_<method>.png
           outputs/tables/pca_variance_<cohort>.csv
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.decomposition import PCA, FastICA
from sklearn.manifold import MDS, TSNE
from sklearn.preprocessing import StandardScaler

import config
from utils import get_logger, save_fig, save_table, set_plot_style

log = get_logger("dimred")
set_plot_style()


# ------------------------------------------------------------------
# LOAD
# ------------------------------------------------------------------
def load_expr_and_labels(cohort: str) -> tuple[pd.DataFrame, pd.Series]:
    expr = pd.read_csv(config.PROCESSED_DIR / f"{cohort}_expr_topvar.tsv",
                       sep="\t", index_col=0)
    labels = pd.read_csv(config.TABLES_DIR / f"labels_{cohort}.csv",
                         index_col=0)["label"]
    common = expr.columns.intersection(labels.index)
    return expr[common], labels.loc[common]


# ------------------------------------------------------------------
# INDIVIDUAL METHODS
# ------------------------------------------------------------------
def run_pca(X: np.ndarray, n: int = config.PCA_N_COMPONENTS) -> tuple[PCA, np.ndarray]:
    pca = PCA(n_components=min(n, X.shape[0] - 1, X.shape[1]),
              random_state=config.RANDOM_STATE)
    Z = pca.fit_transform(X)
    return pca, Z


def run_ica(X: np.ndarray, n: int = 20) -> np.ndarray:
    ica = FastICA(n_components=min(n, X.shape[0] - 1, X.shape[1]),
                  random_state=config.RANDOM_STATE, max_iter=1000,
                  whiten="unit-variance")
    return ica.fit_transform(X)


def run_mds(X: np.ndarray, n: int = 2, metric: bool = True) -> np.ndarray:
    mds = MDS(n_components=n, metric=metric, n_init=2,
              random_state=config.RANDOM_STATE, normalized_stress="auto")
    return mds.fit_transform(X)


def run_tsne(X: np.ndarray, perplexity: float = 30) -> np.ndarray:
    perp = min(perplexity, (X.shape[0] - 1) / 3)
    t = TSNE(n_components=2, perplexity=perp,
             random_state=config.RANDOM_STATE, init="pca",
             learning_rate="auto")
    return t.fit_transform(X)


def run_umap(X: np.ndarray) -> np.ndarray:
    try:
        import umap
    except ImportError as e:
        raise RuntimeError("pip install umap-learn") from e
    reducer = umap.UMAP(n_components=2, random_state=config.RANDOM_STATE)
    return reducer.fit_transform(X)


# ------------------------------------------------------------------
# PLOT HELPER
# ------------------------------------------------------------------
def scatter_embedding(Z: np.ndarray, labels: pd.Series,
                      method: str, cohort: str,
                      explained: str | None = None) -> None:
    fig, ax = plt.subplots(figsize=(7, 6))
    df = pd.DataFrame({"x": Z[:, 0], "y": Z[:, 1], "label": labels.values})
    sns.scatterplot(data=df, x="x", y="y", hue="label",
                    palette=config.PALETTE_LABEL, s=40, alpha=0.85, ax=ax)
    title = f"{method} — {cohort}"
    if explained:
        title += f"  ({explained})"
    ax.set_title(title)
    ax.set_xlabel(f"{method} 1")
    ax.set_ylabel(f"{method} 2")
    ax.legend(title="Tumour phenotype", frameon=False)
    save_fig(fig, f"{cohort}_{method.lower().replace('-', '')}",
             subfolder="dim_reduction")


def pca_variance_plot(pca: PCA, cohort: str) -> None:
    var = pca.explained_variance_ratio_
    cum = np.cumsum(var)
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar(range(1, len(var) + 1), var, alpha=0.6, label="individual")
    ax.plot(range(1, len(var) + 1), cum, "-o", color="black", label="cumulative")
    ax.axhline(0.95, color="red", ls="--", label="95 %")
    ax.set_xlabel("Principal component")
    ax.set_ylabel("Variance explained")
    ax.set_title(f"PCA variance — {cohort}")
    ax.legend(frameon=False)
    save_fig(fig, f"{cohort}_pca_variance", subfolder="dim_reduction")

    tbl = pd.DataFrame({"PC": np.arange(1, len(var) + 1),
                        "variance_ratio": var,
                        "cumulative": cum})
    save_table(tbl, f"pca_variance_{cohort}")


# ------------------------------------------------------------------
# ENTRY POINT
# ------------------------------------------------------------------
def process_cohort(cohort: str) -> None:
    log.info("=== %s ===", cohort)
    expr, labels = load_expr_and_labels(cohort)
    X = StandardScaler().fit_transform(expr.T.values)   # samples x genes

    # 1. PCA
    pca, Zp = run_pca(X)
    pca_variance_plot(pca, cohort)
    scatter_embedding(Zp[:, :2], labels, "PCA", cohort,
                      explained=f"PC1 {pca.explained_variance_ratio_[0]:.1%}, "
                                f"PC2 {pca.explained_variance_ratio_[1]:.1%}")

    # 2. ICA
    Zi = run_ica(X)
    scatter_embedding(Zi[:, :2], labels, "ICA", cohort)

    # 3. MDS (metric + non-metric)
    scatter_embedding(run_mds(X, metric=True),   labels, "MDS-metric",    cohort)
    scatter_embedding(run_mds(X, metric=False),  labels, "MDS-nonmetric", cohort)

    # 4. t-SNE
    scatter_embedding(run_tsne(Zp[:, :20]), labels, "t-SNE", cohort)   # t-SNE on top-20 PCs

    # 5. UMAP
    scatter_embedding(run_umap(Zp[:, :20]), labels, "UMAP", cohort)


def main() -> None:
    log.info("=== DIMENSIONALITY REDUCTION ===")
    cohorts = config.TCGA_PROJECTS + [config.VALIDATION_COHORT]
    for c in cohorts:
        if (config.TABLES_DIR / f"labels_{c}.csv").exists():
            process_cohort(c)


if __name__ == "__main__":
    main()
