"""
05_clustering.py
----------------
OWNER: P3 (Exploratory Analyst)

Purpose
    Unsupervised class-discovery:  do Hot / Cold tumours fall out naturally
    from the expression matrix?

Methods
    - k-means                  (k = 2, 3, 4)
    - hierarchical (Ward)      + dendrogram
    - (optional) DBSCAN

Metrics
    - Silhouette score for k selection
    - Adjusted Rand Index (ARI) vs the Hot/Cold label
    - Adjusted Mutual Information (AMI)
    - Contingency table (Chi-squared can be run in 06)

Inputs   : data/processed/<cohort>_expr_topvar.tsv
           outputs/tables/labels_<cohort>.csv
Outputs  : outputs/figures/clustering/*.png
           outputs/tables/clustering_metrics_<cohort>.csv
           outputs/tables/clusters_<cohort>.csv        <- per-sample cluster ID
"""

from __future__ import annotations
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from scipy.cluster.hierarchy import linkage, dendrogram, fcluster
from scipy.spatial.distance import pdist
from sklearn.cluster import KMeans, AgglomerativeClustering, DBSCAN
from sklearn.metrics import (silhouette_score, silhouette_samples,
                             adjusted_rand_score, adjusted_mutual_info_score)
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA

import config
from utils import get_logger, save_fig, save_table, set_plot_style

log = get_logger("clustering")
set_plot_style()


def load_expr_and_labels(cohort: str) -> tuple[pd.DataFrame, pd.Series]:
    expr = pd.read_csv(config.PROCESSED_DIR / f"{cohort}_expr_topvar.tsv",
                       sep="\t", index_col=0)
    labels = pd.read_csv(config.TABLES_DIR / f"labels_{cohort}.csv",
                         index_col=0)["label"]
    common = expr.columns.intersection(labels.index)
    return expr[common], labels.loc[common]


# ------------------------------------------------------------------
# K selection
# ------------------------------------------------------------------
def silhouette_sweep(X: np.ndarray, k_range: range, cohort: str) -> pd.DataFrame:
    rows = []
    fig, ax = plt.subplots(figsize=(8, 4))
    for k in k_range:
        km = KMeans(n_clusters=k, random_state=config.RANDOM_STATE, n_init="auto")
        lab = km.fit_predict(X)
        s = silhouette_score(X, lab) if k > 1 else np.nan
        rows.append({"k": k, "inertia": km.inertia_, "silhouette": s})
    tbl = pd.DataFrame(rows)
    save_table(tbl, f"kmeans_sweep_{cohort}")

    ax.plot(tbl["k"], tbl["silhouette"], "-o", color="teal")
    ax.set_xlabel("k"); ax.set_ylabel("Silhouette")
    ax.set_title(f"Silhouette sweep — {cohort}")
    save_fig(fig, f"{cohort}_silhouette_sweep", subfolder="clustering")
    return tbl


# ------------------------------------------------------------------
# ALGORITHMS
# ------------------------------------------------------------------
def run_kmeans(X: np.ndarray, k: int) -> np.ndarray:
    return KMeans(n_clusters=k, random_state=config.RANDOM_STATE,
                  n_init="auto").fit_predict(X)


def run_hierarchical(X: np.ndarray, k: int,
                     linkage_method: str = "ward") -> tuple[np.ndarray, np.ndarray]:
    Z = linkage(X, method=linkage_method)
    lab = fcluster(Z, t=k, criterion="maxclust") - 1
    return lab, Z


def run_dbscan(X: np.ndarray, eps: float = 3.0, min_samples: int = 5) -> np.ndarray:
    return DBSCAN(eps=eps, min_samples=min_samples).fit_predict(X)


# ------------------------------------------------------------------
# PLOTS
# ------------------------------------------------------------------
def plot_dendrogram(Z: np.ndarray, cohort: str,
                    labels: pd.Series | None = None) -> None:
    fig, ax = plt.subplots(figsize=(12, 5))
    dendrogram(Z, no_labels=True, color_threshold=None,
               above_threshold_color="grey", ax=ax)
    ax.set_title(f"Hierarchical (Ward) dendrogram — {cohort}")
    ax.set_xlabel("samples")
    save_fig(fig, f"{cohort}_dendrogram", subfolder="clustering")


def plot_clusters_on_pca(X: np.ndarray, cluster_ids: np.ndarray,
                         method: str, cohort: str) -> None:
    pcs = PCA(n_components=2, random_state=config.RANDOM_STATE).fit_transform(X)
    fig, ax = plt.subplots(figsize=(7, 6))
    palette = sns.color_palette("tab10", n_colors=len(np.unique(cluster_ids)))
    for i, c in enumerate(np.unique(cluster_ids)):
        m = cluster_ids == c
        ax.scatter(pcs[m, 0], pcs[m, 1], color=palette[i], label=f"cluster {c}",
                   alpha=0.8, edgecolor="white")
    ax.set_title(f"{method} clusters — {cohort}")
    ax.set_xlabel("PC1"); ax.set_ylabel("PC2")
    ax.legend(frameon=False)
    save_fig(fig, f"{cohort}_{method}_on_pca", subfolder="clustering")


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def process_cohort(cohort: str) -> None:
    log.info("=== %s ===", cohort)
    expr, labels = load_expr_and_labels(cohort)
    X = StandardScaler().fit_transform(expr.T.values)

    # 1. k selection
    silhouette_sweep(X, range(2, 7), cohort)

    # 2. k-means (k=2 since our target is Hot vs Cold; also k=3 for 3-label)
    results = {}
    for k in (2, 3):
        km_lab = run_kmeans(X, k)
        hc_lab, Z = run_hierarchical(X, k)
        results[f"kmeans_k{k}"] = km_lab
        results[f"hier_k{k}"]   = hc_lab
        if k == 3:
            plot_dendrogram(Z, cohort)

    # 3. DBSCAN (exploratory, k not pre-chosen)
    results["dbscan"] = run_dbscan(X)

    # 4. Compare to Hot/Cold ground-truth label
    y_true = labels.map({"Cold": 0, "Intermediate": 1, "Hot": 2}).fillna(-1).astype(int)
    rows = []
    for name, lab in results.items():
        mask = y_true >= 0
        rows.append({
            "method": name,
            "n_clusters": len(np.unique(lab[lab >= 0])),
            "ARI":  adjusted_rand_score(y_true[mask], lab[mask]),
            "AMI":  adjusted_mutual_info_score(y_true[mask], lab[mask]),
            "silhouette": silhouette_score(X, lab) if len(np.unique(lab)) > 1 else np.nan,
        })
        plot_clusters_on_pca(X, lab, name, cohort)
    tbl = pd.DataFrame(rows).sort_values("ARI", ascending=False)
    save_table(tbl, f"clustering_metrics_{cohort}")

    # 5. Save cluster ids per sample
    cluster_df = pd.DataFrame(results, index=labels.index)
    cluster_df["label"] = labels
    save_table(cluster_df, f"clusters_{cohort}")


def main() -> None:
    log.info("=== CLUSTERING ===")
    cohorts = config.TCGA_PROJECTS + [config.GEO_VALIDATION]
    for c in cohorts:
        if (config.TABLES_DIR / f"labels_{c}.csv").exists():
            process_cohort(c)


if __name__ == "__main__":
    main()
