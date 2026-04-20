"""
config.py
---------
Central configuration for the Cold/Hot Tumour project.
Edit paths / thresholds here; every other script imports from this file so a
single change propagates through the whole pipeline.
"""

from __future__ import annotations
from pathlib import Path

# ------------------------------------------------------------------
# PROJECT ROOT
# ------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent

# ------------------------------------------------------------------
# DATA PATHS
# ------------------------------------------------------------------
DATA_DIR       = PROJECT_ROOT / "data"
RAW_DIR        = DATA_DIR / "raw"
PROCESSED_DIR  = DATA_DIR / "processed"
GENE_SETS_DIR  = PROJECT_ROOT / "gene_sets"

# ------------------------------------------------------------------
# OUTPUT PATHS
# ------------------------------------------------------------------
OUTPUTS_DIR    = PROJECT_ROOT / "outputs"
FIGURES_DIR    = OUTPUTS_DIR / "figures"
TABLES_DIR     = OUTPUTS_DIR / "tables"
MODELS_DIR     = OUTPUTS_DIR / "models"

for _p in (RAW_DIR, PROCESSED_DIR, GENE_SETS_DIR,
           FIGURES_DIR, TABLES_DIR, MODELS_DIR):
    _p.mkdir(parents=True, exist_ok=True)

# ------------------------------------------------------------------
# DATASET IDs
# ------------------------------------------------------------------
# Primary (training / discovery) cohort
TCGA_PROJECTS = ["TCGA-COAD", "TCGA-READ"]

# External GEO cohorts
GEO_VALIDATION = "GSE39582"   # 557 CRC samples  -> external validation
GEO_SURVIVAL   = "GSE17538"   # 224 CRC samples  -> survival follow-up

# ------------------------------------------------------------------
# PREPROCESSING
# ------------------------------------------------------------------
MIN_EXPR_CPM       = 1.0      # filter: gene expressed >=1 CPM in >= N% samples
MIN_PCT_SAMPLES    = 0.20     # 20 % samples
LOG_PSEUDOCOUNT    = 1.0      # log2(x + pseudocount)
VARIANCE_TOP_K     = 5000     # top-variance genes kept for downstream ML

# ------------------------------------------------------------------
# IMMUNE SCORING
# ------------------------------------------------------------------
# ESTIMATE gene sets (immune + stromal); downloaded by 03_immune_scoring.py
ESTIMATE_URL = "https://bioinformatics.mdanderson.org/estimate/data/SI_geneset.gmt"
IMMUNE_GMT_FALLBACK = GENE_SETS_DIR / "immune_signatures.gmt"

# Thresholds for Hot / Cold / Intermediate tumour labelling
# (based on quantiles of the immune score within the cohort)
HOT_QUANTILE      = 0.66      # top 1/3  -> Hot
COLD_QUANTILE     = 0.33      # bottom 1/3 -> Cold
# middle third    -> Intermediate

# ------------------------------------------------------------------
# ML + STATS
# ------------------------------------------------------------------
RANDOM_STATE     = 42
N_SPLITS_CV      = 5
PCA_N_COMPONENTS = 50
ALPHA            = 0.05       # significance threshold

# ------------------------------------------------------------------
# FIGURE DEFAULTS
# ------------------------------------------------------------------
FIG_DPI   = 150
FIG_STYLE = "whitegrid"        # seaborn style
PALETTE_LABEL = {
    "Hot":          "#d62728",  # red
    "Cold":         "#1f77b4",  # blue
    "Intermediate": "#7f7f7f",  # grey
}
