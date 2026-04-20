# Cold vs Hot Tumour Classification from Bulk RNA-seq
**BT3041 Analysis and Interpretation of Biological Project Data — Term Paper**

## Project Title
**A Multi-Statistical Approach to Classifying Cold vs. Hot Tumours from Bulk RNA-seq**

## Abstract (draft)
Immune "hot" and "cold" tumours show drastically different responses to immunotherapy, yet robust, reproducible, bulk-RNA-seq-based classifiers remain scarce. Using ~1,400 colorectal cancer (CRC) samples pooled from TCGA-COAD, TCGA-READ, GSE39582, and GSE17538, we derive patient-level immune phenotypes via ssGSEA and ESTIMATE, probe the resulting groups with the full spectrum of dimensionality-reduction techniques taught in BT3041 (PCA, ICA, MDS, t-SNE, UMAP), validate them with unsupervised clustering (k-means, hierarchical), identify immune-defining genes via t-tests / ANOVA under strict FDR control, train SVM / kNN / Logistic-Regression classifiers with cross-validation, link tumour phenotype to clinical outcome via Kaplan–Meier and Cox regression, and finally validate everything on an independent external cohort (GSE39582). The pipeline is engineered to exercise every statistical and ML concept taught in the course.

## Datasets (all public, ~1,400 patients)
| Dataset | Type | Samples | Source |
|--------|------|---------|--------|
| TCGA-COAD | RNA-seq (FPKM / HTSeq-counts) | 456 | GDC portal |
| TCGA-READ | RNA-seq | 163 | GDC portal |
| GSE39582 | Microarray (Affymetrix HG-U133 Plus 2.0) | 557 | GEO |
| GSE17538 | Microarray + survival | 224 | GEO |

## Pipeline Overview
```
raw data  →  preprocess  →  immune scoring (ssGSEA + ESTIMATE)
                                   ↓
                       Hot / Cold / Intermediate labels
                                   ↓
        ┌────────────────────┬──────────────┬──────────────┐
        ↓                    ↓              ↓              ↓
 dim reduction         clustering    hypothesis     classification
 (PCA/ICA/MDS/         (k-means,      testing        (SVM/kNN/LR)
  t-SNE/UMAP)          hierarchical)  (t-test/                ↓
                                       ANOVA/FDR/      survival (KM/Cox)
                                       chi-sq/PLS)            ↓
                                                      external validation
                                                      (GSE39582)
```

## Team Assignments
| Member | Modules | Files |
|--------|---------|-------|
| P1 (Data eng.) | Acquisition + Preprocessing | `01_data_acquisition.py`, `02_preprocessing.py` |
| P2 (Immune phenotyper) | Immune scoring & labelling | `03_immune_scoring.py` |
| P3 (Explorer) | Dim reduction + Clustering | `04_dimensionality_reduction.py`, `05_clustering.py` |
| P4 (Statistician) | Hypothesis testing | `06_hypothesis_testing.py` |
| P5 (ML) | Classification | `07_classification.py` |
| P6 (Survival + validation) | Survival + external cohort | `08_survival_analysis.py`, `09_validation_external.py` |
| ALL | Final figures & report | `10_final_report_figures.py` |

## How to Run
```bash
# 1. Create env
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# 2. Download data (≈ 30 min, 2–3 GB)
python 01_data_acquisition.py

# 3. Preprocess
python 02_preprocessing.py

# 4. Score immune phenotype and assign Hot / Cold labels
python 03_immune_scoring.py

# 5. Dim reduction + clustering (produces figures)
python 04_dimensionality_reduction.py
python 05_clustering.py

# 6. Statistics
python 06_hypothesis_testing.py

# 7. ML
python 07_classification.py

# 8. Survival
python 08_survival_analysis.py

# 9. External validation
python 09_validation_external.py

# 10. Generate combined report figures
python 10_final_report_figures.py
```

## Folder Layout
```
cold_hot_tumour_project/
├── README.md
├── requirements.txt
├── config.py
├── utils.py
├── 01_data_acquisition.py
├── 02_preprocessing.py
├── 03_immune_scoring.py
├── 04_dimensionality_reduction.py
├── 05_clustering.py
├── 06_hypothesis_testing.py
├── 07_classification.py
├── 08_survival_analysis.py
├── 09_validation_external.py
├── 10_final_report_figures.py
├── data/
│   ├── raw/            # untouched downloads
│   └── processed/      # normalised expression matrices
├── gene_sets/          # ssGSEA / ESTIMATE signatures
└── outputs/
    ├── figures/
    ├── tables/
    └── models/
```

## Course-concept coverage checklist
- [x] Standard normal distribution / z-score (preprocessing)
- [x] Hypothesis testing (t-test, ANOVA, chi-squared, F-test)
- [x] Multiple testing correction (Bonferroni + Benjamini–Hochberg)
- [x] Regression (linear, logistic, Cox, PLS)
- [x] Nearest-neighbour classifier (kNN)
- [x] Feature selection (variance filter + DE-based)
- [x] Hierarchical + k-means clustering
- [x] Support Vector Machines
- [x] PCA, ICA, MDS
- [x] t-SNE, UMAP
- [x] Partial Least Squares (PLS)
