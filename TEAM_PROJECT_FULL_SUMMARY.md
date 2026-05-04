# BT3041 Cold vs Hot Tumour Project — Full Team Briefing

**Purpose of this document:** Everyone on the team should read this before the presentation or viva. It explains the **main idea**, **what we actually built and ran**, **how train vs validation differs**, **every pipeline step**, **key numbers**, **limitations**, and **how to talk about it confidently**.

**Repository:** [https://github.com/Aryan7865/cold-hot-tumour-classification](https://github.com/Aryan7865/cold-hot-tumour-classification)

**Generated PDF report (figures + tables):** `report/BT3041_Cold_Hot_Tumour_Report.pdf` — regenerate with `python report/generate_report_pdf.py` after activating the venv.

---

## 1. Main idea (one minute version)

**Biology:** Some tumours are **immune-hot** — heavily infiltrated by immune cells and more likely to respond to immunotherapy (e.g. checkpoint inhibitors). Others are **immune-cold** — poorly infiltrated and harder to treat that way. There is also an **intermediate** group.

**Engineering / stats question:** Can we infer this **Hot / Cold / Intermediate** phenotype from **bulk RNA-seq** (gene expression counts thousands of genes measure at once) using the **full statistical and ML toolkit from BT3041** — not just one classifier, but preprocessing, distributions, hypothesis tests with multiple-testing correction, dimensionality reduction, clustering, supervised learning, survival analysis, and **external validation on data the model never saw during training**?

**Answer we give in the project:** Yes, with strong evidence on discovery data and **good generalisation** to a held-out TCGA cohort (different patients, same technology).

---

## 2. Original plan vs what we shipped (important for honesty)

| Aspect | Early plan (README draft) | **What we actually executed** |
|--------|---------------------------|--------------------------------|
| Cohorts | TCGA-COAD + READ + GEO (GSE39582, GSE17538), ~1,400 samples | **TCGA-COAD + TCGA-READ only** — GEO dropped due to download instability and scope |
| Discovery / training | Mixed | **TCGA-COAD only** for model training and all discovery analyses |
| External test | GSE39582 microarray | **TCGA-READ** — independent patients, **same RNA-seq pipeline** as COAD (cleaner than mixing microarray + RNA-seq) |
| Sample counts | Various | **294 COAD + 109 READ = 403 tumour samples** in final outputs |

If a professor asks *“Why only TCGA?”* — **Same platform, two independent colorectal cohorts, no cross-platform normalisation mess, and still real external validation (READ never used in training).**

---

## 3. Labels: how we define Hot / Cold / Intermediate

We do **not** use pathology slides as gold standard. We use **computational immune scoring**:

1. **ssGSEA** (single-sample Gene Set Enrichment Analysis) on curated immune gene sets + **ESTIMATE** immune/stromal signatures (`03_immune_scoring.py`).
2. Rank samples by immune-related score (ESTIMATE immune score is primary for splitting).
3. **Tertile split within each cohort:** top third → **Hot**, bottom third → **Cold**, middle third → **Intermediate** (`config.py`: `HOT_QUANTILE` 0.66, `COLD_QUANTILE` 0.33).

**Final label counts** (from `outputs/tables/final_summary.csv`):

| Cohort | n samples | Hot | Cold | Intermediate |
|--------|-----------|-----|------|----------------|
| TCGA-COAD | 294 | 100 | 97 | 97 |
| TCGA-READ | 109 | 37 | 36 | 36 |

So classes are roughly balanced — good for classification metrics.

**Presentation tip:** Say clearly: *“Labels are ssGSEA-derived immune phenotypes; we then ask whether the rest of the transcriptome and ML recover this partition and whether it links to survival.”*

---

## 4. Train vs validation — this is what you must not confuse

| Role | Cohort | What happens there |
|------|--------|----------------------|
| **Discovery / training** | **TCGA-COAD** (294 samples) | Preprocessing, immune scoring, all plots/tables for “main results”, **5-fold cross-validation** for every classifier, best model picked by **mean balanced accuracy** on COAD |
| **External validation (held-out)** | **TCGA-READ** (109 samples) | **Never used in training.** Same genes aligned to training feature space; missing genes zero-filled. We apply the **frozen** best COAD model and compare predictions to READ’s own ssGSEA labels (`09_validation_external.py` + `config.VALIDATION_COHORT`) |

**This is real external validation** in the sense of **different patients and different tumour collection** — it is **not** random 80/20 split on one dataset (we also have CV on COAD, but READ is the strong story).

**Numbers (READ external test, best model = RandomForest):**  
From `outputs/tables/external_validation_summary_TCGA-READ.csv`:

- **Accuracy ≈ 0.743**
- **Balanced accuracy ≈ 0.744**
- **Macro F1 ≈ 0.742**
- **n = 109**

For **3 classes**, random guessing ≈ **33%** accuracy — so **~74% is clearly meaningful**, not luck. Errors are often **Hot ↔ Intermediate**, which is biologically believable (continuous immune infiltration, not perfect boxes).

---

## 5. End-to-end pipeline (scripts 01 → 10)

Run order (after `pip install -r requirements.txt` and venv):

### `01_data_acquisition.py` — Data engineering

- Pulls **TCGA** STAR-count expression + clinical metadata via **GDC REST API**.
- **Important fix we implemented:** bulk download in **batches** (e.g. 50 files) with **retries**, because one giant tar stream often **truncated** on slow/unstable connections.
- Builds per-cohort expression matrices: `data/raw/<cohort>/expression_counts.tsv` (genes × samples).

### `02_preprocessing.py` — Uniform tables for all downstream steps

- **CPM** normalisation, filter low expression (`MIN_EXPR_CPM`, `MIN_PCT_SAMPLES` in `config.py`).
- **log₂(count + 1)** transform.
- Keeps **top 5,000 highest-variance genes** for ML-heavy steps (`VARIANCE_TOP_K`).
- Outputs: `data/processed/*_expr_log.tsv`, `*_expr_topvar.tsv`, etc.

### `03_immune_scoring.py` — Phenotype labels

- ssGSEA + ESTIMATE → immune scores → **Hot / Cold / Intermediate** labels per sample.
- Outputs: `outputs/tables/labels_<cohort>.csv`, `immune_scores_<cohort>.csv`, figures under `outputs/figures/immune/`.

### `04_dimensionality_reduction.py` — Course “representation” methods

- **PCA**, **ICA (FastICA)**, **MDS** (metric + non-metric), **t-SNE**, **UMAP**.
- t-SNE / UMAP typically run on **PCA-reduced** input for stability.
- Figures: `outputs/figures/dim_reduction/` (e.g. PCA scatter coloured by label, variance scree).

**Concrete PCA fact (COAD, from `outputs/tables/pca_variance_TCGA-COAD.csv`):**  
**PC1 alone explains ~32.5%** of variance; **PC1+PC2 ~42%**. (We store up to 50 PCs in that table; cumulative at 50 PCs is ~**80%** — high-dimensional structure remains beyond 2D plots.)

### `05_clustering.py` — Unsupervised structure vs labels

- **k-means** (multiple k), **hierarchical (Ward)**, **DBSCAN**.
- Compare clusters to ssGSEA labels: **ARI** (Adjusted Rand Index), **AMI** (Adjusted Mutual Information), **silhouette** where applicable.
- Example (COAD, `outputs/tables/clustering_metrics_TCGA-COAD.csv`): **k-means k=3** gives **ARI ≈ 0.114**, **AMI ≈ 0.119** — partial agreement (unsupervised clustering does not perfectly reproduce tertile labels; that is expected).

### `06_hypothesis_testing.py` — Core “statistics course” muscle

Per gene and clinical tables:

- **Welch t-test** (Hot vs Cold)
- **One-way ANOVA** (Hot vs Cold vs Intermediate)
- **Kruskal–Wallis** (non-parametric omnibus test across groups)
- **Mann–Whitney U** (non-parametric Hot vs Cold)
- **F-test** for variance (Hot vs Cold)
- **Chi-squared** (and related) for **categorical clinical** variables vs label
- **Multiple testing:** **Benjamini–Hochberg (FDR)** and **Bonferroni** on the relevant p-value columns

Outputs: `outputs/tables/dge_<cohort>.csv`, `dge_sig_<cohort>.csv`, `chi2_clinical_<cohort>.csv`, volcano + heatmaps in `outputs/figures/hypothesis/`.

**High-confidence DE list:** genes with **BH-adjusted** Hot–Cold t-test significant **and** large effect (`dge_sig` — thousands of genes; the PDF report and volcano plot summarise this).

**Example clinical chi-square (COAD):** vital status vs label shows raw p ≈ 0.026; **BH-adjusted** p for that row is ~0.077 — interpret carefully (*trend*, not ironclad after FDR*).

### `07_classification.py` — Six models, not “three or four”

Each model is a **scikit-learn Pipeline** (scaling + **SelectKBest** ANOVA F-score with k=500 features + classifier). Models:

1. **kNN** (k=5)  
2. **SVM linear**  
3. **SVM RBF**  
4. **Logistic regression**  
5. **Random forest** (500 trees)  
6. **XGBoost** (500 estimators, with internal **label encoding** for string class names)

**5-fold stratified cross-validation on TCGA-COAD** — metrics: accuracy, **balanced accuracy**, macro-F1.

**Ranked by balanced accuracy (mean on COAD, `outputs/tables/cv_results_TCGA-COAD.csv`):**

| Model | Balanced accuracy (mean) | Notes |
|-------|--------------------------|--------|
| **RandomForest** | **~0.785** | **Best** — chosen for external validation |
| **SVM-linear** | **~0.785** | Essentially tied with RF |
| Logistic | ~0.752 | Strong linear boundary |
| SVM-RBF | ~0.735 | |
| XGBoost | ~0.704 | Tree boosting; less edge here after heavy feature pre-selection |
| kNN-5 | ~0.648 | Worst — curse of dimensionality even after selection |

Confusion matrices + ROC (one-vs-rest) per top models: `outputs/figures/classification/`.

### `08_survival_analysis.py` — Time-to-event + regression flavour

- **Kaplan–Meier** curves by Hot/Cold/Intermediate.
- **Log-rank** (multivariate across groups).
- **Cox** model on selected immune-score covariates (penalised).
- **PLS** relating high-dimensional expression to survival time (training-set fit; see KM/PLS figures).

**Log-rank (COAD):** test statistic ≈ 5.02, **p ≈ 0.081** — **not significant at α=0.05**, but directionally consistent with “Cold worse than Hot” (underpowered for rare events / censoring — good **limitation** to mention).

### `09_validation_external.py` — The “new dataset” question

- Loads **best model name** from COAD CV results.
- Trains that pipeline on **full COAD** (as implemented), evaluates on **READ** with **gene alignment** and **zero imputation** for missing training genes.
- Output: `external_validation_summary_TCGA-READ.csv`, confusion matrix figure in `outputs/figures/validation/`.

**Note:** The **file header comment** still mentions GEO in places — **actual behaviour follows `config.py`** (READ validation). If asked, say: *“We validated on TCGA-READ per final config.”*

### `10_final_report_figures.py` — Composite figures

- Stitches key panels for the written report / slides.

### `report/generate_report_pdf.py` — Auto PDF

- Builds the **two-column** BT3041 PDF with embedded tables/figures from `outputs/`.

---

## 6. BT3041 syllabus mapping (for “did we use what we learned?”)

Use this as a **checklist slide** in the presentation:

| Course theme | Where it lives in our project |
|--------------|-------------------------------|
| Distributions / normality | QQ / histogram checks in preprocessing figures |
| **t-test**, **ANOVA**, **non-parametric** tests | `06_hypothesis_testing.py` (incl. **Kruskal–Wallis**, **Mann–Whitney U**) |
| **Multiple testing** (Bonferroni, BH/FDR) | Same module, `*_bh`, `*_bonf` columns in `dge_*.csv` |
| **Chi-squared** / contingency thinking | `chi2_clinical_*.csv` |
| **PCA, ICA, MDS** | `04_dimensionality_reduction.py` |
| **t-SNE, UMAP** | Same |
| **k-means, hierarchical, DBSCAN** | `05_clustering.py` |
| **kNN, SVM, logistic, ensembles, boosting** | `07_classification.py` (incl. **XGBoost**) |
| **Cross-validation** | Stratified **5-fold** on COAD |
| **Survival: KM, log-rank, Cox** | `08_survival_analysis.py` |
| **PLS** | Survival module (high-dim → outcome) |
| **External validation** | `09_validation_external.py` (COAD → READ) |

---

## 7. Biological “does it make sense?” (convincing story)

Top differentially expressed genes between Hot and Cold (from `dge_sig_TCGA-COAD.csv` — antigen presentation, interferon-related, immune cell markers) align with **known immunology** (MHC class II, cytotoxic / NK-associated genes, etc.). That supports: **ssGSEA labels + DE analysis are capturing real immune biology**, not random noise.

---

## 8. Conclusions you can state confidently

1. We built a **reproducible 10-step pipeline** from raw TCGA counts to PDF report.  
2. We used **many** statistical tests **with FDR and Bonferroni**, not a single p-hacked test.  
3. **Six** classifiers with **5-fold CV** on COAD; **Random Forest** (tied with linear SVM on CV) was evaluated on **held-out READ**.  
4. **External balanced accuracy ~0.74** on 109 READ samples shows **generalisation**, not just memorising COAD.  
5. **Survival** shows a **trend** consistent with literature but **does not reach p < 0.05** for log-rank on COAD — state this honestly; it shows scientific maturity.  
6. **Limitations:** labels are **computational** (ssGSEA), not IHC; READ validation is **same platform** as COAD, not e.g. microarray; survival is **underpowered** for strong significance.

---

## 9. Team roles (from README — align with who speaks)

| Area | Scripts | Suggested speaker focus |
|------|---------|-------------------------|
| Data + preprocess | `01`, `02` | GDC, batching/retry, CPM/log/filter, why uniform matrices |
| Immune labels | `03` | ssGSEA, ESTIMATE, tertiles, label balance |
| Exploration | `04`, `05` | PCA/UMAP story + clustering vs labels (ARI/AMI) |
| Statistics | `06` | Test battery + FDR + volcano/heatmap |
| ML | `07` | Six models, CV metrics, why RF for deployment |
| Survival + validation | `08`, `09` | KM/log-rank/Cox + **COAD train → READ test** |
| Report / figures | `10`, `report/` | PDF, GitHub, reproducibility one-liner |

Everyone should know **section 4** (train vs READ) and **section 8** (conclusions + limitations) regardless of module.

---

## 10. Quick FAQ for practice

**Q: Did we test on a new dataset?**  
**A:** Yes — **TCGA-READ** was **not** used to train or tune the final external evaluation; it’s a separate cohort of **109** patients.

**Q: Was external validation “positive”?**  
**A:** **Yes** — ~**74%** balanced accuracy vs ~**33%** random baseline for 3 classes.

**Q: How many models?**  
**A:** **Six** (kNN, 2× SVM, Logistic, RandomForest, XGBoost).

**Q: Did we use ANOVA and Kruskal–Wallis?**  
**A:** **Yes**, per gene, plus t, Mann–Whitney, F-variance, chi-square on clinicals, with **BH and Bonferroni**.

**Q: What if they say README still mentions GEO?**  
**A:** Early design included GEO; **final run uses two TCGA cohorts** — practical and scientifically clean. Code truth is `config.py`.

---

## 11. Regenerating everything (for demos)

```bash
cd cold_hot_tumour_project
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python 01_data_acquisition.py   # if raw data missing
python 02_preprocessing.py
python 03_immune_scoring.py
python 04_dimensionality_reduction.py
python 05_clustering.py
python 06_hypothesis_testing.py
python 07_classification.py
python 08_survival_analysis.py
python 09_validation_external.py
python 10_final_report_figures.py
python report/generate_report_pdf.py
```

Large raw/processed data may be **gitignored** — teammates need the data folder per README / Drive instructions.

---

*Document generated for internal team use — BT3041 Cold vs Hot Tumour Classification Project.*
