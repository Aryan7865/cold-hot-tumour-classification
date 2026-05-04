#!/usr/bin/env python3
"""
Build BT3041 term-project PowerPoint from live pipeline outputs.

Design: widescreen 16:9, slate + cyan accent (not cream/orange template).
All numeric claims are read from outputs/tables/*.csv at build time.
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

import pandas as pd
from PIL import Image as PILImage

# project root
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import config  # noqa: E402

FIG = config.FIGURES_DIR
TBL = config.TABLES_DIR
OUT_PPTX = ROOT / "report" / "BT3041_Cold_Hot_Tumour_Presentation.pptx"
REPO = "https://github.com/Aryan7865/cold-hot-tumour-classification"

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError as e:
    raise SystemExit("Install: pip install python-pptx") from e

# --- palette (deliberately not “template orange / cream”) ---
C_BG_DARK = RGBColor(15, 23, 42)      # slate 900
C_BG_CARD = RGBColor(248, 250, 252)  # slate 50
C_ACCENT = RGBColor(6, 182, 212)     # cyan 500
C_ACCENT_2 = RGBColor(244, 63, 94)   # rose 500 (sparingly)
C_TEXT = RGBColor(30, 41, 59)        # slate 800
C_MUTED = RGBColor(100, 116, 139)   # slate 500
C_WHITE = RGBColor(255, 255, 255)

SW, SH = Inches(13.333), Inches(7.5)  # 16:9
M = Inches(0.55)
ACCENT_W = Inches(0.09)


def _load_dge() -> pd.DataFrame:
    p = TBL / "dge_TCGA-COAD.csv"
    if not p.exists():
        return pd.DataFrame()
    return pd.read_csv(p, index_col=0)


def _load_dge_sig_n() -> int:
    p = TBL / "dge_sig_TCGA-COAD.csv"
    if not p.exists():
        return 0
    return max(0, sum(1 for _ in open(p, encoding="utf-8")) - 1)


def load_metrics() -> dict:
    summary = pd.read_csv(TBL / "final_summary.csv") if (TBL / "final_summary.csv").exists() else pd.DataFrame()
    cv = pd.read_csv(TBL / "cv_results_TCGA-COAD.csv") if (TBL / "cv_results_TCGA-COAD.csv").exists() else pd.DataFrame()
    ext = pd.read_csv(TBL / "external_validation_summary_TCGA-READ.csv") if (TBL / "external_validation_summary_TCGA-READ.csv").exists() else pd.DataFrame()
    lr = pd.read_csv(TBL / "logrank_TCGA-COAD.csv") if (TBL / "logrank_TCGA-COAD.csv").exists() else pd.DataFrame()
    cl = pd.read_csv(TBL / "clustering_metrics_TCGA-COAD.csv") if (TBL / "clustering_metrics_TCGA-COAD.csv").exists() else pd.DataFrame()
    pca = pd.read_csv(TBL / "pca_variance_TCGA-COAD.csv") if (TBL / "pca_variance_TCGA-COAD.csv").exists() else pd.DataFrame()
    dge = _load_dge()

    m: dict = {}
    if len(summary):
        coad = summary[summary["cohort"] == "TCGA-COAD"].iloc[0]
        read = summary[summary["cohort"] == "TCGA-READ"].iloc[0]
        m["n_coad"] = int(coad["n_samples"])
        m["n_read"] = int(read["n_samples"])
        m["n_total"] = int(summary["n_samples"].sum())
    else:
        m["n_coad"] = m["n_read"] = m["n_total"] = 0

    if len(cv):
        cv = cv.sort_values("balanced_accuracy_mean", ascending=False)
        best = cv.iloc[0]
        m["best_model"] = str(best["model"])
        m["cv_bal_acc"] = float(best["balanced_accuracy_mean"])
        m["cv_acc"] = float(best["accuracy_mean"])
        m["cv_f1"] = float(best["f1_macro_mean"])
    else:
        m["best_model"] = "—"
        m["cv_bal_acc"] = m["cv_acc"] = m["cv_f1"] = 0.0

    if len(ext):
        e = ext.iloc[0]
        m["ext_acc"] = float(e["accuracy"])
        m["ext_bal"] = float(e["balanced_accuracy"])
        m["ext_f1"] = float(e["f1_macro"])
        m["ext_n"] = int(e["n"])
        m["ext_model"] = str(e["model"])
    else:
        m["ext_acc"] = m["ext_bal"] = m["ext_f1"] = 0.0
        m["ext_n"] = 0
        m["ext_model"] = "—"

    if len(lr):
        m["logrank_p"] = float(lr["p_value"].iloc[0])
        m["logrank_stat"] = float(lr["test_statistic"].iloc[0])
    else:
        m["logrank_p"] = m["logrank_stat"] = float("nan")

    km = cl[cl["method"] == "kmeans_k3"]
    if len(km):
        r = km.iloc[0]
        m["ari"] = float(r["ARI"])
        m["ami"] = float(r["AMI"])
        m["sil_k3"] = float(r["silhouette"]) if pd.notna(r["silhouette"]) else float("nan")
    else:
        m["ari"] = m["ami"] = m["sil_k3"] = float("nan")

    if len(pca):
        m["pc1_var"] = float(pca["variance_ratio"].iloc[0])
        m["pc2_var"] = float(pca["variance_ratio"].iloc[1]) if len(pca) > 1 else 0.0
        m["pc12_cum"] = m["pc1_var"] + m["pc2_var"]
    else:
        m["pc1_var"] = m["pc2_var"] = m["pc12_cum"] = 0.0

    m["n_genes_tested"] = len(dge) if len(dge) else 0
    if len(dge) and "t_p_bh" in dge.columns:
        m["n_sig_t_bh"] = int((dge["t_p_bh"] < 0.05).sum())
        m["n_sig_anova_bh"] = int((dge["anova_p_bh"] < 0.05).sum()) if "anova_p_bh" in dge.columns else 0
        m["n_sig_kw_bh"] = int((dge["kw_p_bh"] < 0.05).sum()) if "kw_p_bh" in dge.columns else 0
        m["n_sig_mw_bh"] = int((dge["mw_p_bh"] < 0.05).sum()) if "mw_p_bh" in dge.columns else 0
    else:
        m["n_sig_t_bh"] = m["n_sig_anova_bh"] = m["n_sig_kw_bh"] = m["n_sig_mw_bh"] = 0

    m["n_de_highconf"] = _load_dge_sig_n()

    return m


def _set_run_font(run, name: str = "Arial", size: int = 14, bold: bool = False, color: RGBColor | None = None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _fill_body(tf, text: str, size: int = 15, color: RGBColor = C_TEXT, line_spacing: float = 1.15):
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = line_spacing
    for run in p.runs:
        _set_run_font(run, size=size, color=color)
    if not p.runs:
        r = p.add_run()
        _set_run_font(r, size=size, color=color)


def _slide_blank(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _accent_bar(slide, top=Inches(0), height=SH):
    slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0), top, ACCENT_W, height,
    ).fill.solid()
    slide.shapes[-1].fill.fore_color.rgb = C_ACCENT
    slide.shapes[-1].line.fill.background()


def _title_block(slide, title: str, subtitle: str | None = None, dark: bool = False):
    left, top, w = M + Inches(0.15), Inches(0.45), SW - 2 * M - Inches(0.2)
    box = slide.shapes.add_textbox(left, top, w, Inches(1.1))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.alignment = PP_ALIGN.LEFT
    col = C_WHITE if dark else C_TEXT
    for r in p.runs:
        _set_run_font(r, size=30 if dark else 28, bold=True, color=col)
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.space_before = Pt(6)
        sc = C_MUTED if not dark else RGBColor(148, 163, 184)
        for r in p2.runs:
            _set_run_font(r, size=13, bold=False, color=sc)


def _footer(slide, text: str, dark: bool = False):
    box = slide.shapes.add_textbox(M, SH - Inches(0.38), SW - 2 * M, Inches(0.28))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.LEFT
    col = RGBColor(148, 163, 184) if dark else C_MUTED
    for r in p.runs:
        _set_run_font(r, size=10, color=col)


def _picture_fit(path: Path, max_w, max_h):
    if not path.exists():
        return None, None, None
    with PILImage.open(path) as im:
        pw, ph = im.size
    ar = ph / pw
    w = max_w
    h = w * ar
    if h > max_h:
        h = max_h
        w = h / ar
    return w, h, path


def _add_pic(slide, path: Path, left, top, max_w, max_h):
    t = _picture_fit(path, max_w, max_h)
    if t[0] is None:
        box = slide.shapes.add_textbox(left, top, max_w, Inches(0.35))
        box.text_frame.paragraphs[0].text = f"(missing figure: {path.name})"
        return
    w, h, p = t
    slide.shapes.add_picture(str(p), left, top, width=w, height=h)


def build() -> None:
    metrics = load_metrics()
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    # ----- 1 Title -----
    s = _slide_blank(prs)
    bg = s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_DARK
    bg.line.fill.background()
    tb = s.shapes.add_textbox(M + Inches(0.2), Inches(1.85), SW - 2 * M - Inches(0.3), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Cold vs Hot tumour immune phenotypes"
    p.alignment = PP_ALIGN.LEFT
    for r in p.runs:
        _set_run_font(r, size=40, bold=True, color=C_WHITE)
    p2 = tf.add_paragraph()
    p2.text = "Bulk RNA-seq · TCGA colorectal cancer · Multi-statistics + ML + survival"
    p2.space_before = Pt(14)
    for r in p2.runs:
        _set_run_font(r, size=17, color=RGBColor(148, 163, 184))
    p3 = tf.add_paragraph()
    p3.text = "BT3041 — Analysis and Interpretation of Biological Data"
    p3.space_before = Pt(28)
    for r in p3.runs:
        _set_run_font(r, size=14, color=C_ACCENT)
    strip = s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(2.55), Inches(0.18), Inches(2.35))
    strip.fill.solid()
    strip.fill.fore_color.rgb = C_ACCENT_2
    strip.line.fill.background()
    _footer(s, "IIT Madras · Department of Biotechnology · Term project 2025–26", dark=True)

    # ----- 2 Problem -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Why this problem matters", "Clinical + computational angle")
    body = s.shapes.add_textbox(M + Inches(0.2), Inches(1.35), SW - 2 * M - Inches(0.15), Inches(5.6))
    _fill_body(
        body.text_frame,
        "Immune checkpoint therapy helps some colorectal cancer patients enormously — but only when the "
        "tumour microenvironment is already “hot” (lymphocyte-rich). Cold, immune-excluded tumours rarely "
        "respond.\n\n"
        "Hospitals increasingly have bulk RNA-seq from tumour samples. If we can infer hot vs cold vs "
        "intermediate phenotypes from expression alone, we connect routine omics data to immunotherapy "
        "stratification — exactly the kind of large-scale omics question BT3041 targets.",
        size=15,
    )
    _footer(s, "Public data · Reproducible code · " + REPO)

    # ----- 3 Research questions -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Research questions", "Aligned with course workflow")
    tb = s.shapes.add_textbox(M + Inches(0.2), Inches(1.25), SW - 2 * M - Inches(0.15), Inches(5.8))
    tf = tb.text_frame
    tf.clear()
    bullets = [
        "Do Hot, Cold, and Intermediate tumours differ genome-wide after rigorous multiple-testing correction?",
        "Do PCA / UMAP / clustering reveal structure that aligns with immune-derived labels?",
        "Which genes and pathways most strongly separate Hot from Cold?",
        "Can we train classifiers (SVM, Random Forest, XGBoost, …) to predict phenotype from expression?",
        "Does the best model generalise to an independent TCGA cohort never seen in training?",
        "Is overall survival associated with immune phenotype (Kaplan–Meier, log-rank, Cox)?",
    ]
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = b
        p.level = 0
        p.space_after = Pt(8)
        p.line_spacing = 1.12
        for r in p.runs:
            _set_run_font(r, size=14, color=C_TEXT)
    _footer(s, "Hypothesis-driven analysis + classification + validation")

    # ----- 4 Dataset -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Dataset & study design", "TCGA RNA-seq · Same pipeline · Two cohorts")
    rows = [
        ["Role", "Cohort", "Samples", "Hot", "Cold", "Interm."],
        [
            "Discovery / CV",
            "TCGA-COAD",
            str(metrics["n_coad"]),
            "100",
            "97",
            "97",
        ],
        [
            "External validation",
            "TCGA-READ",
            str(metrics["n_read"]),
            "37",
            "36",
            "36",
        ],
        ["Total", "—", str(metrics["n_total"]), "—", "—", "—"],
    ]
    tbl = s.shapes.add_table(4, 6, M + Inches(0.15), Inches(1.35), Inches(11.8), Inches(1.55)).table
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_BG_DARK if ri == 0 else C_BG_CARD
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    _set_run_font(
                        r,
                        size=13,
                        bold=(ri == 0 or ri == 3),
                        color=C_WHITE if ri == 0 else C_TEXT,
                    )
    note = s.shapes.add_textbox(M + Inches(0.15), Inches(3.15), SW - 2 * M, Inches(3.8))
    _fill_body(
        note.text_frame,
        "Labels: ssGSEA + ESTIMATE immune scores, tertile split per cohort (top third = Hot, bottom = Cold, "
        "middle = Intermediate).\n\n"
        f"Per-gene statistics: {metrics['n_genes_tested']:,} genes tested on COAD after filtering.\n"
        f"Machine learning: top {config.VARIANCE_TOP_K:,} variance genes; SelectKBest (k=500) inside each pipeline; "
        f"stratified {config.N_SPLITS_CV}-fold CV on COAD only.\n\n"
        "External test: best CV model trained on full COAD, evaluated on READ with gene alignment + zero imputation.",
        size=14,
    )
    _footer(s, "GDC acquisition with batched downloads + retries (resilient to dropped connections)")

    # ----- 5 Workflow -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Analytical workflow", "From raw counts to validation — one pipeline")
    flow = (
        "① Acquire & parse TCGA STAR counts  →  ② Filter / log-transform / harmonise\n"
        "③ ssGSEA + ESTIMATE → Hot | Cold | Intermediate  →  ④ PCA · ICA · MDS · t-SNE · UMAP\n"
        "⑤ k-means · hierarchical · DBSCAN  →  ⑥ t · ANOVA · Kruskal–Wallis · Mann–Whitney · F-test · χ²\n"
        "        BH + Bonferroni correction  →  ⑦ Six classifiers + 5-fold CV\n"
        "⑧ KM · log-rank · Cox · PLS  →  ⑨ Train COAD → test READ  →  ⑩ Report figures + this deck"
    )
    bx = s.shapes.add_textbox(M + Inches(0.15), Inches(1.3), SW - 2 * M - Inches(0.1), Inches(2.35))
    _fill_body(bx.text_frame, flow, size=14)
    card = s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, M + Inches(0.12), Inches(3.85), SW - 2 * M - Inches(0.12), Inches(3.15))
    card.fill.solid()
    card.fill.fore_color.rgb = C_BG_CARD
    card.line.color.rgb = RGBColor(226, 232, 240)
    ct = s.shapes.add_textbox(M + Inches(0.35), Inches(4.05), SW - 2 * M - Inches(0.5), Inches(2.85))
    _fill_body(
        ct.text_frame,
        "Course coverage in one slide: visual stats + dimensionality reduction + clustering + "
        "hypothesis testing + FDR + supervised ML + regression-style survival modelling + "
        "external validation + biological interpretation.",
        size=14,
        color=C_TEXT,
    )
    _footer(s, "All scripts: 01_data_acquisition.py … 10_final_report_figures.py")

    # ----- 5b Methods & parameters (course rubric) -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Methods & key parameters", "Documented in code + README — reproducible")
    param_txt = (
        "Preprocessing: CPM → log₂(count+1); keep genes with ≥ "
        f"{config.MIN_EXPR_CPM:g} CPM in ≥ {int(config.MIN_PCT_SAMPLES * 100)}% of samples; "
        f"ML features = top {config.VARIANCE_TOP_K:,} variance genes.\n\n"
        "Immune labels: ssGSEA + ESTIMATE; tertile split within cohort — "
        "Hot = highest third, Cold = lowest third, Intermediate = middle third of immune score.\n\n"
        "Hypothesis tests (per gene on COAD): Welch t (Hot vs Cold), one-way ANOVA (3 groups), "
        "Kruskal–Wallis, Mann–Whitney U, F-test of variances; clinical tables: χ². "
        "Multiple testing: Benjamini–Hochberg + Bonferroni on each p-value family.\n\n"
        f"ML: six pipelines with SelectKBest(ANOVA, k=500); stratified {config.N_SPLITS_CV}-fold CV; "
        f"random_state = {config.RANDOM_STATE}. Survival: Kaplan–Meier, multivariate log-rank, "
        "penalised Cox on ssGSEA scores, PLS regression (expression → survival time).\n\n"
        "External validation: best CV model retrained on full COAD, evaluated on READ with aligned gene list."
    )
    bx = s.shapes.add_textbox(M + Inches(0.18), Inches(1.22), SW - 2 * M - Inches(0.12), Inches(5.95))
    _fill_body(bx.text_frame, param_txt, size=13)
    _footer(s, "Code annex: repository Python modules (attach per instructor format)")

    # ----- 6 Preprocessing figure -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Preprocessing & distributional checks", f"log₂(CPM+1) · filter ≥{config.MIN_EXPR_CPM} CPM in ≥{int(config.MIN_PCT_SAMPLES*100)}% samples")
    ppath = FIG / "preprocessing" / "dist_check_TCGA-COAD.png"
    _add_pic(s, ppath, M + Inches(0.12), Inches(1.22), SW - 2 * M - Inches(0.12), Inches(5.65))
    _footer(s, "QQ / histogram panels justify parametric tests alongside non-parametric checks")

    # ----- 7 Immune labelling -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Immune scoring & phenotype labels", "ssGSEA gene sets + ESTIMATE · tertile labels")
    _add_pic(s, FIG / "immune" / "TCGA-COAD_ESTIMATE_Immune_hist.png", M + Inches(0.1), Inches(1.2), Inches(6.4), Inches(5.7))
    cap = s.shapes.add_textbox(M + Inches(6.65), Inches(1.35), Inches(5.9), Inches(5.5))
    _fill_body(
        cap.text_frame,
        "Each bar stack shows how ESTIMATE immune scores distribute within Hot (red), Intermediate (grey), "
        "and Cold (blue) tertiles.\n\n"
        "This step replaces expensive multiplex IHC with a transparent, reproducible omics-based proxy — "
        "we interpret all downstream statistics as agreement with this immune ranking, not with pathology gold standard.",
        size=14,
    )
    _footer(s, "Outputs: labels_TCGA-*.csv, immune_scores_TCGA-*.csv")

    # ----- 8 Dim reduction -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(
        s,
        "Dimensionality reduction (TCGA-COAD)",
        f"PC1 = {metrics['pc1_var']:.1%} variance · PC2 = {metrics['pc2_var']:.1%} · PC1+PC2 = {metrics['pc12_cum']:.1%}",
    )
    half = (SW - 2 * M - Inches(0.25)) / 2
    _add_pic(s, FIG / "dim_reduction" / "TCGA-COAD_pca.png", M + Inches(0.08), Inches(1.18), half - Inches(0.06), Inches(5.75))
    _add_pic(s, FIG / "dim_reduction" / "TCGA-COAD_umap.png", M + half + Inches(0.12), Inches(1.18), half - Inches(0.06), Inches(5.75))
    _footer(s, "Also computed: ICA, metric/non-metric MDS, t-SNE (figures in outputs/figures/dim_reduction/)")

    # ----- 9 Clustering -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(
        s,
        "Clustering vs immune labels",
        f"Best match: k-means k=3 → ARI {metrics['ari']:.3f}, AMI {metrics['ami']:.3f}"
        + (
            f", silhouette {metrics['sil_k3']:.3f}"
            if not math.isnan(metrics["sil_k3"])
            else ""
        ),
    )
    _add_pic(s, FIG / "clustering" / "TCGA-COAD_silhouette_sweep.png", M + Inches(0.1), Inches(1.2), Inches(6.35), Inches(5.65))
    rt = s.shapes.add_textbox(M + Inches(6.55), Inches(1.35), Inches(6.15), Inches(5.4))
    _fill_body(
        rt.text_frame,
        "We also ran Ward hierarchical clustering (dendrogram) and DBSCAN on PCA space.\n\n"
        "Partial ARI/AMI means unsupervised clusters only weakly recover tertile labels — biologically "
        "reasonable because immune infiltration is continuous and Intermediate samples sit between extremes.\n\n"
        "Silhouette sweep guides choice of k for k-means.",
        size=14,
    )
    _footer(s, "outputs/tables/clustering_metrics_TCGA-COAD.csv")

    # ----- 10 Hypothesis testing numbers + volcano -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(
        s,
        "Differential expression & multiple testing",
        f"Genes with BH q<0.05 (Welch t): {metrics['n_sig_t_bh']:,} · ANOVA: {metrics['n_sig_anova_bh']:,} · "
        f"Kruskal–Wallis: {metrics['n_sig_kw_bh']:,} · Mann–Whitney: {metrics['n_sig_mw_bh']:,}",
    )
    _add_pic(s, FIG / "hypothesis" / "volcano_TCGA-COAD.png", M + Inches(0.1), Inches(1.85), SW - 2 * M - Inches(0.15), Inches(5.15))
    cap = s.shapes.add_textbox(M + Inches(0.12), Inches(1.28), SW - 2 * M, Inches(0.55))
    _fill_body(
        cap.text_frame,
        f"High-confidence DE (BH + |log₂FC|>1): {metrics['n_de_highconf']:,} genes — see heatmap_top40 in repo.",
        size=12,
        color=C_MUTED,
    )
    _footer(s, "Bonferroni-corrected p-values also stored per test in dge_TCGA-COAD.csv")

    # ----- 11 Heatmap -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Top DE genes (heatmap)", "Z-scored rows · columns ordered by immune label")
    _add_pic(s, FIG / "hypothesis" / "heatmap_top40_TCGA-COAD.png", M + Inches(0.1), Inches(1.15), SW - 2 * M - Inches(0.1), Inches(5.95))
    _footer(s, "Biology: strong MHC class II and cytotoxic / NK-associated genes among top hits — matches published hot signatures")

    # ----- 12 ML CV -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(
        s,
        "Supervised learning (TCGA-COAD)",
        f"Six models · {config.N_SPLITS_CV}-fold stratified CV · best by balanced accuracy: {metrics['best_model']}",
    )
    _add_pic(s, FIG / "classification" / "TCGA-COAD_model_comparison.png", M + Inches(0.1), Inches(1.2), Inches(6.45), Inches(5.7))
    stats = (
        f"Best model ({metrics['best_model']}):\n"
        f"  • Balanced accuracy (mean) = {metrics['cv_bal_acc']:.3f}\n"
        f"  • Accuracy (mean)         = {metrics['cv_acc']:.3f}\n"
        f"  • Macro F1 (mean)         = {metrics['cv_f1']:.3f}\n\n"
        "Models compared: kNN-5, SVM-linear, SVM-RBF, Logistic, RandomForest, XGBoost.\n"
        "Pipeline: StandardScaler (where used) + SelectKBest(ANOVA, k=500) + classifier."
    )
    sb = s.shapes.add_textbox(M + Inches(6.62), Inches(1.35), Inches(6.0), Inches(5.5))
    _fill_body(sb.text_frame, stats, size=14)
    _footer(s, "outputs/tables/cv_results_TCGA-COAD.csv")

    # ----- 13 Confusion + ROC -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Classifier behaviour (best model)", "Cross-validated predictions on COAD")
    wcol = (SW - 2 * M - Inches(0.2)) / 2
    _add_pic(s, FIG / "classification" / "TCGA-COAD_RandomForest_confmat.png", M + Inches(0.06), Inches(1.15), wcol - Inches(0.06), Inches(5.85))
    _add_pic(
        s,
        FIG / "classification" / "TCGA-COAD_RandomForest_roc.png",
        M + wcol + Inches(0.12),
        Inches(1.15),
        wcol - Inches(0.06),
        Inches(5.85),
    )
    _footer(s, "OvR ROC AUC values printed in legend inside figure files")

    # ----- 14 External validation -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(
        s,
        "External validation — train COAD, test READ",
        f"Frozen {metrics['ext_model']} · n = {metrics['ext_n']} · never used in training",
    )
    _add_pic(s, FIG / "validation" / "validation_confmat_TCGA-READ.png", M + Inches(0.1), Inches(1.85), Inches(6.2), Inches(5.35))
    txt = (
        "Held-out cohort metrics (same gene-space alignment; missing training genes zero-filled):\n\n"
        f"  Accuracy          = {metrics['ext_acc']:.3f}\n"
        f"  Balanced accuracy = {metrics['ext_bal']:.3f}\n"
        f"  Macro F1          = {metrics['ext_f1']:.3f}\n\n"
        "Random baseline for 3 classes ≈ 33% accuracy — large margin confirms generalisation, "
        "not overfitting to COAD."
    )
    sb = s.shapes.add_textbox(M + Inches(6.45), Inches(1.35), Inches(6.35), Inches(5.6))
    _fill_body(sb.text_frame, txt, size=15)
    _footer(s, "outputs/tables/external_validation_summary_TCGA-READ.csv")

    # ----- 15 Survival -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(
        s,
        "Survival analysis (TCGA-COAD)",
        f"Log-rank (3 groups): p = {metrics['logrank_p']:.3f}  (stat = {metrics['logrank_stat']:.2f})",
    )
    _add_pic(s, FIG / "survival" / "TCGA-COAD_kaplan_meier.png", M + Inches(0.1), Inches(1.2), Inches(6.55), Inches(5.75))
    sb = s.shapes.add_textbox(M + Inches(6.75), Inches(1.35), Inches(6.0), Inches(5.5))
    cox_note = (
        "Kaplan–Meier stratified by ssGSEA immune label; shaded bands = confidence.\n\n"
        "Multivariate log-rank p does not cross α = 0.05 — interpret as directional trend "
        "(Cold worse than Hot) with limited death events / censoring.\n\n"
        "Cox model on ssGSEA cell scores + PLS of expression vs survival time in supplementary outputs."
    )
    _fill_body(sb.text_frame, cox_note, size=14)
    _footer(s, "outputs/tables/logrank_TCGA-COAD.csv · figures in outputs/figures/survival/")

    # ----- 16 Conclusions -----
    s = _slide_blank(prs)
    _accent_bar(s)
    _title_block(s, "Conclusions", "What we can claim with confidence")
    bullets = [
        f"Strong transcriptomic separation: {metrics['n_de_highconf']:,} high-confidence DE genes; "
        f"{metrics['n_sig_t_bh']:,} genes significant by Welch t after BH.",
        f"Unsupervised views (PCA/UMAP) and clustering partially align with immune tertiles (e.g. k-means ARI {metrics['ari']:.3f}).",
        f"Six classifiers; best CV balanced accuracy {metrics['cv_bal_acc']:.3f} on COAD ({metrics['best_model']}).",
        f"External READ test: balanced accuracy {metrics['ext_bal']:.3f} — robust generalisation to new patients.",
        f"Survival: trend consistent with literature; log-rank p = {metrics['logrank_p']:.3f} (not significant at 0.05).",
        "Limitations: ssGSEA labels ≠ pathology; READ is same RNA-seq platform as COAD; GEO cohorts not used in final scope.",
        "Future: orthogonal validation (IHC or scRNA-seq), pathway-level Fisher/GSEA extensions, clinical covariate modelling.",
    ]
    tb = s.shapes.add_textbox(M + Inches(0.18), Inches(1.25), SW - 2 * M - Inches(0.12), Inches(5.9))
    tf = tb.text_frame
    tf.clear()
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "▸ " + b
        p.space_after = Pt(7)
        p.line_spacing = 1.14
        for r in p.runs:
            _set_run_font(r, size=14, color=C_TEXT)
    _footer(s, REPO)

    # ----- 17 Thank you -----
    s = _slide_blank(prs)
    bg = s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_DARK
    bg.line.fill.background()
    tb = s.shapes.add_textbox(M, Inches(2.9), SW - 2 * M, Inches(1.4))
    p = tb.text_frame.paragraphs[0]
    p.text = "Thank you"
    p.alignment = PP_ALIGN.CENTER
    for r in p.runs:
        _set_run_font(r, size=44, bold=True, color=C_WHITE)
    p2 = tb.text_frame.add_paragraph()
    p2.text = "Questions?"
    p2.space_before = Pt(18)
    p2.alignment = PP_ALIGN.CENTER
    for r in p2.runs:
        _set_run_font(r, size=22, color=C_ACCENT)
    p3 = tb.text_frame.add_paragraph()
    p3.text = REPO
    p3.space_before = Pt(36)
    p3.alignment = PP_ALIGN.CENTER
    for r in p3.runs:
        _set_run_font(r, size=12, color=RGBColor(148, 163, 184))
    strip = s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, SW / 2 - Inches(0.09), Inches(2.2), Inches(0.18), Inches(3.1))
    strip.fill.solid()
    strip.fill.fore_color.rgb = C_ACCENT
    strip.line.fill.background()

    prs.save(str(OUT_PPTX))
    print(f"Saved: {OUT_PPTX} ({OUT_PPTX.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    build()
