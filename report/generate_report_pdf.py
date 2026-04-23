"""
Generate the BT3041 final project report in an IEEE-style two-column
research-paper layout, with every figure strictly sized to the column width
(no overflow), and with rich end-to-end content pulled live from outputs/.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Sequence

import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    CondPageBreak,
    Frame,
    FrameBreak,
    Image,
    KeepTogether,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


# --------------------------------------------------------------------------- #
# Paths and constants                                                         #
# --------------------------------------------------------------------------- #

PROJECT_ROOT = Path(__file__).resolve().parents[1]
FIG_DIR = PROJECT_ROOT / "outputs" / "figures"
TBL_DIR = PROJECT_ROOT / "outputs" / "tables"
OUT_PDF = PROJECT_ROOT / "report" / "BT3041_Cold_Hot_Tumour_Report.pdf"

REPO_URL = "https://github.com/Aryan7865/cold-hot-tumour-classification"

PAGE_W, PAGE_H = A4
LEFT = 0.55 * inch
RIGHT = 0.55 * inch
TOP = 0.75 * inch
BOTTOM = 0.65 * inch
GUTTER = 0.25 * inch

CONTENT_W = PAGE_W - LEFT - RIGHT
COL_W = (CONTENT_W - GUTTER) / 2  # ~ 3.24 inch
MAX_FIG_W_IN = (COL_W - 4) / inch  # inches, leave 4 pts safety

# Full-width figure width (spans both columns)
MAX_WIDE_FIG_W_IN = (CONTENT_W - 6) / inch

SHORT_TITLE = "Cold vs Hot Tumour Classification from Bulk RNA-seq"
AUTHOR_LINE = "BT3041 Project Team"

# --------------------------------------------------------------------------- #
# Styles                                                                      #
# --------------------------------------------------------------------------- #

_base = getSampleStyleSheet()

STYLES = {
    "title": ParagraphStyle(
        "title",
        parent=_base["Title"],
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=21,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#0c2d57"),
        spaceAfter=8,
    ),
    "affil_line": ParagraphStyle(
        "affil_line",
        parent=_base["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1a1a1a"),
        spaceAfter=2,
    ),
    "meta_line": ParagraphStyle(
        "meta_line",
        parent=_base["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=9,
        leading=12,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#4a4a4a"),
        spaceAfter=4,
    ),
    "abstract_label": ParagraphStyle(
        "abstract_label",
        parent=_base["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        textColor=colors.HexColor("#0c2d57"),
        spaceAfter=3,
    ),
    "abstract_body": ParagraphStyle(
        "abstract_body",
        parent=_base["BodyText"],
        fontName="Times-Roman",
        fontSize=9.4,
        leading=12.5,
        alignment=TA_JUSTIFY,
    ),
    "h1": ParagraphStyle(
        "h1",
        parent=_base["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=11.5,
        leading=14,
        textColor=colors.HexColor("#123b64"),
        spaceBefore=8,
        spaceAfter=4,
    ),
    "h2": ParagraphStyle(
        "h2",
        parent=_base["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=9.8,
        leading=12,
        textColor=colors.HexColor("#123b64"),
        spaceBefore=5,
        spaceAfter=2,
    ),
    "body": ParagraphStyle(
        "body",
        parent=_base["BodyText"],
        fontName="Times-Roman",
        fontSize=9.3,
        leading=12.5,
        alignment=TA_JUSTIFY,
        spaceAfter=4,
    ),
    "bullet": ParagraphStyle(
        "bullet",
        parent=_base["BodyText"],
        fontName="Times-Roman",
        fontSize=9.3,
        leading=12.2,
        leftIndent=12,
        bulletIndent=2,
        alignment=TA_JUSTIFY,
        spaceAfter=1,
    ),
    "cap": ParagraphStyle(
        "cap",
        parent=_base["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=7.8,
        leading=10,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#2a2a2a"),
        spaceAfter=6,
    ),
    "link": ParagraphStyle(
        "link",
        parent=_base["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=11.5,
        alignment=TA_LEFT,
    ),
    "refs": ParagraphStyle(
        "refs",
        parent=_base["BodyText"],
        fontName="Times-Roman",
        fontSize=8.7,
        leading=11,
        alignment=TA_JUSTIFY,
        leftIndent=12,
        bulletIndent=0,
        spaceAfter=2,
    ),
}


# --------------------------------------------------------------------------- #
# Page decorations                                                            #
# --------------------------------------------------------------------------- #


def _page_deco(canvas, doc) -> None:
    canvas.saveState()
    # Top header rule
    canvas.setStrokeColor(colors.HexColor("#123b64"))
    canvas.setLineWidth(0.6)
    y_rule = PAGE_H - TOP + 16
    canvas.line(LEFT, y_rule, PAGE_W - RIGHT, y_rule)
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor("#123b64"))
    canvas.drawString(LEFT, y_rule + 4, SHORT_TITLE)
    canvas.drawRightString(PAGE_W - RIGHT, y_rule + 4, AUTHOR_LINE + " — IIT Madras")
    # Footer
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawCentredString(PAGE_W / 2, 16, f"Page {doc.page}")
    canvas.restoreState()


# --------------------------------------------------------------------------- #
# Doc and frame layout                                                        #
# --------------------------------------------------------------------------- #


def build_doc() -> BaseDocTemplate:
    doc = BaseDocTemplate(
        str(OUT_PDF),
        pagesize=A4,
        leftMargin=LEFT,
        rightMargin=RIGHT,
        topMargin=TOP,
        bottomMargin=BOTTOM,
        title="BT3041 Term Project Report — Cold vs Hot Tumour Classification",
        author=AUTHOR_LINE,
    )

    # Full-width page (page 1 only) - for title and abstract.
    full_frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height,
        id="full",
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )

    # Two-column page template (pages 2+).
    frame_l = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        COL_W,
        doc.height,
        id="L",
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )
    frame_r = Frame(
        doc.leftMargin + COL_W + GUTTER,
        doc.bottomMargin,
        COL_W,
        doc.height,
        id="R",
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )

    doc.addPageTemplates(
        [
            PageTemplate(id="First", frames=[full_frame], onPage=_page_deco),
            PageTemplate(id="TwoCol", frames=[frame_l, frame_r], onPage=_page_deco),
        ]
    )
    return doc


# --------------------------------------------------------------------------- #
# Image / figure helpers with STRICT sizing                                   #
# --------------------------------------------------------------------------- #


def _sized_image(path: Path, max_width_in: float) -> Image | None:
    if not path.exists():
        return None
    with PILImage.open(path) as im:
        px_w, px_h = im.size
    aspect = px_h / px_w
    w_pts = max_width_in * inch
    h_pts = w_pts * aspect
    img = Image(str(path), width=w_pts, height=h_pts)
    img.hAlign = "CENTER"
    return img


def fig_col(path: Path, caption: str) -> list:
    """Column-width figure. Max width = MAX_FIG_W_IN inches."""
    img = _sized_image(path, MAX_FIG_W_IN)
    if img is None:
        return [Paragraph(f"<i>[missing: {path.name}]</i>", STYLES["cap"])]
    return [img, Spacer(1, 1), Paragraph(caption, STYLES["cap"])]


def fig_wide(path: Path, caption: str) -> list:
    """Alias of fig_col: we keep all figures within a single column so that
    they never bleed into the neighbouring column in the 2-col layout."""
    return fig_col(path, caption)


# --------------------------------------------------------------------------- #
# Table helpers                                                               #
# --------------------------------------------------------------------------- #


def _academic_style(n_rows: int) -> TableStyle:
    return TableStyle(
        [
            # Header row: bold, double line separators
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.2),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f8")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#0c2d57")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            # Thin horizontal borders only — classic academic look
            ("LINEABOVE", (0, 0), (-1, 0), 0.75, colors.HexColor("#0c2d57")),
            ("LINEBELOW", (0, 0), (-1, 0), 0.55, colors.HexColor("#0c2d57")),
            ("LINEBELOW", (0, -1), (-1, -1), 0.75, colors.HexColor("#0c2d57")),
            # Inner row separators — very light
            ("LINEBELOW", (0, 1), (-1, -2), 0.2, colors.HexColor("#c9d3e2")),
            # Padding
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
    )


def col_table(
    rows: Sequence[Sequence[str]], col_widths_in: Sequence[float], caption: str
) -> list:
    t = Table(rows, colWidths=[w * inch for w in col_widths_in])
    t.setStyle(_academic_style(len(rows)))
    t.hAlign = "CENTER"
    return [t, Spacer(1, 2), Paragraph(caption, STYLES["cap"])]


# --------------------------------------------------------------------------- #
# Content loader                                                              #
# --------------------------------------------------------------------------- #


def load_tables() -> dict[str, pd.DataFrame]:
    def _load(name: str) -> pd.DataFrame:
        p = TBL_DIR / name
        if not p.exists():
            return pd.DataFrame()
        if name.startswith("dge_"):
            return pd.read_csv(p, index_col=0)
        return pd.read_csv(p)

    return {
        "summary": _load("final_summary.csv"),
        "cv_coad": _load("cv_results_TCGA-COAD.csv"),
        "cv_read": _load("cv_results_TCGA-READ.csv"),
        "ext": _load("external_validation_summary_TCGA-READ.csv"),
        "dge_coad": _load("dge_TCGA-COAD.csv"),
        "dge_sig_coad": _load("dge_sig_TCGA-COAD.csv"),
        "dge_read": _load("dge_TCGA-READ.csv"),
        "dge_sig_read": _load("dge_sig_TCGA-READ.csv"),
        "logrank_coad": _load("logrank_TCGA-COAD.csv"),
        "logrank_read": _load("logrank_TCGA-READ.csv"),
        "clust_coad": _load("clustering_metrics_TCGA-COAD.csv"),
        "clust_read": _load("clustering_metrics_TCGA-READ.csv"),
        "pca_coad": _load("pca_variance_TCGA-COAD.csv"),
        "pca_read": _load("pca_variance_TCGA-READ.csv"),
        "cox_coad": _load("cox_TCGA-COAD.csv"),
        "cox_read": _load("cox_TCGA-READ.csv"),
        "chi2_coad": _load("chi2_clinical_TCGA-COAD.csv"),
        "chi2_read": _load("chi2_clinical_TCGA-READ.csv"),
    }


# --------------------------------------------------------------------------- #
# Story builder                                                               #
# --------------------------------------------------------------------------- #


def P(txt: str, style: str = "body") -> Paragraph:
    return Paragraph(txt, STYLES[style])


def bullet(txt: str) -> Paragraph:
    return Paragraph("&bull; " + txt, STYLES["bullet"])


def abstract_box(text: str) -> Table:
    inner = [
        [Paragraph("<b>Abstract</b>", STYLES["abstract_label"])],
        [Paragraph(text, STYLES["abstract_body"])],
    ]
    tbl = Table(inner, colWidths=[CONTENT_W - 8])
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eaf1f8")),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#9fb3c8")),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return tbl


def build_story(t: dict[str, pd.DataFrame]) -> list:
    story: list = []

    # ========= PAGE 1 — Title + Abstract (full width) ========= #
    story.append(P("BT3041 — Analysis and Interpretation of Biological Project Data : Report", "meta_line"))
    story.append(Spacer(1, 2))
    story.append(
        P(
            "Cold vs Hot Tumour Classification from Bulk RNA-seq "
            "using a Multi-Statistical and Machine-Learning Pipeline",
            "title",
        )
    )
    story.append(P("BT3041 Project Team &mdash; Group of 6, 2025&ndash;26", "affil_line"))
    story.append(P("Indian Institute of Technology Madras &mdash; Department of Biotechnology", "affil_line"))
    story.append(P(f"Date of Submission : {date.today().strftime('%B %d, %Y')}", "meta_line"))
    story.append(Spacer(1, 6))

    abstract_text = (
        "The tumour immune micro-environment is a primary determinant of response to immune checkpoint "
        "blockade therapy. Tumours are broadly divided into three immune phenotypes: <b>hot</b> (strongly "
        "infiltrated by cytotoxic immune cells, typically responsive to immunotherapy), <b>cold</b> "
        "(immune-excluded, poor responders), and <b>intermediate</b>. In this project we design, execute "
        "and externally validate a complete and reproducible pipeline that predicts the immune phenotype "
        "of colorectal tumours from bulk RNA-sequencing profiles alone. We use two independent TCGA "
        "cohorts: TCGA-COAD (<b>n=294</b>, discovery/training) and TCGA-READ (<b>n=109</b>, external "
        "validation) — 403 patients in total, 633 linked clinical records, and 59,429 annotated genes per "
        "cohort. The pipeline integrates CPM filtering, log-transform and top-variance feature selection, "
        "ssGSEA / ESTIMATE immune scoring, five dimensionality-reduction methods (PCA, ICA, MDS, t-SNE, "
        "UMAP), three clustering algorithms (k-means, Ward hierarchical, DBSCAN), six hypothesis tests "
        "(Welch t, one-way ANOVA, Kruskal-Wallis, Mann-Whitney U, F-variance, χ²) with both Benjamini-"
        "Hochberg and Bonferroni multiple-testing correction, six supervised classifiers (kNN, linear/"
        "RBF SVM, logistic regression, random forest, XGBoost) under stratified 5-fold cross-validation, "
        "Kaplan-Meier and Cox proportional-hazards regression, and Partial Least Squares regression of "
        "survival. The best classifier reaches <b>0.785 balanced accuracy</b> on the discovery cohort "
        "(5-fold CV) and <b>0.744 balanced accuracy</b> on the held-out external cohort. More than "
        "<b>2,269 genes</b> are significantly differentially expressed between hot and cold tumours after "
        "BH correction with |log<sub>2</sub>FC|&nbsp;&gt;&nbsp;1, led by HLA class-II antigen-presentation "
        "genes (HLA-DPB1, HLA-DRA, HLA-DPA1, HLA-DQA1) and cytotoxic effectors (PLA2G2D, SLAMF7) — a "
        "textbook hot-tumour immune signature."
    )
    story.append(abstract_box(abstract_text))
    story.append(Spacer(1, 6))
    story.append(
        P(
            "<b>Keywords:</b> colorectal cancer, tumour immune phenotype, ssGSEA, ESTIMATE, RNA-seq, PCA, "
            "UMAP, hypothesis testing, machine learning, survival analysis, external validation.",
            "body",
        )
    )
    story.append(Spacer(1, 6))

    # Switch to 2-col for all remaining pages
    story.append(NextPageTemplate("TwoCol"))
    story.append(PageBreak())

    # ========= 1. Introduction ========= #
    story.append(P("1. Introduction", "h1"))
    story.append(
        P(
            "<b>C</b>olorectal cancer (CRC) remains one of the leading causes of cancer-related death "
            "worldwide. Although cytotoxic chemotherapy and molecularly targeted agents have improved "
            "outcomes, response remains heterogeneous and many patients still progress. Immunotherapy, in "
            "particular immune checkpoint inhibitors, has transformed the management of several cancers, "
            "but its efficacy in CRC depends strongly on the <i>tumour immune phenotype</i>. Tumours that "
            "are densely infiltrated with CD8<sup>+</sup> T-cells, natural killer (NK) cells and type-I "
            "interferon signalling are termed <b>immune-hot</b> and often respond to checkpoint blockade, "
            "whereas <b>immune-cold</b> tumours are depleted of lymphocytes and usually fail to respond."
        )
    )
    story.append(
        P(
            "Measuring the immune phenotype directly &mdash; for instance via multiplexed "
            "immunohistochemistry &mdash; is expensive and is not routinely available outside tertiary "
            "centres. Bulk RNA sequencing, on the other hand, is increasingly accessible and captures "
            "signatures of immune infiltration indirectly through gene expression. The goal of this "
            "project is therefore to build, rigorously statistically validate, and externally test a "
            "machine-learning pipeline that infers the immune phenotype of a colorectal tumour from its "
            "transcriptome alone."
        )
    )

    story.append(P("1.1 Research Objectives", "h2"))
    story.append(P("The project is organised around four objectives:"))
    story.append(bullet("<b>(O1)</b> <b>Data acquisition and harmonisation.</b> Reproducibly download TCGA STAR-count expression data and clinical annotation for two CRC cohorts, and build pre-processed expression matrices on a shared gene-symbol feature space."))
    story.append(bullet("<b>(O2)</b> <b>Immune phenotype labelling.</b> Derive per-sample immune scores using ssGSEA and ESTIMATE and assign every patient to a Hot / Cold / Intermediate class by tertile-splitting."))
    story.append(bullet("<b>(O3)</b> <b>Descriptive and inferential statistics.</b> Quantify differences between groups using parametric and non-parametric hypothesis tests with rigorous multiple-testing correction; compare unsupervised clustering to ssGSEA labels."))
    story.append(bullet("<b>(O4)</b> <b>Predictive modelling and external validation.</b> Train six supervised classifiers on TCGA-COAD and evaluate generalisation on the held-out TCGA-READ cohort; additionally relate the immune phenotype to overall survival via Kaplan-Meier, log-rank and Cox regression."))

    story.append(P("1.2 Why this topic satisfies the course rubric", "h2"))
    story.append(
        P(
            "The Hot/Cold classification problem is a rare natural fit for BT3041 because it "
            "legitimately requires, in a biologically meaningful way: (i) distribution checks on the "
            "normality of expression; (ii) hypothesis testing across two and three groups with both "
            "parametric and non-parametric tests and FDR correction; (iii) dimensionality reduction "
            "with multiple methods; (iv) unsupervised clustering benchmarked by ARI / AMI; (v) "
            "supervised classifiers with cross-validation; (vi) linear / logistic / PLS regression; "
            "and (vii) survival analysis. Every BT3041 module is therefore used in a biologically "
            "purposeful way rather than as an isolated exercise."
        )
    )

    # ========= 2. Materials and Methods ========= #
    story.append(P("2. Materials and Methods", "h1"))

    story.append(P("2.1 Cohorts and data acquisition", "h2"))
    story.append(
        P(
            "We used two open-access RNA-seq cohorts from the NCI Genomic Data Commons (GDC). "
            "TCGA-COAD (colon adenocarcinoma) served as the primary discovery cohort and "
            "TCGA-READ (rectum adenocarcinoma) as a fully independent external validation cohort. Both "
            "were generated on the same TCGA pipeline (Illumina paired-end sequencing, STAR alignment, "
            "unstranded counts), eliminating cross-platform normalisation artefacts. Expression files "
            "and clinical metadata were retrieved programmatically via the GDC REST API. Because GDC "
            "frequently drops the connection on large single-shot bulk downloads, we implemented a "
            "resumable batched downloader (50 files per batch with exponential back-off retries). After "
            "download, per-sample STAR-count files were parsed and merged into cohort-level expression "
            "matrices using <i>gene_name</i> (HGNC symbol) as the row index; duplicate PAR_Y aliases "
            "were collapsed by summation and replicate aliquots were disambiguated by suffixing."
        )
    )

    # Table 1 — cohort composition
    summary = t["summary"]
    row = [["Cohort", "n", "Hot", "Cold", "Int.", "Role"]]
    role_map = {"TCGA-COAD": "Discovery", "TCGA-READ": "External"}
    for _, r in summary.iterrows():
        row.append(
            [
                r["cohort"],
                int(r["n_samples"]),
                int(r["Hot"]),
                int(r["Cold"]),
                int(r["Intermediate"]),
                role_map.get(r["cohort"], ""),
            ]
        )
    row.append(
        [
            "Total",
            int(summary["n_samples"].sum()),
            int(summary["Hot"].sum()),
            int(summary["Cold"].sum()),
            int(summary["Intermediate"].sum()),
            "—",
        ]
    )
    story.extend(
        col_table(
            row,
            col_widths_in=[0.80, 0.32, 0.34, 0.34, 0.34, 0.65],
            caption="Table 1. Final cohort composition after preprocessing and ssGSEA-based labelling.",
        )
    )

    story.append(P("2.2 Preprocessing and harmonisation", "h2"))
    story.append(
        P(
            "Raw counts were converted to CPM (counts per million), and lowly expressed genes were "
            "filtered with the criterion <i>CPM &ge; 1 in &ge; 20% of samples</i>, retaining "
            "<b>16,723</b> genes in TCGA-COAD and <b>16,623</b> genes in TCGA-READ. Expression values "
            "were then log<sub>2</sub>-transformed with a pseudocount of 1. For downstream machine "
            "learning we retained the <b>top 5,000 most-variable genes</b> per cohort. For cross-cohort "
            "analyses, log-expression matrices were intersected on HGNC gene symbols, producing "
            "<b>16,453 common genes</b> that share identical preprocessing and distributional properties. "
            "Shapiro-Wilk tests and QQ-plots on a random sample of genes (Fig.&nbsp;1) supported "
            "approximate normality of the log-transformed values, legitimising the use of parametric "
            "tests in later sections."
        )
    )
    # QQ / hist figure as wide figure
    story.extend(
        fig_wide(
            FIG_DIR / "preprocessing" / "dist_check_TCGA-COAD.png",
            "Figure&nbsp;1. Distribution diagnostics for five randomly-sampled genes in TCGA-COAD "
            "after log<sub>2</sub>(CPM+1) transformation. Top row: histograms; bottom row: QQ-plots "
            "against a standard normal. Near-linear QQ-plots support the use of parametric Welch-t "
            "and ANOVA tests.",
        )
    )

    story.append(P("2.3 Immune phenotype scoring and Hot/Cold/Intermediate labelling", "h2"))
    story.append(
        P(
            "Per-sample immune scores were computed using single-sample Gene Set Enrichment Analysis "
            "(ssGSEA) via <i>gseapy</i> on twelve curated immune gene sets, combining the ESTIMATE "
            "immune / stromal signatures with Bindea cell-type signatures (CD8 T-cells, cytotoxic "
            "cells, Tregs, Th1, Th2, macrophages, NK-cells, B-cells, dendritic cells, neutrophils). "
            "The <i>ESTIMATE_Immune</i> score was used to rank patients and split them into equal "
            "tertiles: top third — <b>Hot</b>, bottom third — <b>Cold</b>, middle — "
            "<b>Intermediate</b>. This yielded approximately balanced classes in both cohorts "
            "(Table&nbsp;1, Fig.&nbsp;2)."
        )
    )
    story.append(
        KeepTogether(
            fig_col(
                FIG_DIR / "immune" / "TCGA-COAD_ESTIMATE_Immune_hist.png",
                "Figure&nbsp;2. ESTIMATE_Immune score distribution in TCGA-COAD, stratified by the "
                "assigned Hot / Cold / Intermediate label.",
            )
        )
    )

    story.append(P("2.4 Dimensionality reduction", "h2"))
    story.append(
        P(
            "Every sample was projected into 2-D using five complementary techniques: linear Principal "
            "Component Analysis (PCA) and Independent Component Analysis (FastICA); Multidimensional "
            "Scaling in both metric and non-metric variants; t-Stochastic Neighbour Embedding (t-SNE); "
            "and Uniform Manifold Approximation and Projection (UMAP). To avoid perplexity degeneracy "
            "on high-dimensional inputs, both t-SNE and UMAP were trained on the first 20 PCA "
            "components rather than on raw expression. Scatter plots were coloured by the ssGSEA-"
            "derived label to qualitatively assess whether the immune phenotype is captured by leading "
            "modes of transcriptomic variation."
        )
    )

    story.append(P("2.5 Unsupervised clustering", "h2"))
    story.append(
        P(
            "We benchmarked three fundamentally different clustering algorithms: k-means "
            "(k&nbsp;=&nbsp;2..6 with a silhouette sweep for model selection), agglomerative "
            "hierarchical clustering with Ward linkage (producing a full dendrogram), and density-"
            "based DBSCAN. Each cluster assignment was compared with the ssGSEA ground-truth label "
            "using Adjusted Rand Index (ARI) and Adjusted Mutual Information (AMI); the silhouette "
            "coefficient was reported as an internal validity index."
        )
    )

    story.append(P("2.6 Hypothesis testing and multiple-testing correction", "h2"))
    story.append(
        P(
            "For every gene we computed five independent statistical tests: (i) Welch&rsquo;s two-"
            "sample t-test comparing Hot vs Cold expression; (ii) one-way ANOVA across all three "
            "immune classes; (iii) the non-parametric Kruskal-Wallis test (rank-based ANOVA "
            "analogue); (iv) the pairwise non-parametric Mann-Whitney U test (Hot vs Cold); and "
            "(v) the F-test of equality of variance between Hot and Cold groups. P-values from each "
            "test were corrected for multiple comparisons using both the Benjamini-Hochberg (FDR) and "
            "the Bonferroni procedures. Clinical categorical variables (pathologic stage, vital "
            "status) were tested for association with the immune phenotype using Pearson&rsquo;s "
            "&chi;<sup>2</sup> on contingency tables. Genes satisfying BH-adjusted t-test "
            "q&nbsp;&lt;&nbsp;0.05 <i>and</i> |log<sub>2</sub>FC|&nbsp;&gt;&nbsp;1 were flagged as "
            "high-confidence differentially expressed (DE) and used for the volcano plot and the "
            "top-40 heatmap."
        )
    )

    story.append(P("2.7 Supervised classification", "h2"))
    story.append(
        P(
            "Six classifiers were evaluated in a unified scikit-learn pipeline "
            "(<font face=\"Courier\">StandardScaler &rarr; SelectKBest(ANOVA, k=500) &rarr; "
            "model</font>): k-Nearest Neighbours (k&nbsp;=&nbsp;5), linear and RBF Support Vector "
            "Machines (C&nbsp;=&nbsp;1), multinomial Logistic Regression, Random Forest (500 trees), "
            "and XGBoost (500 boosted trees with a label-encoder wrapper). Each model was evaluated "
            "with <b>stratified 5-fold cross-validation</b> on the 294 COAD samples (and "
            "independently on 109 READ samples) using accuracy, balanced accuracy and macro-F1. The "
            "top three models per cohort were re-fit on the full training set and stored; confusion "
            "matrices and one-vs-rest ROC curves were generated via cross-validated predictions."
        )
    )

    story.append(P("2.8 Survival analysis", "h2"))
    story.append(
        P(
            "Overall survival was derived from <i>days_to_death</i> (for deceased patients) and "
            "<i>days_to_last_follow_up</i> (for censored patients); vital status provided the event "
            "indicator. Non-parametric Kaplan-Meier curves were fit per immune label and the equality "
            "of survival distributions was tested with the multivariate log-rank test. A multivariate "
            "Cox proportional-hazards model (L<sub>2</sub> penalty = 0.1) was fit on the five "
            "highest-variance ssGSEA cell-type scores. Finally, Partial Least Squares (PLS) "
            "regression with three components was used to predict survival time directly from the "
            "5,000-gene expression matrix, demonstrating dimension-reduced supervised regression as "
            "covered in BT3041."
        )
    )

    story.append(P("2.9 External validation", "h2"))
    story.append(
        P(
            "The best discovery-cohort classifier (highest 5-fold CV balanced accuracy on TCGA-COAD) "
            "was frozen and applied to TCGA-READ expression after aligning feature spaces on the "
            "training feature list. Genes present in training but absent in test were zero-imputed; "
            "genes present only in test were discarded. Predictions were compared to the ssGSEA-"
            "derived labels on READ to produce an external confusion matrix, balanced accuracy and "
            "macro-F1."
        )
    )

    story.append(P("2.10 Software", "h2"))
    story.append(
        P(
            "All analyses were performed in Python 3.12 using pandas, numpy, scipy, scikit-learn, "
            "statsmodels, gseapy, umap-learn, lifelines, seaborn, xgboost, and reportlab (for this "
            "PDF). Every figure and table in this report is auto-generated by the 10-module pipeline "
            "under <font face=\"Courier\">01_*.py &hellip; 10_*.py</font>. The full source tree, "
            f"including this PDF, is publicly available at "
            f"<font color=\"#0645AD\">{REPO_URL}</font>."
        )
    )

    # ========= 3. Results ========= #
    story.append(P("3. Results", "h1"))

    # ----- 3.1 DimRed ----- #
    story.append(P("3.1 Representation learning and immune-phenotype gradient", "h2"))
    pca_coad = t["pca_coad"]
    pcs95 = int((pca_coad["cumulative"] < 0.95).sum() + 1) if len(pca_coad) else -1
    pc1_var = pca_coad["variance_ratio"].iloc[0] if len(pca_coad) else 0
    story.append(
        P(
            f"The first principal component alone captured <b>{pc1_var:.1%}</b> of total variance in "
            f"TCGA-COAD, and <b>{pcs95}</b> components were needed to explain 95% of it. Low-"
            "dimensional projections (Fig.&nbsp;3) reveal a clear gradient from Cold (blue) to Hot "
            "(red) along PC1 and along the first UMAP axis. Non-linear projections (t-SNE, UMAP) "
            "produce more compact Hot / Cold islands than linear ones, consistent with non-linear "
            "structure in the transcriptomic space; importantly, Intermediate samples sit between "
            "Hot and Cold rather than outside them, matching the tertile-split interpretation."
        )
    )

    story.extend(
        fig_col(
            FIG_DIR / "dim_reduction" / "TCGA-COAD_pca_variance.png",
            "Figure&nbsp;3a. PCA scree plot, TCGA-COAD: individual and cumulative variance explained.",
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "dim_reduction" / "TCGA-COAD_pca.png",
            "Figure&nbsp;3b. PCA projection, TCGA-COAD, coloured by immune label.",
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "dim_reduction" / "TCGA-COAD_umap.png",
            "Figure&nbsp;3c. UMAP on top-20 PCA components, TCGA-COAD.",
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "dim_reduction" / "TCGA-COAD_tsne.png",
            "Figure&nbsp;3d. t-SNE on top-20 PCA components, TCGA-COAD.",
        )
    )

    # ----- 3.2 Clustering ----- #
    story.append(P("3.2 Unsupervised clustering partially recovers the immune structure", "h2"))
    clust = t["clust_coad"].sort_values("ARI", ascending=False)
    crows = [["Method", "k", "ARI", "AMI", "Silh."]]
    for _, r in clust.iterrows():
        sil = "—" if pd.isna(r["silhouette"]) else f"{r['silhouette']:.3f}"
        crows.append(
            [
                str(r["method"]),
                str(int(r["n_clusters"])),
                f"{r['ARI']:.3f}",
                f"{r['AMI']:.3f}",
                sil,
            ]
        )
    story.extend(
        col_table(
            crows,
            col_widths_in=[0.95, 0.28, 0.50, 0.50, 0.55],
            caption="Table&nbsp;2. Clustering performance on TCGA-COAD (vs ssGSEA labels).",
        )
    )
    story.append(
        P(
            "k-means with k=3 achieves the highest Adjusted Rand Index (0.114) and Adjusted Mutual "
            "Information (0.119), confirming that the biological signal is partially — but not "
            "fully — captured by unsupervised structure alone. The modest ARI reflects the "
            "continuous nature of immune infiltration: Intermediate tumours dilute any discrete "
            "cluster separation, which is a known limitation of using tertile labels as "
            "pseudo-ground-truth."
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "clustering" / "TCGA-COAD_silhouette_sweep.png",
            "Figure&nbsp;4a. Silhouette sweep for k-means across k&nbsp;=&nbsp;2..6.",
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "clustering" / "TCGA-COAD_dendrogram.png",
            "Figure&nbsp;4b. Ward-linkage hierarchical dendrogram, TCGA-COAD.",
        )
    )

    # ----- 3.3 Hypothesis testing ----- #
    story.append(P("3.3 Thousands of genes differ between hot and cold tumours", "h2"))
    dge = t["dge_coad"]
    t_raw = int((dge["t_p"] < 0.05).sum())
    t_bh = int((dge["t_p_bh"] < 0.05).sum())
    t_bonf = int((dge["t_p_bonf"] < 0.05).sum())
    an_raw = int((dge["anova_p"] < 0.05).sum())
    an_bh = int((dge["anova_p_bh"] < 0.05).sum())
    an_bonf = int((dge["anova_p_bonf"] < 0.05).sum())
    kw_raw = int((dge["kw_p"] < 0.05).sum())
    kw_bh = int((dge["kw_p_bh"] < 0.05).sum())
    kw_bonf = int((dge["kw_p_bonf"] < 0.05).sum())
    mw_raw = int((dge["mw_p"] < 0.05).sum()) if "mw_p" in dge.columns else 0
    mw_bh = int((dge["mw_p_bh"] < 0.05).sum()) if "mw_p_bh" in dge.columns else 0
    mw_bonf = int((dge["mw_p_bonf"] < 0.05).sum()) if "mw_p_bonf" in dge.columns else 0
    fvar_raw = int((dge["Fvar_p"] < 0.05).sum())

    story.append(
        P(
            f"Of the 16,723 expressed genes in TCGA-COAD, <b>{t_bh:,}</b> were significant at the "
            f"BH 0.05 threshold by Welch t-test and <b>{t_bonf:,}</b> survived the stricter "
            f"Bonferroni 0.05 correction. Agreement with non-parametric alternatives was strong: "
            f"Mann-Whitney U flagged <b>{mw_bh:,}</b> and Kruskal-Wallis <b>{kw_bh:,}</b> at BH 0.05, "
            f"indicating that the signal is not an artefact of violated distributional assumptions. "
            f"ANOVA across all three immune groups was even more sensitive "
            f"(<b>{an_bh:,}</b> at BH 0.05). Combining BH q &lt; 0.05 with |log<sub>2</sub>FC| &gt; 1 "
            f"gave a curated list of <b>2,269</b> high-confidence DE genes."
        )
    )

    trows = [
        ["Test", "Scope", "Raw", "BH q&lt;0.05", "Bonf."],
        ["Welch t", "Hot vs Cold", f"{t_raw:,}", f"{t_bh:,}", f"{t_bonf:,}"],
        ["ANOVA", "3-group", f"{an_raw:,}", f"{an_bh:,}", f"{an_bonf:,}"],
        ["Kruskal", "3-group", f"{kw_raw:,}", f"{kw_bh:,}", f"{kw_bonf:,}"],
        ["MWU", "Hot vs Cold", f"{mw_raw:,}", f"{mw_bh:,}", f"{mw_bonf:,}"],
        ["F-var.", "Hot vs Cold", f"{fvar_raw:,}", "—", "—"],
    ]
    story.extend(
        col_table(
            trows,
            col_widths_in=[0.72, 0.80, 0.55, 0.65, 0.55],
            caption="Table&nbsp;3. Per-gene significance counts across tests (TCGA-COAD).",
        )
    )

    # Top DE genes
    story.append(P("3.3.1 Top differentially expressed genes", "h2"))
    top = t["dge_sig_coad"].sort_values("t_p_bh").head(10)
    derows = [["Gene", "μ(Hot)", "μ(Cold)", "log2FC", "BH q"]]
    for gene, row in top.iterrows():
        derows.append(
            [
                str(gene),
                f"{row['mean_Hot']:.2f}",
                f"{row['mean_Cold']:.2f}",
                f"{row['log2FC_HotVsCold']:+.2f}",
                f"{row['t_p_bh']:.1e}",
            ]
        )
    story.extend(
        col_table(
            derows,
            col_widths_in=[0.82, 0.55, 0.55, 0.55, 0.70],
            caption=(
                "Table&nbsp;4. Top-10 Hot-vs-Cold differentially expressed genes in TCGA-COAD "
                "(ranked by BH-adjusted Welch t q-value)."
            ),
        )
    )
    story.append(
        P(
            "The list is strikingly interpretable: six MHC class-II antigen-presentation genes "
            "(HLA-DPB1, HLA-DRA, HLA-DPA1, HLA-DQA1, HLA-DMB, HLA-DOA) together with CD48 (activation "
            "marker on lymphocytes), C1QB (complement), SLAMF7 (NK-cell / cytotoxic effector) and "
            "PLA2G2D (inflammatory lipid mediator). This is a textbook immune-hot signature, which "
            "strongly suggests that our ssGSEA-derived labels capture genuine biology rather than "
            "technical noise."
        )
    )

    story.extend(
        fig_wide(
            FIG_DIR / "hypothesis" / "volcano_TCGA-COAD.png",
            "Figure&nbsp;5. Volcano plot of Hot-vs-Cold differential expression in TCGA-COAD. Red: "
            "BH q&lt;0.05 and |log<sub>2</sub>FC|&nbsp;&gt;&nbsp;1 (2,269 genes).",
        )
    )
    story.extend(
        fig_wide(
            FIG_DIR / "hypothesis" / "heatmap_top40_TCGA-COAD.png",
            "Figure&nbsp;6. Expression heatmap of the 40 most-significant Hot/Cold DE genes in "
            "TCGA-COAD; rows z-scored, columns ordered by ssGSEA label.",
        )
    )

    # ----- 3.4 Classification ----- #
    story.append(P("3.4 Supervised classification achieves balanced accuracy &gt; 0.78", "h2"))
    cv = t["cv_coad"].sort_values("balanced_accuracy_mean", ascending=False)
    cvrows = [["Model", "Acc.", "Bal.Acc.", "F1"]]
    for _, r in cv.iterrows():
        cvrows.append(
            [
                str(r["model"]),
                f"{r['accuracy_mean']:.3f}",
                f"{r['balanced_accuracy_mean']:.3f}",
                f"{r['f1_macro_mean']:.3f}",
            ]
        )
    story.extend(
        col_table(
            cvrows,
            col_widths_in=[1.05, 0.55, 0.65, 0.55],
            caption="Table&nbsp;5. 5-fold stratified CV results on TCGA-COAD.",
        )
    )
    story.append(
        P(
            "Random Forest and linear SVM tied for the best balanced accuracy (<b>0.785</b>). Logistic "
            "Regression followed closely (0.752). XGBoost underperformed (0.704) in this regime, "
            "likely because aggressive ANOVA-based feature pre-selection already linearises the "
            "decision boundary and removes much of the signal that boosted trees typically exploit. "
            "kNN performed worst (0.648), reflecting the curse of dimensionality even after feature "
            "selection."
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "classification" / "TCGA-COAD_model_comparison.png",
            "Figure&nbsp;7. Ranking of the six classifiers by CV balanced accuracy on TCGA-COAD.",
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "classification" / "TCGA-COAD_RandomForest_confmat.png",
            "Figure&nbsp;8a. Confusion matrix for the best model (Random Forest, COAD, CV).",
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "classification" / "TCGA-COAD_RandomForest_roc.png",
            "Figure&nbsp;8b. One-vs-rest ROC curves for Random Forest on TCGA-COAD "
            "(AUC: Cold 0.97, Hot 0.96, Intermediate 0.84).",
        )
    )

    # ----- 3.5 External validation ----- #
    story.append(P("3.5 External validation on TCGA-READ confirms generalisation", "h2"))
    ext = t["ext"].iloc[0] if len(t["ext"]) else None
    if ext is not None:
        story.append(
            P(
                f"The frozen Random Forest model trained on TCGA-COAD was applied to the fully "
                f"independent TCGA-READ cohort (n=<b>{int(ext['n'])}</b>), which was never seen "
                f"during training, hyper-parameter selection, or feature filtering. It reached "
                f"accuracy&nbsp;=&nbsp;<b>{float(ext['accuracy']):.3f}</b>, balanced "
                f"accuracy&nbsp;=&nbsp;<b>{float(ext['balanced_accuracy']):.3f}</b>, and "
                f"macro-F1&nbsp;=&nbsp;<b>{float(ext['f1_macro']):.3f}</b>. The gap between "
                f"discovery (CV) and external performance is only ~4 percentage points &mdash; a "
                f"small, expected domain-shift effect that strongly supports the transferability of "
                f"our transcriptomic immune signature."
            )
        )
    story.extend(
        fig_col(
            FIG_DIR / "validation" / "validation_confmat_TCGA-READ.png",
            "Figure&nbsp;9. External validation confusion matrix (COAD-trained RF on READ). "
            "Most errors are Hot&harr;Intermediate confusions, consistent with a continuous "
            "immune gradient rather than hard categorical boundaries.",
        )
    )

    # ----- 3.6 Survival ----- #
    story.append(P("3.6 Survival analysis links the immune phenotype to outcome", "h2"))
    lrp = float(t["logrank_coad"]["p_value"].iloc[0]) if len(t["logrank_coad"]) else None
    story.append(
        P(
            f"Kaplan-Meier analysis stratified by ssGSEA immune label "
            f"(Fig.&nbsp;10) shows that Cold tumours trend towards shorter overall survival than "
            f"Hot tumours in TCGA-COAD; the multivariate log-rank test yielded "
            f"p&nbsp;=&nbsp;<b>{lrp:.3f}</b>. With roughly 90 death events the trend does not cross "
            f"the p&nbsp;&lt;&nbsp;0.05 threshold, but the direction is biologically expected and "
            f"consistent with the hot-tumour-better-prognosis literature. A five-covariate Cox "
            f"proportional-hazards model on the ssGSEA cell-type scores corroborates the pattern "
            f"(Table&nbsp;6): neutrophil infiltration carries a protective hazard ratio "
            f"(HR&nbsp;=&nbsp;0.32) while macrophage infiltration trends adverse (HR&nbsp;=&nbsp;1.52), "
            f"although the wide confidence intervals reflect the modest event count."
        )
    )
    cox = t["cox_coad"]
    if len(cox):
        crows = [["Covariate", "coef", "HR", "SE", "z", "p"]]
        for _, r in cox.iterrows():
            crows.append(
                [
                    str(r["covariate"]),
                    f"{r['coef']:+.2f}",
                    f"{r['exp(coef)']:.2f}",
                    f"{r['se(coef)']:.2f}",
                    f"{r['z']:+.2f}",
                    f"{r['p']:.2f}",
                ]
            )
        story.extend(
            col_table(
                crows,
                col_widths_in=[0.90, 0.48, 0.48, 0.48, 0.38, 0.38],
                caption="Table&nbsp;6. Cox proportional-hazards regression (penalizer = 0.1), TCGA-COAD.",
            )
        )
    story.extend(
        fig_col(
            FIG_DIR / "survival" / "TCGA-COAD_kaplan_meier.png",
            "Figure&nbsp;10a. Kaplan-Meier survival curves by immune phenotype, TCGA-COAD.",
        )
    )
    story.extend(
        fig_col(
            FIG_DIR / "survival" / "TCGA-COAD_pls_survival.png",
            "Figure&nbsp;10b. PLS regression of expression onto survival time "
            "(R<sup>2</sup>&nbsp;=&nbsp;0.37 on training).",
        )
    )

    # ========= 4. Discussion ========= #
    story.append(P("4. Discussion", "h1"))
    story.append(
        P(
            "Our results converge on three main conclusions. <b>First</b>, the tumour immune "
            "phenotype is strongly encoded in the bulk transcriptome: both unsupervised (PCA, t-SNE, "
            "UMAP) and supervised (Random Forest, SVM) approaches recover it with high confidence, "
            "and the top differentially expressed genes form a biologically coherent MHC-II + "
            "cytotoxic effector signature. <b>Second</b>, the multiple-testing framework &mdash; "
            "combining parametric and non-parametric tests with both BH and Bonferroni correction "
            "&mdash; agrees on the identity of the most differentially expressed genes, suggesting "
            "robustness to distributional assumptions. <b>Third</b>, external validation on a "
            "completely independent cohort (TCGA-READ) retains ~74% balanced accuracy, which is the "
            "most important evidence that we are not overfitting the discovery cohort."
        )
    )
    story.append(
        P(
            "Several limitations should be noted. The ssGSEA-derived labels, although biologically "
            "grounded, are themselves computationally inferred rather than obtained from ground-"
            "truth immunohistochemistry; therefore accuracies describe agreement with ssGSEA, not "
            "with pathology. Both cohorts are from the same TCGA platform, which controls for "
            "technical batch effects but does not test cross-platform generalisation &mdash; a "
            "natural follow-up would include a microarray cohort such as GSE39582. The survival "
            "signal is directionally consistent but under-powered at n&nbsp;&asymp;&nbsp;90 events. "
            "Finally, XGBoost under-performed linear SVM and logistic regression, suggesting that "
            "after ANOVA-based feature selection the decision boundary becomes approximately "
            "linear; boosted trees may excel more in the full 5000-gene setting without pre-"
            "selection."
        )
    )

    story.append(P("4.1 Coverage of the BT3041 syllabus", "h2"))
    story.append(bullet("Distributions and normality tests (Shapiro, QQ) &mdash; <b>§2.2, Fig.&nbsp;1</b>"))
    story.append(bullet("Parametric / non-parametric hypothesis testing (t, F-var, ANOVA, Kruskal-Wallis, Mann-Whitney U, χ²) &mdash; <b>§2.6, §3.3, Tables&nbsp;3–4</b>"))
    story.append(bullet("Multiple-testing correction (Bonferroni, BH) &mdash; <b>§3.3, Table&nbsp;3</b>"))
    story.append(bullet("Dimensionality reduction (PCA, ICA, MDS, t-SNE, UMAP) &mdash; <b>§2.4, §3.1, Fig.&nbsp;3</b>"))
    story.append(bullet("Clustering (k-means, Ward hierarchical, DBSCAN) with ARI / AMI / silhouette &mdash; <b>§2.5, §3.2, Table&nbsp;2, Fig.&nbsp;4</b>"))
    story.append(bullet("Supervised classification (kNN, linear / RBF SVM, logistic, random forest, XGBoost) &mdash; <b>§2.7, §3.4, Table&nbsp;5, Figs.&nbsp;7–8</b>"))
    story.append(bullet("Regression (multinomial logistic, PLS) &mdash; <b>§2.7, §2.8, Fig.&nbsp;10b</b>"))
    story.append(bullet("Survival analysis (Kaplan-Meier, log-rank, Cox) &mdash; <b>§2.8, §3.6, Table&nbsp;6, Fig.&nbsp;10</b>"))
    story.append(bullet("External validation on a held-out cohort &mdash; <b>§2.9, §3.5, Fig.&nbsp;9</b>"))

    # ========= 5. Conclusion ========= #
    story.append(P("5. Conclusion", "h1"))
    story.append(
        P(
            "We have built, executed and externally validated a reproducible statistical and "
            "machine-learning pipeline that classifies colorectal tumours into Hot, Cold and "
            "Intermediate immune phenotypes from bulk RNA-seq. In a 403-patient TCGA-COAD / "
            "TCGA-READ setting the pipeline achieves 78.5% cross-validated balanced accuracy and "
            "74.4% balanced accuracy on a fully independent external cohort, uncovers more than "
            "two thousand biologically interpretable differentially expressed genes, and recovers "
            "the expected immune-infiltration gradient in both unsupervised low-dimensional "
            "projections and survival analysis. The pipeline covers every topic of the BT3041 "
            "syllabus and is packaged so that any team member or reviewer can regenerate every "
            "figure and table of this report by running a single command."
        )
    )

    story.append(P("Code, data and reproducibility", "h1"))
    story.append(
        Paragraph(
            f"GitHub repository: <a href='{REPO_URL}' color='#0645AD'>{REPO_URL}</a>",
            STYLES["link"],
        )
    )
    story.append(
        P(
            "The repository contains: (i) ten Python scripts covering data acquisition &rarr; final "
            "figures; (ii) <font face=\"Courier\">report/generate_report_pdf.py</font> — the exact "
            "script that produced this PDF; (iii) <font face=\"Courier\">requirements.txt</font> "
            "pinning reproducible dependencies; and (iv) a README with end-to-end instructions."
        )
    )

    story.append(P("References", "h1"))
    refs = [
        "Yoshihara, K. et al. Inferring tumour purity and stromal and immune cell admixture from expression data. <i>Nature Communications</i>, 2013.",
        "Bindea, G. et al. Spatiotemporal dynamics of intratumoral immune cells reveal the immune landscape in human cancer. <i>Immunity</i>, 2013.",
        "Charoentong, P. et al. Pan-cancer immunogenomic analyses reveal genotype-immunophenotype relationships and predictors of response to checkpoint blockade. <i>Cell Reports</i>, 2017.",
        "Benjamini, Y. &amp; Hochberg, Y. Controlling the False Discovery Rate. <i>JRSS-B</i>, 1995.",
        "McInnes, L., Healy, J. &amp; Melville, J. UMAP: Uniform Manifold Approximation and Projection. <i>J. Open Source Software</i>, 2018.",
        "Davidson-Pilon, C. lifelines: survival analysis in Python. <i>J. Open Source Software</i>, 2019.",
        "Pedregosa, F. et al. Scikit-learn: Machine Learning in Python. <i>JMLR</i>, 2011.",
    ]
    for i, r in enumerate(refs, 1):
        story.append(Paragraph(f"[{i}] {r}", STYLES["refs"]))

    return story


def main() -> None:
    t = load_tables()
    doc = build_doc()
    doc.build(build_story(t))
    print(f"Report written: {OUT_PDF} ({OUT_PDF.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
