"""
Generate a styled BT3041 research report PDF from pipeline outputs.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    NextPageTemplate,
    PageBreak,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageTemplate,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUT_PDF = PROJECT_ROOT / "report" / "BT3041_Cold_Hot_Tumour_Report.pdf"


def load_tables() -> dict[str, pd.DataFrame]:
    tables_dir = PROJECT_ROOT / "outputs" / "tables"
    return {
        "final_summary": pd.read_csv(tables_dir / "final_summary.csv"),
        "cv_coad": pd.read_csv(tables_dir / "cv_results_TCGA-COAD.csv"),
        "ext_val": pd.read_csv(tables_dir / "external_validation_summary_TCGA-READ.csv"),
        "dge_coad": pd.read_csv(tables_dir / "dge_TCGA-COAD.csv"),
        "logrank_coad": pd.read_csv(tables_dir / "logrank_TCGA-COAD.csv"),
    }


def _page_number(canvas, doc) -> None:
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(A4[0] - 25, 15, f"{doc.page}")
    canvas.restoreState()


def make_doc() -> BaseDocTemplate:
    doc = BaseDocTemplate(
        str(OUT_PDF),
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=30,
        title="Cold vs Hot Tumour Classification Report",
        author="BT3041 Project Team",
    )

    # First page: single column (title + abstract).
    frame_first = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height,
        id="first",
    )

    # Body: two-column style to mimic conference report format.
    gap = 18
    col_w = (doc.width - gap) / 2
    frame_col1 = Frame(doc.leftMargin, doc.bottomMargin, col_w, doc.height, id="col1")
    frame_col2 = Frame(
        doc.leftMargin + col_w + gap,
        doc.bottomMargin,
        col_w,
        doc.height,
        id="col2",
    )

    doc.addPageTemplates(
        [
            PageTemplate(id="FirstPage", frames=[frame_first], onPage=_page_number),
            PageTemplate(id="TwoCol", frames=[frame_col1, frame_col2], onPage=_page_number),
        ]
    )
    return doc


def fig(path: Path, width: float, caption: str, caption_style: ParagraphStyle):
    if not path.exists():
        return [Paragraph(f"<i>Missing figure:</i> {path}", caption_style), Spacer(1, 0.1 * inch)]
    img = Image(str(path))
    img.drawWidth = width
    img.drawHeight = width * (img.imageHeight / img.imageWidth)
    return [img, Spacer(1, 0.05 * inch), Paragraph(caption, caption_style), Spacer(1, 0.15 * inch)]


def make_report() -> None:
    t = load_tables()
    doc = make_doc()
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleMain",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#0c2d57"),
        spaceAfter=8,
    )
    subtitle_style = ParagraphStyle(
        "Subtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10.5,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#333333"),
        spaceAfter=8,
    )
    h1 = ParagraphStyle(
        "H1",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        textColor=colors.HexColor("#123b64"),
        spaceBefore=6,
        spaceAfter=5,
    )
    h2 = ParagraphStyle(
        "H2",
        parent=styles["Heading3"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        textColor=colors.HexColor("#123b64"),
        spaceBefore=4,
        spaceAfter=3,
    )
    body = ParagraphStyle(
        "Body",
        parent=styles["BodyText"],
        fontName="Times-Roman",
        fontSize=9.5,
        leading=13,
        alignment=TA_JUSTIFY,
        spaceAfter=5,
    )
    cap = ParagraphStyle(
        "Caption",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        textColor=colors.HexColor("#444444"),
        alignment=TA_CENTER,
    )

    cohort_df = t["final_summary"].copy()
    cv = t["cv_coad"].copy()
    cv = cv[["model", "balanced_accuracy_mean", "f1_macro_mean"]]
    cv.columns = ["Model", "Balanced Accuracy", "Macro F1"]
    cv["Balanced Accuracy"] = cv["Balanced Accuracy"].map(lambda x: f"{x:.3f}")
    cv["Macro F1"] = cv["Macro F1"].map(lambda x: f"{x:.3f}")

    dge = t["dge_coad"]
    sig_hotcold = int((dge["t_p_bh"] < 0.05).sum())
    sig_anova = int((dge["anova_p_bh"] < 0.05).sum())
    sig_kw = int((dge["kw_p_bh"] < 0.05).sum())
    sig_mw = int((dge["mw_p_bh"] < 0.05).sum()) if "mw_p_bh" in dge.columns else None
    logrank_p = float(t["logrank_coad"]["p_value"].iloc[0])
    ext = t["ext_val"].iloc[0]

    story = []

    # Title block.
    story.append(Paragraph("BT3041 Analysis and Interpretation of Biological Project Data: Report", subtitle_style))
    story.append(Paragraph("Cold vs Hot Tumour Classification from Bulk RNA-seq Using a Multi-Statistical Pipeline", title_style))
    story.append(
        Paragraph(
            "Aryan (Team Lead) with BT3041 Project Team",
            subtitle_style,
        )
    )
    story.append(
        Paragraph(
            f"Date of submission: {date.today().strftime('%B %d, %Y')}",
            subtitle_style,
        )
    )
    story.append(Spacer(1, 0.1 * inch))

    abstract_text = (
        "This term-paper project develops and validates a complete end-to-end framework to classify "
        "colorectal tumours into immune-hot, immune-cold, and intermediate phenotypes using transcriptomic "
        "profiles from TCGA. The final dataset includes 403 patients across two independent cohorts "
        "(TCGA-COAD n=294, TCGA-READ n=109). The pipeline integrates preprocessing, ssGSEA/ESTIMATE-based "
        "immune scoring, dimensionality reduction, unsupervised clustering, hypothesis testing (Welch t-test, "
        "ANOVA, Kruskal-Wallis, Mann-Whitney U, chi-square), machine learning classification, survival analysis, "
        "and external validation. Cross-validated model performance in the discovery cohort reached balanced "
        "accuracy of 0.785, while external validation achieved balanced accuracy of 0.744. Differential expression "
        "and survival analyses further support clinically meaningful immune stratification."
    )
    abstract_box = Table(
        [[Paragraph("<b>Abstract</b><br/>" + abstract_text, body)]],
        colWidths=[doc.width - 4],
    )
    abstract_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#e9f0f8")),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#9fb3c8")),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(abstract_box)
    story.append(Spacer(1, 0.12 * inch))
    story.append(NextPageTemplate("TwoCol"))
    story.append(PageBreak())

    # Body sections.
    story.append(Paragraph("1. Introduction", h1))
    story.append(
        Paragraph(
            "Tumour immune context strongly influences treatment response in colorectal cancer. "
            "Immune-hot tumours generally show higher infiltration by cytotoxic and helper immune cells and "
            "often exhibit improved response to checkpoint blockade therapies. Immune-cold tumours, in contrast, "
            "show poor immune infiltration and weaker immunotherapy benefit. This project addresses a practical "
            "clinical informatics question: can we predict tumour immune phenotype from routine bulk RNA-seq data "
            "using a statistically rigorous and reproducible analysis stack?",
            body,
        )
    )

    story.append(Paragraph("2. Materials and Methods", h1))
    story.append(Paragraph("2.1 Cohorts and Data Curation", h2))
    story.append(
        Paragraph(
            "Data were downloaded from GDC as STAR-count RNA-seq files with linked clinical metadata. "
            "To maximize robustness under unstable network conditions, downloads were batched and retried automatically. "
            "TCGA-COAD was used as discovery/training cohort and TCGA-READ as independent validation cohort.",
            body,
        )
    )

    cohort_table_data = [["Cohort", "Samples", "Hot", "Cold", "Intermediate"]]
    for _, r in cohort_df.iterrows():
        cohort_table_data.append(
            [r["cohort"], int(r["n_samples"]), int(r["Hot"]), int(r["Cold"]), int(r["Intermediate"])]
        )
    cohort_tbl = Table(cohort_table_data, colWidths=[1.05 * inch, 0.7 * inch, 0.45 * inch, 0.45 * inch, 0.75 * inch])
    cohort_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dbe7f3")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ]
        )
    )
    story.append(cohort_tbl)
    story.append(Spacer(1, 0.05 * inch))
    story.append(Paragraph("Table 1. Final cohort composition after preprocessing and immune-label assignment.", cap))
    story.append(Spacer(1, 0.12 * inch))

    story.append(Paragraph("2.2 Preprocessing and Harmonization", h2))
    story.append(
        Paragraph(
            "Raw counts were converted with log2(CPM + 1), and genes were filtered at >=1 CPM in >=20% of samples. "
            "Top-variance genes (K=5000) were used for machine learning. Gene-symbol intersection harmonization "
            "between cohorts yielded 16,453 common genes, ensuring consistent feature spaces for downstream analyses.",
            body,
        )
    )

    story.append(Paragraph("2.3 Statistical and Machine Learning Pipeline", h2))
    story.append(
        Paragraph(
            "Immune labels were generated by ssGSEA with ESTIMATE/Bindea signatures and quantile-based stratification "
            "(top third: Hot, bottom third: Cold, middle: Intermediate). Dimensionality reduction included PCA, ICA, "
            "MDS, t-SNE, and UMAP. Unsupervised discovery used k-means, hierarchical clustering, and DBSCAN. "
            "Hypothesis testing included Welch t-test, one-way ANOVA, Kruskal-Wallis, Mann-Whitney U, and chi-square, "
            "with Benjamini-Hochberg and Bonferroni correction. Supervised classification models included kNN, "
            "SVM (linear/RBF), logistic regression, random forest, and XGBoost under stratified 5-fold CV.",
            body,
        )
    )

    story.append(Paragraph("3. Results", h1))
    story.append(Paragraph("3.1 Representation Learning and Immune Separation", h2))
    story.extend(
        fig(
            PROJECT_ROOT / "outputs" / "figures" / "report" / "Fig2_embedding_panel.png",
            width=2.55 * inch,
            caption="Figure 1. PCA/t-SNE/UMAP composite for TCGA-COAD showing partial immune-phenotype separation.",
            caption_style=cap,
        )
    )

    story.append(Paragraph("3.2 Differential Expression and Hypothesis Testing", h2))
    story.append(
        Paragraph(
            f"In TCGA-COAD (16,723 genes tested), significant genes after BH correction were: "
            f"t-test = {sig_hotcold}, ANOVA = {sig_anova}, Kruskal-Wallis = {sig_kw}"
            + (f", Mann-Whitney U = {sig_mw}" if sig_mw is not None else "")
            + ". Applying combined BH + |log2FC| > 1 criteria yielded 2,269 high-confidence DE genes.",
            body,
        )
    )
    story.extend(
        fig(
            PROJECT_ROOT / "outputs" / "figures" / "hypothesis" / "volcano_TCGA-COAD.png",
            width=2.55 * inch,
            caption="Figure 2. Volcano plot (Hot vs Cold, TCGA-COAD) with BH-adjusted significance thresholds.",
            caption_style=cap,
        )
    )

    story.append(Paragraph("3.3 Classification Performance", h2))
    model_table_data = [["Model", "Balanced Acc.", "Macro F1"]]
    for _, row in cv.iterrows():
        model_table_data.append([row["Model"], row["Balanced Accuracy"], row["Macro F1"]])
    model_tbl = Table(model_table_data, colWidths=[1.0 * inch, 0.75 * inch, 0.6 * inch])
    model_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dbe7f3")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ]
        )
    )
    story.append(model_tbl)
    story.append(Spacer(1, 0.05 * inch))
    story.append(Paragraph("Table 2. Five-fold cross-validation performance in TCGA-COAD.", cap))
    story.append(Spacer(1, 0.1 * inch))
    story.extend(
        fig(
            PROJECT_ROOT / "outputs" / "figures" / "classification" / "TCGA-COAD_model_comparison.png",
            width=2.55 * inch,
            caption="Figure 3. Model comparison in discovery cohort (balanced accuracy, 5-fold CV).",
            caption_style=cap,
        )
    )

    story.append(Paragraph("3.4 External Validation and Survival", h2))
    story.append(
        Paragraph(
            f"Using the best discovery model (RandomForest), external testing on TCGA-READ reached "
            f"accuracy = {float(ext['accuracy']):.3f}, balanced accuracy = {float(ext['balanced_accuracy']):.3f}, "
            f"macro-F1 = {float(ext['f1_macro']):.3f} (n = {int(ext['n'])}). "
            f"Kaplan-Meier analysis in TCGA-COAD gave log-rank p = {logrank_p:.3f}.",
            body,
        )
    )
    story.extend(
        fig(
            PROJECT_ROOT / "outputs" / "figures" / "validation" / "validation_confmat_TCGA-READ.png",
            width=2.55 * inch,
            caption="Figure 4. External validation confusion matrix (COAD-trained model evaluated on READ).",
            caption_style=cap,
        )
    )
    story.extend(
        fig(
            PROJECT_ROOT / "outputs" / "figures" / "survival" / "TCGA-COAD_kaplan_meier.png",
            width=2.55 * inch,
            caption="Figure 5. Kaplan-Meier survival curves for Hot/Cold/Intermediate in TCGA-COAD.",
            caption_style=cap,
        )
    )

    story.append(Paragraph("4. Discussion", h1))
    story.append(
        Paragraph(
            "The analysis confirms that a compact two-cohort design (403 total patients) is sufficient for "
            "a rigorous BT3041-grade translational bioinformatics study. The strongest performance was obtained "
            "by Random Forest and linear SVM, while XGBoost did not outperform simpler baselines in this feature "
            "regime. Statistical testing consistently identified robust immune-associated expression differences, "
            "supporting biological interpretability in addition to predictive performance.",
            body,
        )
    )

    story.append(Paragraph("5. Conclusion", h1))
    story.append(
        Paragraph(
            "This project delivers a full reproducible pipeline for immune-phenotype classification in colorectal "
            "tumours, covering all major statistical and machine-learning topics in the BT3041 syllabus. "
            "Future extension can include a cross-platform GEO cohort for stricter external generalization testing.",
            body,
        )
    )

    story.append(Paragraph("Code and Reproducibility", h1))
    story.append(
        Paragraph(
            "GitHub repository: <u><font color='blue'>https://github.com/Aryan7865/cold-hot-tumour-classification</font></u>",
            body,
        )
    )
    story.append(
        Paragraph(
            "All figures and tables in this report are auto-generated from scripts 01–10 and stored under outputs/.",
            body,
        )
    )

    doc.build(story)
    print(f"Report written: {OUT_PDF}")


if __name__ == "__main__":
    make_report()
