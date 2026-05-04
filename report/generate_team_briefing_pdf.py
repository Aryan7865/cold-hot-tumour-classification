#!/usr/bin/env python3
"""
Build TEAM_PROJECT_FULL_SUMMARY.pdf from TEAM_PROJECT_FULL_SUMMARY.md (narrative)
plus auto-generated metric tables from outputs/tables/*.csv.

Uses ReportLab only (no pandoc). Run from repo root:
  python report/generate_team_briefing_pdf.py
"""

from __future__ import annotations

import html
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MD_PATH = ROOT / "TEAM_PROJECT_FULL_SUMMARY.md"
OUT_PDF = ROOT / "TEAM_PROJECT_FULL_SUMMARY.pdf"

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
import config  # noqa: E402

import pandas as pd  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import inch  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

TBL = config.TABLES_DIR
PAGE_W, PAGE_H = A4
M = 0.65 * inch
CONTENT_W = PAGE_W - 2 * M


def load_metrics() -> dict:
    out: dict = {}
    def _read(name: str) -> pd.DataFrame:
        p = TBL / name
        return pd.read_csv(p) if p.exists() else pd.DataFrame()

    summary = _read("final_summary.csv")
    if len(summary):
        out["n_coad"] = int(summary.loc[summary["cohort"] == "TCGA-COAD", "n_samples"].iloc[0])
        out["n_read"] = int(summary.loc[summary["cohort"] == "TCGA-READ", "n_samples"].iloc[0])
        out["n_total"] = int(summary["n_samples"].sum())
    else:
        out["n_coad"] = out["n_read"] = out["n_total"] = 0

    cv = _read("cv_results_TCGA-COAD.csv")
    if len(cv):
        cv = cv.sort_values("balanced_accuracy_mean", ascending=False)
        out["cv_rows"] = cv
        out["best_model"] = str(cv.iloc[0]["model"])
        out["cv_bal"] = float(cv.iloc[0]["balanced_accuracy_mean"])
    else:
        out["cv_rows"] = pd.DataFrame()
        out["best_model"] = "—"
        out["cv_bal"] = 0.0

    ext = _read("external_validation_summary_TCGA-READ.csv")
    if len(ext):
        e = ext.iloc[0]
        out["ext_bal"] = float(e["balanced_accuracy"])
        out["ext_acc"] = float(e["accuracy"])
        out["ext_f1"] = float(e["f1_macro"])
        out["ext_n"] = int(e["n"])
    else:
        out["ext_bal"] = out["ext_acc"] = out["ext_f1"] = 0.0
        out["ext_n"] = 0

    lr = _read("logrank_TCGA-COAD.csv")
    out["logrank_p"] = float(lr["p_value"].iloc[0]) if len(lr) else float("nan")

    cl = _read("clustering_metrics_TCGA-COAD.csv")
    km = cl[cl["method"] == "kmeans_k3"]
    if len(km):
        out["ari"] = float(km.iloc[0]["ARI"])
        out["ami"] = float(km.iloc[0]["AMI"])
    else:
        out["ari"] = out["ami"] = 0.0

    dge = _read("dge_TCGA-COAD.csv")
    if len(dge) and "t_p_bh" in dge.columns:
        out["n_genes"] = len(dge)
        out["n_sig_t"] = int((dge["t_p_bh"] < 0.05).sum())
        out["n_sig_anova"] = int((dge["anova_p_bh"] < 0.05).sum()) if "anova_p_bh" in dge.columns else 0
        out["n_sig_kw"] = int((dge["kw_p_bh"] < 0.05).sum()) if "kw_p_bh" in dge.columns else 0
        out["n_sig_mw"] = int((dge["mw_p_bh"] < 0.05).sum()) if "mw_p_bh" in dge.columns else 0
    else:
        out["n_genes"] = out["n_sig_t"] = out["n_sig_anova"] = out["n_sig_kw"] = out["n_sig_mw"] = 0

    sigp = ROOT / "outputs" / "tables" / "dge_sig_TCGA-COAD.csv"
    out["n_de_conf"] = max(0, sum(1 for _ in open(sigp, encoding="utf-8")) - 1) if sigp.exists() else 0

    pca = _read("pca_variance_TCGA-COAD.csv")
    if len(pca):
        out["pc1"] = float(pca["variance_ratio"].iloc[0])
        out["pc2"] = float(pca["variance_ratio"].iloc[1]) if len(pca) > 1 else 0.0
    else:
        out["pc1"] = out["pc2"] = 0.0
    return out


def _styles():
    base = getSampleStyleSheet()
    title = ParagraphStyle(
        "t", parent=base["Title"], fontName="Helvetica-Bold", fontSize=16,
        textColor=colors.HexColor("#0c2d57"), alignment=TA_CENTER, spaceAfter=10,
    )
    title_md = ParagraphStyle(
        "tm", parent=base["Title"], fontName="Helvetica-Bold", fontSize=14,
        textColor=colors.HexColor("#0c2d57"), alignment=TA_CENTER, spaceAfter=12,
    )
    h1 = ParagraphStyle(
        "h1", parent=base["Heading1"], fontName="Helvetica-Bold", fontSize=13,
        textColor=colors.HexColor("#123b64"), spaceBefore=14, spaceAfter=6,
    )
    h2 = ParagraphStyle(
        "h2", parent=base["Heading2"], fontName="Helvetica-Bold", fontSize=11.5,
        textColor=colors.HexColor("#1e3a5f"), spaceBefore=10, spaceAfter=4,
    )
    body = ParagraphStyle(
        "b", parent=base["BodyText"], fontName="Times-Roman", fontSize=10.5,
        leading=14, alignment=TA_JUSTIFY, spaceAfter=6,
    )
    small = ParagraphStyle(
        "s", parent=body, fontSize=9, leading=12, textColor=colors.HexColor("#444444"),
    )
    bullet = ParagraphStyle(
        "bl", parent=body, leftIndent=14, bulletIndent=4, firstLineIndent=0,
    )
    return title, title_md, h1, h2, body, small, bullet


def _inline_to_xml(s: str) -> str:
    """Convert a subset of Markdown inline to ReportLab XML."""
    s = s.replace("\t", " ")
    # Links [text](url)
    def link_repl(m):
        t, u = m.group(1), m.group(2)
        return f'<a href="{html.escape(u, quote=True)}" color="blue">{html.escape(t)}</a>'

    s = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", link_repl, s)
    # Split **bold**
    parts: list[tuple[str, str]] = []
    pos = 0
    while True:
        a = s.find("**", pos)
        if a == -1:
            parts.append(("t", s[pos:]))
            break
        parts.append(("t", s[pos:a]))
        b = s.find("**", a + 2)
        if b == -1:
            parts.append(("t", s[a:]))
            break
        parts.append(("b", s[a + 2 : b]))
        pos = b + 2
    out = []
    for typ, chunk in parts:
        if typ == "b":
            out.append("<b>" + html.escape(chunk) + "</b>")
        else:
            chunk = re.sub(r"`([^`]+)`", r'<font name="Courier">\1</font>', chunk)
            out.append(html.escape(chunk).replace("\n", "<br/>"))
    return "".join(out)


def _para(text: str, style) -> Paragraph:
    return Paragraph(_inline_to_xml(text.strip()), style)


def _metric_front_matter(metrics: dict, title_style, h1, body, small) -> list:
    story = [
        Paragraph("TEAM PROJECT BRIEFING", title_style),
        Paragraph(
            "Cold vs Hot Tumour Classification — BT3041 Term Project<br/>"
            "<i>Auto-generated metrics snapshot + narrative from TEAM_PROJECT_FULL_SUMMARY.md</i>",
            small,
        ),
        Spacer(1, 12),
        Paragraph("Key numbers (from <font name='Courier'>outputs/tables/</font> at build time)", h1),
    ]
    rows = [
        ["Quantity", "Value"],
        ["TCGA-COAD samples (discovery)", str(metrics["n_coad"])],
        ["TCGA-READ samples (external test)", str(metrics["n_read"])],
        ["Total samples", str(metrics["n_total"])],
        ["Genes in per-gene DE table (COAD)", f"{metrics['n_genes']:,}"],
        ["BH q&lt;0.05 (Welch t)", f"{metrics['n_sig_t']:,}"],
        ["BH q&lt;0.05 (ANOVA)", f"{metrics['n_sig_anova']:,}"],
        ["BH q&lt;0.05 (Kruskal–Wallis)", f"{metrics['n_sig_kw']:,}"],
        ["BH q&lt;0.05 (Mann–Whitney)", f"{metrics['n_sig_mw']:,}"],
        ["High-confidence DE (BH + |log₂FC|&gt;1)", f"{metrics['n_de_conf']:,}"],
        ["PC1 / PC2 variance (COAD)", f"{metrics['pc1']:.1%} / {metrics['pc2']:.1%}"],
        ["k-means k=3 vs labels (ARI / AMI)", f"{metrics['ari']:.3f} / {metrics['ami']:.3f}"],
        ["Best CV model (balanced acc.)", f"{metrics['best_model']} ({metrics['cv_bal']:.3f})"],
        ["External READ balanced accuracy", f"{metrics['ext_bal']:.3f} (n={metrics['ext_n']})"],
        ["Log-rank p (COAD, 3 groups)", f"{metrics['logrank_p']:.4f}"],
    ]
    t = Table(rows, colWidths=[CONTENT_W * 0.55, CONTENT_W * 0.45])
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f8")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(t)
    story.append(Spacer(1, 14))
    if len(metrics["cv_rows"]):
        story.append(Paragraph("Cross-validation leaderboard (TCGA-COAD)", h1))
        crow = [["Model", "Bal.Acc.", "Acc.", "Macro-F1"]]
        for _, r in metrics["cv_rows"].iterrows():
            crow.append(
                [
                    str(r["model"]),
                    f"{float(r['balanced_accuracy_mean']):.3f}",
                    f"{float(r['accuracy_mean']):.3f}",
                    f"{float(r['f1_macro_mean']):.3f}",
                ]
            )
        ct = Table(crow, colWidths=[CONTENT_W * 0.38, CONTENT_W * 0.2, CONTENT_W * 0.2, CONTENT_W * 0.22])
        ct.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f8")),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
                    ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(ct)
    story.append(PageBreak())
    return story


def _parse_md_to_story(lines: list[str], title_c, h1, h2, body, bullet) -> list:
    story: list = []
    i = 0
    n = len(lines)

    def flush_para(buf: list[str]):
        if not buf:
            return
        text = " ".join(x.strip() for x in buf if x.strip())
        if text:
            story.append(_para(text, body))

    while i < n:
        raw = lines[i]
        line = raw.rstrip("\n")
        stripped = line.strip()

        if stripped == "---":
            story.append(Spacer(1, 8))
            i += 1
            continue

        if stripped.startswith("# ") and not stripped.startswith("##"):
            ttl = stripped[2:].strip()
            story.append(Paragraph(_inline_to_xml(ttl), title_c))
            i += 1
            continue

        if stripped.startswith("## "):
            story.append(_para(stripped[3:].strip(), h1))
            i += 1
            continue

        if stripped.startswith("### "):
            story.append(_para(stripped[4:].strip(), h2))
            i += 1
            continue

        if stripped.startswith("|") and stripped.endswith("|"):
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                rowline = lines[i].strip()
                if re.match(r"^\|\s*-+", rowline):
                    i += 1
                    continue
                cells = [re.sub(r"\*\*(.+?)\*\*", r"\1", c.strip()) for c in rowline.split("|")[1:-1]]
                rows.append(cells)
                i += 1
            if rows:
                ncol = max(len(r) for r in rows)
                for r in rows:
                    while len(r) < ncol:
                        r.append("")
                cw = CONTENT_W / ncol
                tbl = Table(rows, colWidths=[cw] * ncol)
                tbl.setStyle(
                    TableStyle(
                        [
                            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                            ("FONTSIZE", (0, 0), (-1, -1), 7.5 if ncol > 5 else 8.5),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f8")),
                            ("GRID", (0, 0), (-1, -1), 0.2, colors.HexColor("#cbd5e1")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 3),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                        ]
                    )
                )
                story.append(Spacer(1, 4))
                story.append(tbl)
                story.append(Spacer(1, 6))
            continue

        if stripped.startswith("- ") or stripped.startswith("* "):
            items = []
            while i < n:
                s2 = lines[i].strip()
                if not (s2.startswith("- ") or s2.startswith("* ")):
                    break
                items.append(ListItem(_para(s2[2:].strip(), body), leftIndent=12))
                i += 1
            if items:
                story.append(ListFlowable(items, bulletType="bullet", start="•"))
            continue

        if not stripped:
            i += 1
            continue

        buf = []
        while i < n:
            s2 = lines[i].strip()
            if not s2 or s2.startswith("#") or s2.startswith("|") or s2.startswith("- ") or s2.startswith("* ") or s2 == "---":
                break
            buf.append(lines[i])
            i += 1
        flush_para(buf)
    return story


def build() -> None:
    if not MD_PATH.exists():
        raise SystemExit(f"Missing {MD_PATH}")
    metrics = load_metrics()
    title_style, title_md, h1, h2, body, small, bullet = _styles()

    doc = SimpleDocTemplate(
        str(OUT_PDF),
        pagesize=A4,
        leftMargin=M,
        rightMargin=M,
        topMargin=M,
        bottomMargin=M,
        title="Team briefing — BT3041 Cold vs Hot Tumours",
    )
    story = _metric_front_matter(metrics, title_style, h1, body, small)

    md_text = MD_PATH.read_text(encoding="utf-8")
    lines = md_text.splitlines()
    story.extend(_parse_md_to_story(lines, title_md, h1, h2, body, bullet))

    doc.build(story)
    print(f"Wrote {OUT_PDF} ({OUT_PDF.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    build()
