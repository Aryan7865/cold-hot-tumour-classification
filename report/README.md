# BT3041 Final Report

This folder contains the styled final report PDF and its generator script.

- Output PDF: `BT3041_Cold_Hot_Tumour_Report.pdf`
- Generator: `generate_report_pdf.py`

## Regenerate

```bash
source .venv/bin/activate
python report/generate_report_pdf.py
```

The script pulls metrics/tables from `outputs/tables/` and figures from `outputs/figures/`.

## Presentation (PowerPoint)

- Output: `BT3041_Cold_Hot_Tumour_Presentation.pptx`
- Generator: `build_presentation.py` (metrics are read live from the same CSVs — regenerate after re-running the pipeline).

```bash
source .venv/bin/activate
pip install -r requirements.txt   # includes python-pptx
python report/build_presentation.py
```
