"""
01_data_acquisition.py
----------------------
OWNER: P1 (Data Engineer)

Purpose
    Download all raw cohorts required for the project:
      (1)  TCGA-COAD + TCGA-READ   - RNA-seq + clinical  (primary / training)
      (2)  GSE39582                - microarray          (external validation)
      (3)  GSE17538                - microarray + survival (secondary)

Outputs (all dropped in  data/raw/<cohort>/)
    * expression matrices (gene x sample, raw counts or intensities)
    * clinical/phenotype tables (sample x clinical attributes)

Notes
    * TCGA download uses the public GDC API (no authentication needed for open
      data).  If the GDC endpoint is unreachable on a given machine, the
      script falls back to a pre-baked manifest URL — set `USE_MANIFEST = True`.
    * GEO downloads use the GEOparse library.
    * Every function is idempotent:  re-running does not re-download files
      that already exist locally.
"""

from __future__ import annotations
import json
import os
import tarfile
import shutil
import gzip
from pathlib import Path
from typing import Iterable

import pandas as pd
import requests
from tqdm import tqdm

import config
from utils import get_logger

log = get_logger("acquisition")

# ------------------------------------------------------------------
# TCGA  (GDC open API)
# ------------------------------------------------------------------
GDC_FILES   = "https://api.gdc.cancer.gov/files"
GDC_DATA    = "https://api.gdc.cancer.gov/data"
GDC_CLINICAL= "https://api.gdc.cancer.gov/cases"


def _gdc_query(project: str,
               data_type: str = "Gene Expression Quantification",
               workflow: str = "STAR - Counts") -> list[dict]:
    """Return GDC file manifest for a given project + workflow."""
    filters = {
        "op": "and",
        "content": [
            {"op": "in", "content": {"field": "cases.project.project_id",
                                     "value": [project]}},
            {"op": "in", "content": {"field": "files.data_type",
                                     "value": [data_type]}},
            {"op": "in", "content": {"field": "files.analysis.workflow_type",
                                     "value": [workflow]}},
            {"op": "in", "content": {"field": "access", "value": ["open"]}},
        ],
    }
    params = {
        "filters": json.dumps(filters),
        "fields":  "file_id,file_name,cases.submitter_id,cases.samples.sample_type",
        "format":  "JSON",
        "size":    "5000",
    }
    r = requests.get(GDC_FILES, params=params, timeout=60)
    r.raise_for_status()
    return r.json()["data"]["hits"]


def download_tcga_project(project: str, batch_size: int = 50,
                          max_retries: int = 3) -> Path:
    """Download STAR-Counts gene-expression files + clinical for a TCGA project.

    Downloads in BATCHES of `batch_size` files to avoid GDC's habit of dropping
    the connection on very large single-shot POSTs. Each batch is resumable —
    if it fails, we retry; if it succeeds, the extracted folder is kept and
    future runs skip it.
    """
    out_dir = config.RAW_DIR / project
    out_dir.mkdir(parents=True, exist_ok=True)

    manifest_path = out_dir / "manifest.json"
    if manifest_path.exists():
        hits = json.loads(manifest_path.read_text())
    else:
        log.info("Querying GDC for %s ...", project)
        hits = _gdc_query(project)
        manifest_path.write_text(json.dumps(hits, indent=2))
        log.info("Found %d files for %s", len(hits), project)

    extract_dir = out_dir / "star_counts"
    extract_dir.mkdir(exist_ok=True)

    file_ids = [h["file_id"] for h in hits]
    # Skip IDs we already extracted
    pending = [fid for fid in file_ids if not (extract_dir / fid).exists()]
    if not pending:
        log.info("%s: all %d files already downloaded", project, len(file_ids))
        _parse_tcga_expression(extract_dir, hits, out_dir)
        _download_tcga_clinical(project, out_dir)
        return out_dir

    log.info("%s: downloading %d / %d files in batches of %d",
             project, len(pending), len(file_ids), batch_size)

    batches = [pending[i:i + batch_size]
               for i in range(0, len(pending), batch_size)]
    for bi, batch in enumerate(batches, 1):
        _download_batch(batch, extract_dir, f"{project} batch {bi}/{len(batches)}",
                        max_retries=max_retries)

    _parse_tcga_expression(extract_dir, hits, out_dir)
    _download_tcga_clinical(project, out_dir)
    return out_dir


def _download_batch(file_ids: list[str], extract_dir: Path,
                    desc: str, max_retries: int = 3) -> None:
    """Download a batch of GDC files as a tar.gz, extract in place, delete tar.

    Retries up to max_retries on network/decode errors. Skips files already
    extracted (idempotent).
    """
    pending = [fid for fid in file_ids if not (extract_dir / fid).exists()]
    if not pending:
        return

    tmp_tar = extract_dir / f".batch_{abs(hash(tuple(pending))) & 0xFFFFFF}.tar.gz"
    for attempt in range(1, max_retries + 1):
        try:
            with requests.post(GDC_DATA,
                               json={"ids": pending},
                               headers={"Content-Type": "application/json"},
                               stream=True, timeout=600) as r:
                r.raise_for_status()
                total = int(r.headers.get("content-length", 0))
                with open(tmp_tar, "wb") as fh, tqdm(total=total,
                                                    unit="B", unit_scale=True,
                                                    desc=desc, leave=False) as pbar:
                    for chunk in r.iter_content(chunk_size=1 << 20):
                        if chunk:
                            fh.write(chunk)
                            pbar.update(len(chunk))
            # Extract
            with tarfile.open(tmp_tar) as tf:
                tf.extractall(extract_dir)
            tmp_tar.unlink(missing_ok=True)
            return
        except (requests.exceptions.RequestException,
                tarfile.ReadError, EOFError, OSError) as e:
            log.warning("  [attempt %d/%d] %s failed: %s",
                        attempt, max_retries, desc, e)
            tmp_tar.unlink(missing_ok=True)
            if attempt == max_retries:
                raise
            import time
            time.sleep(5 * attempt)


def _parse_tcga_expression(extract_dir: Path,
                           hits: list[dict],
                           out_dir: Path) -> None:
    """Consolidate GDC STAR-Counts files into one gene x sample matrix."""
    matrix_path = out_dir / "expression_counts.tsv"
    if matrix_path.exists():
        log.info("Expression matrix already parsed at %s", matrix_path.name)
        return

    id_to_sample = {
        h["file_id"]: h["cases"][0]["submitter_id"] for h in hits
    }

    frames = []
    for fid, sample in tqdm(id_to_sample.items(), desc="parsing"):
        tsv = next((extract_dir / fid).glob("*.tsv"), None)
        if tsv is None:
            continue
        df = pd.read_csv(tsv, sep="\t", comment="#",
                         skiprows=[2, 3, 4, 5])
        df = df.set_index("gene_name")[["unstranded"]]
        df.columns = [sample]
        frames.append(df)
    expr = pd.concat(frames, axis=1).groupby(level=0).sum()
    expr.to_csv(matrix_path, sep="\t")
    log.info("Wrote %s  (%s genes x %s samples)",
             matrix_path.name, *expr.shape)


def _download_tcga_clinical(project: str, out_dir: Path) -> None:
    """Grab minimal clinical (stage, MSI status, survival) via GDC API."""
    path = out_dir / "clinical.tsv"
    if path.exists():
        return
    filters = {"op": "in",
               "content": {"field": "project.project_id", "value": [project]}}
    params = {
        "filters": json.dumps(filters),
        "fields": ",".join([
            "submitter_id",
            "demographic.vital_status",
            "demographic.days_to_death",
            "diagnoses.days_to_last_follow_up",
            "diagnoses.ajcc_pathologic_stage",
            "diagnoses.tumor_stage",
        ]),
        "format": "TSV",
        "size":   "5000",
    }
    r = requests.get(GDC_CLINICAL, params=params, timeout=60)
    r.raise_for_status()
    path.write_bytes(r.content)
    log.info("Saved clinical  -> %s", path.name)


# ------------------------------------------------------------------
# GEO  (GEOparse)
# ------------------------------------------------------------------
def download_geo(gse_id: str, max_retries: int = 5) -> Path:
    """Download a GEO series and build a simple expression + pheno table.

    Retries up to `max_retries` times with exponential back-off — NCBI's FTP
    server frequently drops big transfers, and GEOparse fails hard on the
    slightest size mismatch. Partial SOFT files are removed before each retry.
    """
    try:
        import GEOparse
    except ImportError as e:
        raise RuntimeError("Install GEOparse:  pip install GEOparse") from e

    out_dir = config.RAW_DIR / gse_id
    out_dir.mkdir(parents=True, exist_ok=True)

    expr_path  = out_dir / "expression.tsv"
    pheno_path = out_dir / "phenotype.tsv"
    if expr_path.exists() and pheno_path.exists():
        log.info("%s already downloaded.", gse_id)
        return out_dir

    import time
    last_err: Exception | None = None
    for attempt in range(1, max_retries + 1):
        try:
            log.info("Downloading %s  (attempt %d/%d) ...",
                     gse_id, attempt, max_retries)
            # Remove any partial SOFT file from previous failed attempt
            for partial in out_dir.glob("*_family.soft.gz"):
                partial.unlink(missing_ok=True)

            gse = GEOparse.get_GEO(geo=gse_id, destdir=str(out_dir), silent=True)

            # Expression matrix
            frames = {}
            for name, gsm in gse.gsms.items():
                if not gsm.table.empty and "VALUE" in gsm.table.columns:
                    frames[name] = gsm.table.set_index("ID_REF")["VALUE"]
            if not frames:
                raise RuntimeError(f"{gse_id}: no GSM tables contained VALUE column")
            expr = pd.concat(frames, axis=1)
            expr.to_csv(expr_path, sep="\t")

            # Phenotype
            pheno = pd.DataFrame({name: gsm.metadata
                                  for name, gsm in gse.gsms.items()}).T
            pheno.to_csv(pheno_path, sep="\t")

            log.info("%s:  expression %s,  phenotype %s",
                     gse_id, expr.shape, pheno.shape)
            return out_dir

        except Exception as e:       # noqa: BLE001
            last_err = e
            log.warning("  [%s attempt %d/%d] failed: %s",
                        gse_id, attempt, max_retries, e)
            time.sleep(10 * attempt)    # 10, 20, 30, 40, 50 s

    raise RuntimeError(f"{gse_id} failed after {max_retries} retries: {last_err}")


# ------------------------------------------------------------------
# ENTRY POINT
# ------------------------------------------------------------------
def main() -> None:
    log.info("=== DATA ACQUISITION ===")
    for project in config.TCGA_PROJECTS:
        try:
            download_tcga_project(project)
        except Exception as e:     # noqa: BLE001
            log.error("TCGA download for %s failed: %s", project, e)
            log.error("  Manual fallback: use GDC Data Transfer Tool or cBioPortal.")

    for gse in (config.GEO_VALIDATION, config.GEO_SURVIVAL):
        try:
            download_geo(gse)
        except Exception as e:     # noqa: BLE001
            log.error("GEO download for %s failed: %s", gse, e)


if __name__ == "__main__":
    main()
