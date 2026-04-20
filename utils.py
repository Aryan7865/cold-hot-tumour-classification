"""
utils.py
--------
Shared helpers used by every module: IO, figure saving, logging, etc.
"""

from __future__ import annotations
import logging
import pickle
from pathlib import Path
from typing import Any

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

import config

# ------------------------------------------------------------------
# LOGGER
# ------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
    datefmt="%H:%M:%S",
)


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)


log = get_logger(__name__)

# ------------------------------------------------------------------
# PLOT STYLE
# ------------------------------------------------------------------
def set_plot_style() -> None:
    sns.set_theme(style=config.FIG_STYLE, context="talk")
    plt.rcParams["figure.dpi"] = config.FIG_DPI
    plt.rcParams["savefig.dpi"] = config.FIG_DPI
    plt.rcParams["axes.titleweight"] = "bold"


def save_fig(fig: plt.Figure, name: str, subfolder: str | None = None) -> Path:
    """Save a figure to outputs/figures/<subfolder>/<name>.png"""
    target_dir = config.FIGURES_DIR / subfolder if subfolder else config.FIGURES_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    out = target_dir / f"{name}.png"
    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight")
    plt.close(fig)
    log.info("Saved figure: %s", out.relative_to(config.PROJECT_ROOT))
    return out


# ------------------------------------------------------------------
# DATA IO
# ------------------------------------------------------------------
def save_table(df: pd.DataFrame, name: str, subfolder: str | None = None) -> Path:
    target_dir = config.TABLES_DIR / subfolder if subfolder else config.TABLES_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    out = target_dir / f"{name}.csv"
    df.to_csv(out, index=True)
    log.info("Saved table : %s (%s rows)",
             out.relative_to(config.PROJECT_ROOT), len(df))
    return out


def load_table(name: str, subfolder: str | None = None, **kwargs) -> pd.DataFrame:
    target_dir = config.TABLES_DIR / subfolder if subfolder else config.TABLES_DIR
    path = target_dir / f"{name}.csv"
    return pd.read_csv(path, **kwargs)


def save_pickle(obj: Any, name: str, subfolder: str | None = None) -> Path:
    target_dir = config.MODELS_DIR / subfolder if subfolder else config.MODELS_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    out = target_dir / f"{name}.pkl"
    with open(out, "wb") as fh:
        pickle.dump(obj, fh)
    log.info("Saved object: %s", out.relative_to(config.PROJECT_ROOT))
    return out


def load_pickle(name: str, subfolder: str | None = None) -> Any:
    target_dir = config.MODELS_DIR / subfolder if subfolder else config.MODELS_DIR
    path = target_dir / f"{name}.pkl"
    with open(path, "rb") as fh:
        return pickle.load(fh)


# ------------------------------------------------------------------
# CONVENIENCE
# ------------------------------------------------------------------
def show_summary(df: pd.DataFrame, name: str = "") -> None:
    log.info("── %s summary ──", name)
    log.info("shape: %s", df.shape)
    log.info("head:\n%s", df.iloc[:3, :5])
