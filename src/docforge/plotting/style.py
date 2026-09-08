"""Shared Matplotlib visual system for docforge-generated figures.

The palette and typography follow the restrained style used across the
INVAR-2025 manuscripts: white backgrounds, sans-serif faces, thin charcoal
lines, muted semantic colors, direct labels, and editable vector PDF output.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

CHARCOAL = "#2F2F2F"
DARK_GRAY = "#686868"
MID_GRAY = "#B7B7B7"
LIGHT_GRAY = "#D9D9D9"
PALE_GRAY = "#F2F2F2"
WHITE = "#FFFFFF"
NAVY = "#26516B"          # existing / published evidence
NAVY_LIGHT = "#8EA8B5"    # preparation / in progress
SOFT_NAVY = "#E6EFF4"
CRIMSON = "#B50030"       # risk / high-error region
CRIMSON_LIGHT = "#D97085"
SOFT_CRIMSON = "#FAEBEE"
JADE = "#007065"          # selected/adapted component
JADE_LIGHT = "#4DA698"
SOFT_JADE = "#E8F4F1"
OCHRE = "#C9A24B"
ORANGE = "#C66A4E"
SOFT_ORANGE = "#FBEEE6"
PURPLE = "#705080"
SOFT_PURPLE = "#F0EDF5"
SKY = "#77B9D1"

MM_PER_INCH = 25.4


@dataclass(frozen=True)
class FigureSize:
    width_mm: float = 170.0
    height_mm: float = 85.0

    @property
    def inches(self) -> tuple[float, float]:
        return self.width_mm / MM_PER_INCH, self.height_mm / MM_PER_INCH


def set_style() -> None:
    """Apply the shared plotting style."""
    mpl.rcParams.update(
        {
            "font.family": "sans-serif",
            # Put the CJK-capable font first. It also contains Latin glyphs and
            # avoids Matplotlib selecting Arial for a mixed Chinese string.
            "font.sans-serif": ["Microsoft YaHei", "Arial", "DejaVu Sans"],
            "font.size": 7.5,
            "text.color": CHARCOAL,
            "axes.labelcolor": CHARCOAL,
            "axes.edgecolor": CHARCOAL,
            "axes.linewidth": 0.55,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "xtick.color": CHARCOAL,
            "ytick.color": CHARCOAL,
            "xtick.direction": "out",
            "ytick.direction": "out",
            "xtick.major.size": 2.5,
            "ytick.major.size": 2.5,
            "xtick.major.width": 0.5,
            "ytick.major.width": 0.5,
            "legend.frameon": False,
            "figure.facecolor": WHITE,
            "axes.facecolor": WHITE,
            "savefig.facecolor": WHITE,
            "savefig.transparent": False,
            "pdf.fonttype": 42,
            "ps.fonttype": 42,
            "svg.fonttype": "none",
        }
    )


def new_figure(size: FigureSize = FigureSize()) -> plt.Figure:
    set_style()
    return plt.figure(figsize=size.inches, facecolor=WHITE)


def panel_label(ax: plt.Axes, label: str, x: float = -0.08, y: float = 1.04) -> None:
    ax.text(
        x,
        y,
        label,
        transform=ax.transAxes,
        fontsize=9,
        fontweight="bold",
        ha="left",
        va="bottom",
        clip_on=False,
    )


def prepare_schematic_axis(ax: plt.Axes) -> None:
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")


def rounded_box(
    ax: plt.Axes,
    xy: tuple[float, float],
    width: float,
    height: float,
    *,
    facecolor: str = WHITE,
    edgecolor: str = LIGHT_GRAY,
    linewidth: float = 0.7,
    linestyle: str = "-",
    radius: float = 0.025,
    transform=None,
    zorder: int = 2,
) -> FancyBboxPatch:
    transform = transform or ax.transData
    patch = FancyBboxPatch(
        xy,
        width,
        height,
        boxstyle=f"round,pad=0.012,rounding_size={radius}",
        facecolor=facecolor,
        edgecolor=edgecolor,
        linewidth=linewidth,
        linestyle=linestyle,
        transform=transform,
        zorder=zorder,
    )
    ax.add_patch(patch)
    return patch


def arrow(
    ax: plt.Axes,
    start: tuple[float, float],
    end: tuple[float, float],
    *,
    color: str = DARK_GRAY,
    linewidth: float = 0.8,
    linestyle: str = "-",
    connectionstyle: str = "arc3,rad=0",
    mutation_scale: float = 7,
    transform=None,
    zorder: int = 3,
) -> FancyArrowPatch:
    transform = transform or ax.transData
    patch = FancyArrowPatch(
        start,
        end,
        arrowstyle="-|>",
        mutation_scale=mutation_scale,
        linewidth=linewidth,
        linestyle=linestyle,
        color=color,
        connectionstyle=connectionstyle,
        transform=transform,
        zorder=zorder,
    )
    ax.add_patch(patch)
    return patch


def status_chip(
    ax: plt.Axes,
    x: float,
    y: float,
    text: str,
    *,
    facecolor: str,
    edgecolor: str,
    textcolor: str | None = None,
    width: float | None = None,
) -> None:
    textcolor = textcolor or edgecolor
    width = width or max(0.16, 0.016 * len(text))
    rounded_box(
        ax,
        (x, y),
        width,
        0.085,
        facecolor=facecolor,
        edgecolor=edgecolor,
        linewidth=0.65,
        radius=0.03,
        transform=ax.transAxes,
        zorder=4,
    )
    ax.text(
        x + width / 2,
        y + 0.042,
        text,
        transform=ax.transAxes,
        fontsize=6.6,
        color=textcolor,
        ha="center",
        va="center",
        zorder=5,
    )


def save_figure(
    fig: plt.Figure,
    stem: str,
    out_dir: Path,
    *,
    formats: tuple[str, ...] = ("pdf", "png"),
    dpi: int = 600,
) -> list[Path]:
    """Save *fig* under *stem* into *out_dir* in the requested formats.

    PNG is written at *dpi* (default 600); PDF and SVG remain editable vector
    files. Returns the written paths.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for fmt in formats:
        path = out_dir / f"{stem}.{fmt}"
        if fmt == "pdf":
            with PdfPages(path) as pdf:
                pdf.savefig(fig)
        elif fmt == "svg":
            fig.savefig(path, format="svg")
        elif fmt == "png":
            fig.savefig(path, format="png", dpi=dpi)
        else:
            raise ValueError(f"Unsupported figure format: {fmt}")
        written.append(path)
    return written


__all__ = [
    "CHARCOAL",
    "DARK_GRAY",
    "MID_GRAY",
    "LIGHT_GRAY",
    "PALE_GRAY",
    "WHITE",
    "NAVY",
    "NAVY_LIGHT",
    "SOFT_NAVY",
    "CRIMSON",
    "CRIMSON_LIGHT",
    "SOFT_CRIMSON",
    "JADE",
    "JADE_LIGHT",
    "SOFT_JADE",
    "OCHRE",
    "ORANGE",
    "SOFT_ORANGE",
    "PURPLE",
    "SOFT_PURPLE",
    "SKY",
    "FigureSize",
    "set_style",
    "new_figure",
    "panel_label",
    "prepare_schematic_axis",
    "rounded_box",
    "arrow",
    "status_chip",
    "save_figure",
]