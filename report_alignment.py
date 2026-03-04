from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


@dataclass
class ReportArtifacts:
    markdown_path: Path
    profile_plot_path: Path
    delta_plot_path: Path


def _compute_stats(comparison: pd.DataFrame, strain: pd.DataFrame, legacy: pd.DataFrame) -> dict:
    x_diff = pd.to_numeric(comparison["Delta_Xbase_vs_ProfileXminusbottom_m"], errors="coerce").to_numpy(dtype=float)
    y_diff = pd.to_numeric(comparison["Delta_Ybase_vs_ProfileYminusbottom_m"], errors="coerce").to_numpy(dtype=float)

    profile_x_bottom = float(pd.to_numeric(legacy["Profile_X_max_m"], errors="coerce").iloc[-1])
    profile_y_bottom = float(pd.to_numeric(legacy["Profile_Y_max_m"], errors="coerce").iloc[-1])
    profile_rss_bottom = float(pd.to_numeric(legacy["Profile_RSS_total_m"], errors="coerce").iloc[-1])

    return {
        "profile_x_bottom": profile_x_bottom,
        "profile_y_bottom": profile_y_bottom,
        "profile_rss_bottom": profile_rss_bottom,
        "x_mean_abs": float(np.nanmean(np.abs(x_diff))),
        "x_max_abs": float(np.nanmax(np.abs(x_diff))),
        "y_mean_abs": float(np.nanmean(np.abs(y_diff))),
        "y_max_abs": float(np.nanmax(np.abs(y_diff))),
        "surface_profile_x": float(pd.to_numeric(legacy["Profile_X_max_m"], errors="coerce").iloc[0]),
        "surface_profile_x_bc": float(pd.to_numeric(comparison["Profile_X_minus_bottom_m"], errors="coerce").iloc[0]),
        "surface_strain_x": float(pd.to_numeric(strain["X_base_rel_max_m"], errors="coerce").iloc[0]),
    }


def _plot_base_corrected_profiles(comparison: pd.DataFrame, output_png: Path) -> None:
    depth = pd.to_numeric(comparison["Depth_m"], errors="coerce").to_numpy(dtype=float)

    x_base = pd.to_numeric(comparison["X_base_rel_max_m"], errors="coerce").to_numpy(dtype=float)
    x_prof_bc = pd.to_numeric(comparison["Profile_X_minus_bottom_m"], errors="coerce").to_numpy(dtype=float)

    y_base = pd.to_numeric(comparison["Y_base_rel_max_m"], errors="coerce").to_numpy(dtype=float)
    y_prof_bc = pd.to_numeric(comparison["Profile_Y_minus_bottom_m"], errors="coerce").to_numpy(dtype=float)

    total_base = pd.to_numeric(comparison["Total_base_rel_max_m"], errors="coerce").to_numpy(dtype=float)
    total_prof_bc = pd.to_numeric(comparison["Profile_RSS_minus_bottom_m"], errors="coerce").to_numpy(dtype=float)

    fig, axes = plt.subplots(1, 3, figsize=(14.5, 4.8), sharey=True)

    axes[0].plot(x_base, depth, "-o", lw=2.0, ms=3.5, label="X_base_rel_max")
    axes[0].plot(x_prof_bc, depth, "-s", lw=1.8, ms=3.0, label="Profile_X_minus_bottom")
    axes[0].set_title("X Profile")
    axes[0].set_xlabel("Displacement (m)")
    axes[0].set_ylabel("Depth (m)")
    axes[0].grid(True, alpha=0.35)
    axes[0].legend(loc="best", fontsize=8)

    axes[1].plot(y_base, depth, "-o", lw=2.0, ms=3.5, label="Y_base_rel_max")
    axes[1].plot(y_prof_bc, depth, "-s", lw=1.8, ms=3.0, label="Profile_Y_minus_bottom")
    axes[1].set_title("Y Profile")
    axes[1].set_xlabel("Displacement (m)")
    axes[1].grid(True, alpha=0.35)
    axes[1].legend(loc="best", fontsize=8)

    axes[2].plot(total_base, depth, "-o", lw=2.0, ms=3.5, label="Total_base_rel_max")
    axes[2].plot(total_prof_bc, depth, "-s", lw=1.8, ms=3.0, label="Profile_RSS_minus_bottom")
    axes[2].set_title("Total Profile")
    axes[2].set_xlabel("Displacement (m)")
    axes[2].grid(True, alpha=0.35)
    axes[2].legend(loc="best", fontsize=8)

    axes[0].invert_yaxis()
    fig.suptitle("Base-Corrected Profile Alignment", fontsize=14, fontweight="bold")
    fig.tight_layout(rect=[0, 0, 1, 0.95])

    output_png.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_png, dpi=160)
    plt.close(fig)


def _plot_alignment_deltas(comparison: pd.DataFrame, output_png: Path) -> None:
    depth = pd.to_numeric(comparison["Depth_m"], errors="coerce").to_numpy(dtype=float)
    dx = pd.to_numeric(comparison["Delta_Xbase_vs_ProfileXminusbottom_m"], errors="coerce").to_numpy(dtype=float)
    dy = pd.to_numeric(comparison["Delta_Ybase_vs_ProfileYminusbottom_m"], errors="coerce").to_numpy(dtype=float)
    dt = (
        pd.to_numeric(comparison["Total_base_rel_max_m"], errors="coerce").to_numpy(dtype=float)
        - pd.to_numeric(comparison["Profile_RSS_minus_bottom_m"], errors="coerce").to_numpy(dtype=float)
    )

    fig, ax = plt.subplots(figsize=(7.2, 5.2))
    ax.axvline(0.0, color="k", lw=1.0, alpha=0.5)
    ax.plot(dx, depth, "-o", lw=1.8, ms=3.5, label="Delta X")
    ax.plot(dy, depth, "-s", lw=1.8, ms=3.2, label="Delta Y")
    ax.plot(dt, depth, "-^", lw=1.8, ms=3.2, label="Delta Total")

    ax.set_title("Base-Corrected Alignment Delta")
    ax.set_xlabel("Delta (m)")
    ax.set_ylabel("Depth (m)")
    ax.grid(True, alpha=0.35)
    ax.legend(loc="best", fontsize=8)
    ax.invert_yaxis()

    fig.tight_layout()
    output_png.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_png, dpi=160)
    plt.close(fig)


def _build_markdown(
    report_name: str,
    stats: dict,
    sample_rows: pd.DataFrame,
    profile_plot_file: str,
    delta_plot_file: str,
) -> str:
    lines: list[str] = []
    lines.append(f"# {report_name}")
    lines.append("")
    lines.append("## Ozet")
    lines.append("- Deepsoil Profile Maximum Displacement verisinde taban ofseti gorulmektedir.")
    lines.append("- Strain tabanli base-relative profile ile uyum, taban ofseti cikarilinca belirgin sekilde artar.")
    lines.append("")
    lines.append("## Grafikler")
    lines.append("")
    lines.append(f"![Base-Corrected Profiles]({profile_plot_file})")
    lines.append("")
    lines.append(f"![Alignment Deltas]({delta_plot_file})")
    lines.append("")
    lines.append("## Sayisal Bulgular")
    lines.append(f"- Profile X taban ofseti: `{stats['profile_x_bottom']:.6f} m`")
    lines.append(f"- Profile Y taban ofseti: `{stats['profile_y_bottom']:.6f} m`")
    lines.append(f"- Profile RSS taban ofseti: `{stats['profile_rss_bottom']:.6f} m`")
    lines.append(f"- Layer 1 Profile X max: `{stats['surface_profile_x']:.6f} m`")
    lines.append(f"- Layer 1 Profile X (base-corrected): `{stats['surface_profile_x_bc']:.6f} m`")
    lines.append(f"- Layer 1 Strain X base-relative max: `{stats['surface_strain_x']:.6f} m`")
    lines.append(f"- X ortalama mutlak fark: `{stats['x_mean_abs']:.6f} m`")
    lines.append(f"- X maksimum mutlak fark: `{stats['x_max_abs']:.6f} m`")
    lines.append(f"- Y ortalama mutlak fark: `{stats['y_mean_abs']:.6f} m`")
    lines.append(f"- Y maksimum mutlak fark: `{stats['y_max_abs']:.6f} m`")
    lines.append("")
    lines.append("## Ilk 10 Katman Ornek Tablo (X)")
    lines.append("")
    lines.append(
        "|Layer|Depth_m|X_base_rel_max_m|Profile_X_max_m|Profile_X_minus_bottom_m|Delta_Xbase_vs_ProfileXminusbottom_m|"
    )
    lines.append("|---:|---:|---:|---:|---:|---:|")
    for _, row in sample_rows.iterrows():
        lines.append(
            f"|{int(row['Layer_Index'])}|{row['Depth_m']:.3f}|{row['X_base_rel_max_m']:.6f}|"
            f"{row['Profile_X_max_m']:.6f}|{row['Profile_X_minus_bottom_m']:.6f}|"
            f"{row['Delta_Xbase_vs_ProfileXminusbottom_m']:.6f}|"
        )
    lines.append("")
    return "\n".join(lines)


def generate_alignment_report(workbook_path: str | Path) -> ReportArtifacts:
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(workbook_path)

    comparison = pd.read_excel(workbook_path, sheet_name="Comparison")
    strain = pd.read_excel(workbook_path, sheet_name="Strain_Relative")
    legacy = pd.read_excel(workbook_path, sheet_name="Legacy_Methods")

    stats = _compute_stats(comparison, strain, legacy)

    stem = workbook_path.stem
    profile_plot_path = workbook_path.with_name(f"{stem}_base_corrected_profiles.png")
    delta_plot_path = workbook_path.with_name(f"{stem}_alignment_deltas.png")
    markdown_path = workbook_path.with_name(f"{stem}_alignment_report.md")

    _plot_base_corrected_profiles(comparison, profile_plot_path)
    _plot_alignment_deltas(comparison, delta_plot_path)

    sample_rows = comparison[
        [
            "Layer_Index",
            "Depth_m",
            "X_base_rel_max_m",
            "Profile_X_max_m",
            "Profile_X_minus_bottom_m",
            "Delta_Xbase_vs_ProfileXminusbottom_m",
        ]
    ].head(10)

    report_text = _build_markdown(
        report_name="Relative Displacement Alignment Report",
        stats=stats,
        sample_rows=sample_rows,
        profile_plot_file=profile_plot_path.name,
        delta_plot_file=delta_plot_path.name,
    )
    markdown_path.write_text(report_text, encoding="utf-8")

    return ReportArtifacts(
        markdown_path=markdown_path,
        profile_plot_path=profile_plot_path,
        delta_plot_path=delta_plot_path,
    )
