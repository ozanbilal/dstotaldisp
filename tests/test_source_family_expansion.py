from io import BytesIO

import pandas as pd
from openpyxl import Workbook

from disp_core import (
    _pair_derived_family,
    _read_input_motion_curves,
    _single_derived_family,
    _single_profile_family,
)


def test_read_input_motion_curves_accepts_5pct_damped_spectral_alias():
    motion = pd.DataFrame(
        {
            "Time (s)": [0.0, 0.1, 0.2],
            "Acceleration (g)": [0.0, 0.25, -0.1],
            "Period (s)": [0.01, 0.02, 0.03],
            "5% Damped Spectral": [0.74, 0.75, 0.76],
            "Frequency": [0.1, 0.2, 0.3],
            "Fourier Amplitude": [1.0, 1.02, 1.04],
        }
    )

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        motion.to_excel(writer, sheet_name="Input Motion", index=False)
    buffer.seek(0)

    with pd.ExcelFile(buffer, engine="openpyxl") as xl:
        curves = _read_input_motion_curves(xl)

    assert "acceleration" in curves
    assert "psa" in curves
    assert "fourier" in curves
    period, spectrum = curves["psa"]
    assert period.tolist() == [0.01, 0.02, 0.03]
    assert spectrum.tolist() == [0.74, 0.75, 0.76]


def test_single_profile_family_includes_pga_and_effective_stress_charts():
    wb = Workbook()
    ws = wb.active
    ws.title = "Profile"
    ws.append(
        [
            "Maximum Displacement",
            "Displacement (m)",
            "PGA (g)",
            "PGA Value",
            "Effective Stress",
            "Effective Stress (kPa)",
        ]
    )
    ws.append([None, None, None, None, None, None])
    ws.append([0.0, 0.025, 0.0, 0.31, 0.0, 180.0])
    ws.append([5.0, 0.016, 5.0, 0.27, 5.0, 230.0])
    ws.append([10.0, 0.010, 10.0, 0.21, 10.0, 290.0])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    with pd.ExcelFile(buffer, engine="openpyxl") as xl:
        family = _single_profile_family(xl)

    assert family is not None
    chart_keys = [item["chartKey"] for item in family["charts"]]
    assert "profile-max-displacement" in chart_keys
    assert "profile-pga" in chart_keys
    assert "profile-effective-stress" in chart_keys


def test_derived_families_hide_relative_profiles_and_keep_total_profiles():
    single_summary_df = pd.DataFrame(
        {
            "Depth_m": [0.0, 5.0, 10.0],
            "TBDY_total_max_m": [0.035, 0.024, 0.013],
        }
    )
    single_profile_df = pd.DataFrame(
        {
            "Depth_m": [0.0, 5.0, 10.0],
            "Profile_raw_max_m": [0.03, 0.02, 0.01],
            "Profile_relative_m": [0.02, 0.01, 0.0],
        }
    )
    single_family = _single_derived_family(
        "sample.xlsx",
        single_summary_df,
        single_profile_df,
        input_motion_max_abs=0.004,
    )
    assert single_family is not None
    assert [item["chartKey"] for item in single_family["charts"]] == ["derived-total-profile"]

    pair_profile_df = pd.DataFrame(
        {
            "Depth_m": [0.0, 5.0, 10.0],
            "Profile_X_raw_max_m": [0.03, 0.02, 0.01],
            "Profile_Y_raw_max_m": [0.028, 0.019, 0.011],
            "Profile_RSS_raw_max_m": [0.041, 0.028, 0.015],
        }
    )
    comparison_df = pd.DataFrame(
        {
            "Depth_m": [0.0, 5.0, 10.0],
            "X_tbdy_total_max_m": [0.034, 0.023, 0.012],
            "Y_tbdy_total_max_m": [0.033, 0.022, 0.011],
            "Total_tbdy_total_max_m": [0.047, 0.032, 0.017],
            "TimeHist_Resultant_total_m": [0.046, 0.031, 0.016],
        }
    )
    x_input_added_df = pd.DataFrame({"Depth_m": [0.0, 5.0, 10.0], "X_input_added_total_m": [0.032, 0.022, 0.012]})
    y_input_added_df = pd.DataFrame({"Depth_m": [0.0, 5.0, 10.0], "Y_input_added_total_m": [0.031, 0.021, 0.011]})
    pair_family = _pair_derived_family(comparison_df, pair_profile_df, x_input_added_df, y_input_added_df)

    assert pair_family is not None
    assert [item["chartKey"] for item in pair_family["charts"]] == [
        "derived-x-profile",
        "derived-y-profile",
        "derived-resultant-profile",
    ]
    for chart in pair_family["charts"]:
        for series in chart.get("series", []):
            assert "relative" not in str(series.get("name", "")).lower()
