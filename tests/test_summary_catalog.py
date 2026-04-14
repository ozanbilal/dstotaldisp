import sqlite3
import tempfile
from pathlib import Path
import numpy as np
import pandas as pd

from disp_core import _build_pair_summary_entry, _build_single_summary_entry, process_batch_files


def _build_db_bytes() -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".db3", delete=False) as tmp:
        db_path = Path(tmp.name)

    try:
        conn = sqlite3.connect(db_path)
        try:
            conn.execute(
                """
                CREATE TABLE PROFILES (
                    LAYER_NUMBER INTEGER PRIMARY KEY,
                    DEPTH_LAYER_TOP REAL NOT NULL,
                    DEPTH_LAYER_MID REAL NOT NULL,
                    MIN_DISP_RELATIVE REAL NOT NULL,
                    MAX_DISP_RELATIVE REAL NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE VEL_DISP (
                    TIME REAL NOT NULL,
                    LAYER1_DISP_TOTAL REAL NOT NULL,
                    LAYER1_DISP_RELATIVE REAL NOT NULL
                )
                """
            )
            conn.execute(
                "INSERT INTO PROFILES (LAYER_NUMBER, DEPTH_LAYER_TOP, DEPTH_LAYER_MID, MIN_DISP_RELATIVE, MAX_DISP_RELATIVE) VALUES (?, ?, ?, ?, ?)",
                (1, 0.0, 0.5, 0.0, 0.0),
            )
            conn.executemany(
                "INSERT INTO VEL_DISP (TIME, LAYER1_DISP_TOTAL, LAYER1_DISP_RELATIVE) VALUES (?, ?, ?)",
                [
                    (0.0, 0.00, 0.00),
                    (0.5, 0.10, 0.02),
                    (1.0, 0.15, 0.04),
                ],
            )
            conn.commit()
        finally:
            conn.close()

        return db_path.read_bytes()
    finally:
        try:
            db_path.unlink()
        except OSError:
            pass


def test_pair_summary_entry_prefers_strain_input_when_full_profile_exists():
    comparison_df = pd.DataFrame(
        {
            "Depth_m": [0.5, 1.5, 2.5],
            "Total_profile_offset_total_est_m": [0.15, 0.19, 0.26],
            "TimeHist_Resultant_total_m": [0.17, 0.22, 0.29],
            "Profile_RSS_total_m": [0.14, 0.18, 0.24],
        }
    )
    profile_sheet_df = pd.DataFrame({"Depth_m": [0.5, 1.5, 2.5]})
    strain_bundle = {
        "base_reference": "deepest_layer",
        "u_rel_base_x": np.array([[0.01, 0.02, 0.03], [0.02, 0.03, 0.04], [0.03, 0.04, 0.05]], dtype=float),
        "u_rel_base_y": np.array([[0.01, 0.01, 0.02], [0.02, 0.02, 0.03], [0.03, 0.03, 0.04]], dtype=float),
        "u_input_proxy_x": np.array([0.01, 0.01, 0.02], dtype=float),
        "u_input_proxy_y": np.array([0.01, 0.02, 0.02], dtype=float),
        "u_base_ref_x": np.array([0.02, 0.02, 0.03], dtype=float),
        "u_base_ref_y": np.array([0.02, 0.03, 0.03], dtype=float),
    }

    entry = _build_pair_summary_entry(
        object(),
        object(),
        "synthetic_X.xlsx",
        "synthetic_Y.xlsx",
        comparison_df,
        profile_sheet_df,
        strain_bundle,
        {},
        {},
        options={},
    )

    assert entry["summaryKind"] == "pair"
    assert entry["preferredVariantKey"] == "strain_input_total"
    assert entry["coverage"]["limitedData"] is False
    assert entry["detailSourceIds"][0].startswith("source-pair-")
    assert "db_direct_total" not in {variant["variantKey"] for variant in entry["variants"]}


def test_db_direct_summary_catalog_prefers_direct_total():
    summary = process_batch_files(
        {"sample_X.db3": _build_db_bytes()},
        {
            "useDb3Directly": True,
            "method2Enabled": True,
            "method3Enabled": True,
            "_returnWebResults": True,
        },
    )

    assert len(summary["summaryCatalog"]) == 1
    entry = summary["summaryCatalog"][0]

    assert entry["summaryKind"] == "db_pair"
    assert entry["preferredVariantKey"] == "db_direct_total"
    assert entry["coverage"]["limitedData"] is False
    assert [variant["variantKey"] for variant in entry["variants"]] == ["db_direct_total"]


def test_limited_layer_single_summary_hides_full_layer_variants():
    summary_df = pd.DataFrame(
        {
            "Depth_m": [0.5, 1.5],
            "Profile_offset_total_est_m": [0.10, 0.18],
            "TimeHist_maxabs_m": [0.11, 0.19],
            "Profile_max_m": [0.09, 0.16],
        }
    )
    profile_sheet_df = pd.DataFrame({"Depth_m": [0.5, 1.5, 2.5]})
    strain_bundle = {
        "base_reference": "deepest_layer",
        "u_rel_base": np.array([[0.01, 0.02, 0.03], [0.02, 0.03, 0.04]], dtype=float),
        "u_input_proxy": np.array([0.01, 0.01, 0.01], dtype=float),
        "u_base_ref": np.array([0.02, 0.02, 0.02], dtype=float),
    }

    entry = _build_single_summary_entry(
        object(),
        "limited_case_X.xlsx",
        "X",
        summary_df,
        profile_sheet_df,
        strain_bundle,
        {},
        options={},
    )

    variants = {variant["variantKey"]: variant for variant in entry["variants"]}

    assert entry["coverage"]["limitedData"] is True
    assert entry["preferredVariantKey"] == "profile_offset_total"
    assert entry["warnings"]
    assert variants["strain_input_total"]["valid"] is False
    assert variants["strain_deepest_total"]["valid"] is False
    assert variants["time_history_total"]["valid"] is False
    assert variants["profile_offset_total"]["valid"] is True
    assert variants["profile_reference_total"]["valid"] is True
