import sqlite3
import tempfile
from pathlib import Path

import pytest

from disp_core import process_batch_files


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
                """
                CREATE TABLE TIME_HISTORIES (
                    TIME REAL NOT NULL,
                    LAYER1_ACCEL REAL NOT NULL,
                    LAYER1_VEL REAL NOT NULL,
                    LAYER1_DISP REAL NOT NULL,
                    LAYER1_ARIAS REAL NOT NULL,
                    LAYER1_STRAIN REAL NOT NULL,
                    LAYER1_STRESS REAL NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE RESPONSE_SPECTRA (
                    PERIOD REAL NOT NULL,
                    INPUT_MOTION_RS REAL NOT NULL,
                    LAYER1_RS REAL NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE FOURIER_AMPLITUDE_SPECTRA (
                    FREQUENCY REAL NOT NULL,
                    INPUT_MOTION_FAS REAL NOT NULL,
                    LAYER1_FAS REAL NOT NULL,
                    LAYER1_FAS_RATIO REAL NOT NULL
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
            conn.executemany(
                (
                    "INSERT INTO TIME_HISTORIES "
                    "(TIME, LAYER1_ACCEL, LAYER1_VEL, LAYER1_DISP, LAYER1_ARIAS, LAYER1_STRAIN, LAYER1_STRESS) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)"
                ),
                [
                    (0.0, 0.01, 0.00, 0.00, 0.0, 0.001, 1.0),
                    (0.5, 0.03, 0.04, 0.05, 0.2, 0.002, 2.0),
                    (1.0, 0.02, 0.02, 0.03, 0.3, 0.003, 1.5),
                ],
            )
            conn.executemany(
                "INSERT INTO RESPONSE_SPECTRA (PERIOD, INPUT_MOTION_RS, LAYER1_RS) VALUES (?, ?, ?)",
                [(0.1, 0.2, 0.3), (1.0, 0.4, 0.5)],
            )
            conn.executemany(
                "INSERT INTO FOURIER_AMPLITUDE_SPECTRA (FREQUENCY, INPUT_MOTION_FAS, LAYER1_FAS, LAYER1_FAS_RATIO) VALUES (?, ?, ?, ?)",
                [(0.5, 1.0, 1.2, 1.2), (2.0, 0.5, 0.75, 1.5)],
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


def _build_time_history_db_bytes() -> bytes:
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
                    PGA_TOTAL REAL NOT NULL,
                    PGV_RELATIVE REAL NOT NULL,
                    MIN_DISP_RELATIVE REAL NOT NULL,
                    MAX_DISP_RELATIVE REAL NOT NULL,
                    DEPTH_LAYER_MID REAL NOT NULL,
                    INITIAL_EFFECTIVE_STRESS REAL NOT NULL,
                    MAX_STRAIN REAL NOT NULL,
                    MAX_STRESS_RATIO REAL NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE TIME_HISTORIES (
                    TIME REAL NOT NULL,
                    LAYER1_ACCEL REAL NOT NULL,
                    LAYER1_VEL REAL NOT NULL,
                    LAYER1_DISP REAL NOT NULL,
                    LAYER1_ARIAS REAL NOT NULL,
                    LAYER1_STRAIN REAL NOT NULL,
                    LAYER1_STRESS REAL NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE RESPONSE_SPECTRA (
                    PERIOD REAL NOT NULL,
                    INPUT_MOTION_RS REAL NOT NULL,
                    LAYER1_RS REAL NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE FOURIER_AMPLITUDE_SPECTRA (
                    FREQUENCY REAL NOT NULL,
                    INPUT_MOTION_FAS REAL NOT NULL,
                    LAYER1_FAS REAL NOT NULL,
                    LAYER1_FAS_RATIO REAL NOT NULL
                )
                """
            )
            conn.executemany(
                (
                    "INSERT INTO PROFILES "
                    "(LAYER_NUMBER, DEPTH_LAYER_TOP, PGA_TOTAL, PGV_RELATIVE, MIN_DISP_RELATIVE, MAX_DISP_RELATIVE, "
                    "DEPTH_LAYER_MID, INITIAL_EFFECTIVE_STRESS, MAX_STRAIN, MAX_STRESS_RATIO) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                ),
                [
                    (1, 0.0, 0.7, 0.2, -0.010, 0.012, 0.5, 9.0, 0.01, 0.6),
                    (2, 1.0, 0.6, 0.1, -0.008, 0.009, 1.5, 18.0, 0.02, 0.5),
                ],
            )
            conn.executemany(
                (
                    "INSERT INTO TIME_HISTORIES "
                    "(TIME, LAYER1_ACCEL, LAYER1_VEL, LAYER1_DISP, LAYER1_ARIAS, LAYER1_STRAIN, LAYER1_STRESS) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)"
                ),
                [(0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0), (0.1, 0.2, 0.03, 0.01, 0.1, 0.002, 1.0)],
            )
            conn.executemany(
                "INSERT INTO RESPONSE_SPECTRA (PERIOD, INPUT_MOTION_RS, LAYER1_RS) VALUES (?, ?, ?)",
                [(0.1, 0.2, 0.3), (1.0, 0.4, 0.5)],
            )
            conn.executemany(
                "INSERT INTO FOURIER_AMPLITUDE_SPECTRA (FREQUENCY, INPUT_MOTION_FAS, LAYER1_FAS, LAYER1_FAS_RATIO) VALUES (?, ?, ?, ?)",
                [(0.5, 1.0, 1.2, 1.2), (2.0, 0.5, 0.75, 1.5)],
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


def test_db_direct_results_are_classified_as_db_direct():
    summary = process_batch_files(
        {"sample_X.db3": _build_db_bytes()},
        {
            "useDb3Directly": True,
            "method2Enabled": True,
            "method3Enabled": True,
            "_returnWebResults": True,
        },
    )

    assert len(summary["results"]) == 2
    method2 = next(item for item in summary["results"] if item["pairKey"].startswith("DB_METHOD2|"))
    method3 = next(item for item in summary["results"] if item["pairKey"] == "DB_METHOD3|ALL")

    assert method2["viewerGroup"] == "DB Direct"
    assert method2["viewerKind"] == "DB Method-2"
    assert method2["viewerGroupOrder"] == 41
    assert method3["viewerGroup"] == "DB Direct"
    assert method3["viewerKind"] == "DB Method-3"
    assert method3["viewerGroupOrder"] == 42
    assert len(summary["sourceCatalog"]) == 1
    assert summary["sourceCatalog"][0]["sourceKind"] == "db_single"
    assert [family["familyKey"] for family in summary["sourceCatalog"][0]["families"]] == ["db-motion", "db-layer-series"]
    chart_keys = [
        chart["chartKey"]
        for family in summary["sourceCatalog"][0]["families"]
        for chart in family["charts"]
    ]
    assert "db-layer-strain" in chart_keys
    assert "db-layer-response-spectrum" in chart_keys
    assert "db-layer-fourier" in chart_keys
    assert "db-layer-fourier-ratio" in chart_keys


def test_db_direct_pairs_time_history_schema_without_vel_disp():
    db_bytes = _build_time_history_db_bytes()
    summary = process_batch_files(
        {
            "Motion_DD2_X_RSN825_CAPEMEND_CPM000.db3": db_bytes,
            "Motion_DD2_Y_RSN825_CAPEMEND_CPM090.db3": db_bytes,
        },
        {
            "useDb3Directly": True,
            "method2Enabled": True,
            "method3Enabled": True,
            "_returnWebResults": True,
        },
    )

    assert summary["errors"] == []
    assert summary["metrics"]["pairsDetected"] == 1
    assert summary["metrics"]["method2Processed"] == 2
    assert summary["metrics"]["method3Produced"] == 1
    assert len(summary["sourceCatalog"]) == 1
    assert summary["sourceCatalog"][0]["sourceKind"] == "db_pair"
    family_keys = [family["familyKey"] for family in summary["sourceCatalog"][0]["families"]]
    assert family_keys == ["db-motion", "db-layer-series"]
    summary_entry = summary["summaryCatalog"][0]
    assert summary_entry["preferredVariantKey"] == "db_direct_total"
    total_variant = next(item for item in summary_entry["variants"] if item["variantKey"] == "db_direct_total")
    assert total_variant["values"] == pytest.approx([0.01697056274847714, 0.012727922061357855])
