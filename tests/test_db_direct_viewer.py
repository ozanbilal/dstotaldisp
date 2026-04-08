import sqlite3
import tempfile
from pathlib import Path

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
    assert summary["sourceCatalog"][0]["sourceKind"] == "db_pair"
    assert [family["familyKey"] for family in summary["sourceCatalog"][0]["families"]] == ["db-motion"]
