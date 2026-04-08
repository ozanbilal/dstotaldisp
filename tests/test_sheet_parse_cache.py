import io
from pathlib import Path

import pandas as pd

from disp_core import _read_layer_column


ROOT = Path(__file__).resolve().parents[1]


def test_read_layer_column_reuses_cached_sheet_parse():
    workbook = next((ROOT / "_tmp_m23_case_pair_in").glob("*.xlsx"))

    with pd.ExcelFile(io.BytesIO(workbook.read_bytes()), engine="openpyxl") as xl:
        target_sheet = next(name for name in xl.sheet_names if name.startswith("Layer"))
        parse_calls = 0
        original_parse = xl.parse

        def counting_parse(*args, **kwargs):
            nonlocal parse_calls
            parse_calls += 1
            return original_parse(*args, **kwargs)

        xl.parse = counting_parse

        _read_layer_column(xl, target_sheet, "Strain (%)")
        _read_layer_column(xl, target_sheet, "Acceleration (g)")

        assert parse_calls == 1
