from io import BytesIO

import pandas as pd
from openpyxl import Workbook

from disp_core import _read_layer_column


def test_read_layer_column_accepts_shear_strain_alias():
    wb = Workbook()
    ws = wb.active
    ws.title = "Layer 1"
    ws.append(["Time (s)", "Acceleration (g)", "Shear Strain (%)"])
    ws.append([0.0, 0.0, 1.25])
    ws.append([0.1, 0.1, 2.50])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    with pd.ExcelFile(buffer, engine="openpyxl") as xl:
        time, strain = _read_layer_column(xl, "Layer 1", "Strain (%)")

    assert time.tolist() == [0.0, 0.1]
    assert strain.tolist() == [1.25, 2.5]
