import io

import disp_core


class _DummyExcelFile:
    def __init__(self, *_args, **_kwargs):
        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


def _prefetched_method2(name: str, axis: str) -> dict:
    return {
        "skipped": False,
        "axis": axis,
        "result": {
            "pairKey": f"METHOD2|{name}",
            "xFileName": name if axis == "X" else "",
            "yFileName": name if axis == "Y" else "",
            "outputFileName": f"output_method2_{name}.xlsx",
            "outputBytes": b"method2",
            "previewCharts": [],
            "metrics": {"mode": "method2_single", "axis": axis},
        },
    }


def test_batch_reuses_prefetched_method2_for_xlsx_pairs_and_singles(monkeypatch):
    file_map = {
        "pair_x.xlsx": b"x",
        "pair_y.xlsx": b"y",
        "single_x.xlsx": b"s",
    }

    monkeypatch.setattr(disp_core, "_candidate_kind", lambda _name: "xlsx")
    monkeypatch.setattr(disp_core, "_is_candidate_file", lambda _name, _include_manip: True)
    monkeypatch.setattr(
        disp_core,
        "_resolve_xy_pairs",
        lambda *_args, **_kwargs: ([("pair_x.xlsx", "pair_y.xlsx")], [], []),
    )
    monkeypatch.setattr(disp_core.pd, "ExcelFile", _DummyExcelFile)
    monkeypatch.setattr(
        disp_core,
        "_process_xy_pair_xlsx",
        lambda *_args, **_kwargs: {
            "result": {
                "pairKey": "pair",
                "xFileName": "pair_x.xlsx",
                "yFileName": "pair_y.xlsx",
                "outputFileName": "pair.xlsx",
                "outputBytes": b"pair",
                "previewCharts": [],
                "metrics": {"mode": "pair"},
            },
            "method2Extracted": {
                "pair_x.xlsx": _prefetched_method2("pair_x.xlsx", "X"),
                "pair_y.xlsx": _prefetched_method2("pair_y.xlsx", "Y"),
            },
        },
    )
    monkeypatch.setattr(
        disp_core,
        "_process_single_file_xlsx",
        lambda *_args, **_kwargs: {
            "result": {
                "pairKey": "single",
                "xFileName": "single_x.xlsx",
                "yFileName": "",
                "outputFileName": "single.xlsx",
                "outputBytes": b"single",
                "previewCharts": [],
                "metrics": {"mode": "single"},
            },
            "method2Extracted": _prefetched_method2("single_x.xlsx", "X"),
        },
    )
    monkeypatch.setattr(
        disp_core,
        "_extract_method2_single",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("unexpected fallback extraction")),
    )

    summary = disp_core.process_batch_files(
        file_map,
        {
            "method2Enabled": True,
            "method3Enabled": False,
        },
    )

    output_names = [item["outputFileName"] for item in summary["results"]]

    assert output_names == [
        "pair.xlsx",
        "single.xlsx",
        "output_method2_pair_x.xlsx.xlsx",
        "output_method2_pair_y.xlsx.xlsx",
        "output_method2_single_x.xlsx.xlsx",
    ]
    assert summary["metrics"]["method2Processed"] == 3
