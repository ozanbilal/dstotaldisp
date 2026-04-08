from pathlib import Path

from disp_core import process_batch_files


ROOT = Path(__file__).resolve().parents[1]


def _load_inputs(folder: str) -> dict[str, bytes]:
    base = ROOT / folder
    return {item.name: item.read_bytes() for item in base.iterdir() if item.is_file()}


def test_pair_batch_results_include_preview_charts():
    summary = process_batch_files(
        _load_inputs("_tmp_m23_case_pair_in"),
        {
            "method2Enabled": True,
            "method3Enabled": True,
            "_returnWebResults": True,
        },
    )

    preview_counts = {result["pairKey"]: len(result.get("previewCharts", [])) for result in summary["results"]}
    primary_result = next((result for result in summary["results"] if result.get("viewerGroup") == "Primary Outputs"), None)

    assert primary_result is not None
    assert len(primary_result.get("previewCharts", [])) > 0
    assert preview_counts["METHOD2|Results_profile_1_motion_DD1_X_20030501002708_1201_H1"] > 0
    assert preview_counts["METHOD2|Results_profile_1_motion_DD1_Y_20030501002708_1201_H2"] > 0
    assert preview_counts["METHOD3|ALL"] > 0

    viewer_groups = {result["pairKey"]: result["viewerGroup"] for result in summary["results"]}
    assert primary_result.get("viewerGroup") == "Primary Outputs"
    assert viewer_groups["METHOD2|Results_profile_1_motion_DD1_X_20030501002708_1201_H1"] == "Method-2"
    assert viewer_groups["METHOD3|ALL"] == "Method-3 Aggregate"

    preview_titles = {
        result["pairKey"]: [chart.get("title", "") for chart in result.get("previewCharts", [])]
        for result in summary["results"]
    }
    assert "Method-2 X Approx Total (Ubase + Urel)" not in preview_titles[
        "METHOD2|Results_profile_1_motion_DD1_X_20030501002708_1201_H1"
    ]
    assert "Method-3 X Approx Total (Ubase + Urel)" in preview_titles["METHOD3|ALL"]
