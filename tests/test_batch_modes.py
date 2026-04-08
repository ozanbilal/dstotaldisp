from pathlib import Path

from disp_core import process_batch_files


ROOT = Path(__file__).resolve().parents[1]


def _load_inputs(folder: str) -> dict[str, bytes]:
    base = ROOT / folder
    return {item.name: item.read_bytes() for item in base.iterdir() if item.is_file()}


def test_process_batch_files_supports_primary_only_and_method23_only_modes():
    files = _load_inputs("_tmp_m23_case_pair_in")

    primary_only = process_batch_files(
        files,
        {
            "method2Enabled": True,
            "method3Enabled": True,
            "skipMethod23Outputs": True,
        },
    )
    method23_only = process_batch_files(
        files,
        {
            "method2Enabled": True,
            "method3Enabled": True,
            "skipPrimaryOutputs": True,
        },
    )

    primary_keys = [item["pairKey"] for item in primary_only["results"]]
    method23_keys = [item["pairKey"] for item in method23_only["results"]]

    assert primary_keys == [
        "Results_profile_1_motion_DD1_20030501002708_1201|Results_profile_1_motion_DD1_Y_20030501002708_1201_H2"
    ]
    assert method23_keys == [
        "METHOD2|Results_profile_1_motion_DD1_X_20030501002708_1201_H1",
        "METHOD2|Results_profile_1_motion_DD1_Y_20030501002708_1201_H2",
        "METHOD3|ALL",
    ]


def test_process_batch_files_can_return_web_ready_results():
    files = _load_inputs("_tmp_m23_case_pair_in")

    summary = process_batch_files(
        files,
        {
            "method2Enabled": True,
            "method3Enabled": True,
            "_returnWebResults": True,
        },
    )

    assert summary["results"]
    assert all("outputBytesB64" in item for item in summary["results"])
    assert all("outputBytes" not in item for item in summary["results"])
    assert all("viewerGroup" in item for item in summary["results"])
    assert all("viewerKind" in item for item in summary["results"])
    assert all("viewerGroupOrder" in item for item in summary["results"])

    viewer_groups = {item["pairKey"]: (item["viewerGroup"], item["viewerKind"], item["viewerGroupOrder"]) for item in summary["results"]}
    primary_items = [item for item in summary["results"] if item.get("viewerGroup") == "Primary Outputs"]
    assert len(primary_items) == 1
    assert (
        primary_items[0].get("viewerGroup"),
        primary_items[0].get("viewerKind"),
        primary_items[0].get("viewerGroupOrder"),
    ) == ("Primary Outputs", "Pair", 10)
    assert viewer_groups["METHOD2|Results_profile_1_motion_DD1_X_20030501002708_1201_H1"] == ("Method-2", "Method-2 X", 20)
    assert viewer_groups["METHOD2|Results_profile_1_motion_DD1_Y_20030501002708_1201_H2"] == ("Method-2", "Method-2 Y", 20)
    assert viewer_groups["METHOD3|ALL"] == ("Method-3 Aggregate", "Method-3", 30)
