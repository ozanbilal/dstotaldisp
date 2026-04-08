from pathlib import Path

from disp_core import process_batch_files


def test_pair_source_catalog_exposes_file_and_pair_entries():
    fixture = Path(__file__).resolve().parent.parent / "_tmp_m23_case_pair_in"
    file_map = {path.name: path.read_bytes() for path in fixture.glob("*.xlsx")}

    summary = process_batch_files(
        file_map,
        {
            "_returnWebResults": True,
            "method2Enabled": True,
            "method3Enabled": True,
        },
    )

    assert len(summary["results"]) == 4
    assert all(item.get("outputBytesB64") for item in summary["results"])

    source_catalog = summary["sourceCatalog"]
    assert len(source_catalog) == 4

    by_kind = {}
    for entry in source_catalog:
        by_kind.setdefault(entry["sourceKind"], []).append(entry)

    assert len(by_kind["single"]) == 2
    assert len(by_kind["pair"]) == 1
    assert len(by_kind["method3_aggregate"]) == 1

    x_source = next(entry for entry in by_kind["single"] if entry["axis"] == "X")
    y_source = next(entry for entry in by_kind["single"] if entry["axis"] == "Y")
    pair_source = by_kind["pair"][0]

    assert [family["familyKey"] for family in x_source["families"]] == [
        "input-motion",
        "profile",
        "layer-series",
        "derived-profiles",
    ]
    assert [family["familyKey"] for family in y_source["families"]] == [
        "input-motion",
        "profile",
        "layer-series",
        "derived-profiles",
    ]
    assert [family["familyKey"] for family in pair_source["families"]] == [
        "input-motion",
        "profile",
        "layer-series",
        "derived-profiles",
    ]

    x_input_motion = next(family for family in x_source["families"] if family["familyKey"] == "input-motion")
    x_input_motion_keys = [chart["chartKey"] for chart in x_input_motion["charts"]]
    assert "input-psa" in x_input_motion_keys

    x_derived = next(family for family in x_source["families"] if family["familyKey"] == "derived-profiles")
    assert [chart["chartKey"] for chart in x_derived["charts"]] == ["derived-total-profile"]

    pair_derived = next(family for family in pair_source["families"] if family["familyKey"] == "derived-profiles")
    pair_derived_keys = [chart["chartKey"] for chart in pair_derived["charts"]]
    assert pair_derived_keys == ["derived-x-profile", "derived-y-profile", "derived-resultant-profile"]
    assert all("relative" not in chart["chartLabel"].lower() for chart in pair_derived["charts"])

    assert x_source["artifactPairKeys"] == ["METHOD2|Results_profile_1_motion_DD1_X_20030501002708_1201_H1"]
    assert y_source["artifactPairKeys"] == ["METHOD2|Results_profile_1_motion_DD1_Y_20030501002708_1201_H2"]
    assert pair_source["artifactPairKeys"] == [
        "Results_profile_1_motion_DD1_20030501002708_1201|Results_profile_1_motion_DD1_Y_20030501002708_1201_H2"
    ]
