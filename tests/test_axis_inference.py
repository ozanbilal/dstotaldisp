from disp_core import _derive_y_name, _infer_axis_label


def test_infer_axis_label_handles_horizontal_numbered_names():
    assert _infer_axis_label("Results_profile_Profile 1_motion_Yalova_Horizontal_1_Matched_RSN864.xlsx") == "X"
    assert _infer_axis_label("Results_profile_Profile 1_motion_Yalova_Horiontal_2_Matched_RSN864.xlsx") == "Y"


def test_derive_y_name_handles_horizontal_numbered_names():
    assert (
        _derive_y_name("Results_profile_Profile 1_motion_Yalova_Horizontal_1_Matched_RSN864.xlsx")
        == "Results_profile_Profile 1_motion_Yalova_Horizontal_2_Matched_RSN864.xlsx"
    )
    assert (
        _derive_y_name("Results_profile_Profile 1_motion_Yalova_Horiontal_1_Matched_RSN864.xlsx")
        == "Results_profile_Profile 1_motion_Yalova_Horiontal_2_Matched_RSN864.xlsx"
    )
