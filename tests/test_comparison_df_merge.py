import pandas as pd

from disp_core import _build_comparison_df


def test_build_comparison_df_preserves_layers_when_depth_values_drift():
    strain_df = pd.DataFrame(
        {
            "Layer_Index": [1, 2, 3],
            "Depth_m": [0.5, 1.5, 2.5],
            "X_base_rel_max_m": [0.1, 0.2, 0.3],
            "Y_base_rel_max_m": [0.1, 0.2, 0.3],
            "X_tbdy_total_max_m": [0.15, 0.25, 0.35],
            "Y_tbdy_total_max_m": [0.15, 0.25, 0.35],
            "Total_base_rel_max_m": [0.14, 0.28, 0.42],
            "Total_tbdy_total_max_m": [0.2, 0.32, 0.46],
            "Total_input_proxy_rel_max_m": [0.18, 0.3, 0.44],
            "Total_profile_offset_total_est_m": [0.19, 0.31, 0.45],
        }
    )
    legacy_df = pd.DataFrame(
        {
            "Layer_Index": [1, 2, 3],
            "Depth_m": [0.5000001, 1.5000001, 2.5000001],
            "Profile_X_max_m": [0.11, 0.21, 0.31],
            "Profile_Y_max_m": [0.11, 0.21, 0.31],
            "Profile_RSS_total_m": [0.16, 0.29, 0.43],
            "TimeHist_Resultant_total_m": [0.17, 0.3, 0.44],
        }
    )

    merged = _build_comparison_df(strain_df, legacy_df)

    assert len(merged) == 3
    assert merged["Layer_Index"].tolist() == [1, 2, 3]
    assert merged["Depth_m"].tolist() == [0.5, 1.5, 2.5]
    assert merged["Profile_RSS_total_m"].tolist() == [0.16, 0.29, 0.43]
