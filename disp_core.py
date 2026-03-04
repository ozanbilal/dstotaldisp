import io
import math
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Mapping, Sequence, Tuple

import numpy as np
import pandas as pd
from openpyxl.chart import LineChart, Reference, ScatterChart, Series


EXCLUDE_PREFIXES = ("output_", "~$")
EXCLUDE_SUFFIXES = ("-manip.xlsx",)
DEFAULT_HIGHPASS_CUTOFF_HZ = 0.03
DEFAULT_HIGHPASS_TRANSITION_HZ = 0.02


def _log(logs: List[Dict[str, str]], level: str, message: str) -> None:
    logs.append({"level": level, "message": message})


def _to_number_series(series: pd.Series) -> np.ndarray:
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.dropna().to_numpy(dtype=float)


def _cumtrapz(y: np.ndarray, x: np.ndarray) -> np.ndarray:
    if y.size == 0:
        return np.array([], dtype=float)
    if y.size == 1:
        return np.array([0.0], dtype=float)
    dx = np.diff(x)
    area = 0.5 * (y[1:] + y[:-1]) * dx
    return np.concatenate(([0.0], np.cumsum(area)))


def _baseline_correct(acc: np.ndarray, time: np.ndarray) -> np.ndarray:
    if acc.size < 4:
        return acc - np.mean(acc)
    coeff = np.polyfit(time, acc, 3)
    baseline = np.polyval(coeff, time)
    return acc - baseline


def _soft_highpass_fft(
    signal: np.ndarray,
    time: np.ndarray,
    cutoff_hz: float = DEFAULT_HIGHPASS_CUTOFF_HZ,
    transition_hz: float = DEFAULT_HIGHPASS_TRANSITION_HZ,
) -> np.ndarray:
    x = signal.astype(float)
    if x.size < 8:
        return x - np.mean(x)

    dt = float(np.median(np.diff(time.astype(float))))
    if not np.isfinite(dt) or dt <= 0:
        return x - np.mean(x)

    x = x - np.mean(x)
    n = x.size
    freqs = np.fft.rfftfreq(n, d=dt)
    if freqs.size <= 1:
        return x

    nyquist = 0.5 / dt
    cutoff = float(np.clip(cutoff_hz, 0.0, max(0.0, nyquist * 0.999)))
    transition = max(0.0, float(transition_hz))
    stop = max(0.0, cutoff - transition)

    if cutoff <= 0.0:
        return x

    transfer = np.ones_like(freqs)
    transfer[freqs <= stop] = 0.0

    if cutoff > stop:
        mask = (freqs > stop) & (freqs < cutoff)
        xi = (freqs[mask] - stop) / (cutoff - stop)
        transfer[mask] = 0.5 - 0.5 * np.cos(np.pi * xi)

    spectrum = np.fft.rfft(x)
    filtered = np.fft.irfft(spectrum * transfer, n=n)
    return filtered.astype(float)


def _acc_to_disp(
    time: np.ndarray,
    acc_g: np.ndarray,
    *,
    highpass_enabled: bool = True,
    highpass_cutoff_hz: float = DEFAULT_HIGHPASS_CUTOFF_HZ,
    highpass_transition_hz: float = DEFAULT_HIGHPASS_TRANSITION_HZ,
) -> np.ndarray:
    acc_corr = _baseline_correct(acc_g.astype(float), time.astype(float))
    if highpass_enabled:
        acc_proc = _soft_highpass_fft(
            acc_corr,
            time,
            cutoff_hz=highpass_cutoff_hz,
            transition_hz=highpass_transition_hz,
        )
    else:
        acc_proc = acc_corr
    vel = _cumtrapz(acc_proc * 9.81, time)
    return _cumtrapz(vel, time)


def _normalize_options(options: Any) -> Dict[str, Any]:
    if options is None:
        return {}
    if hasattr(options, "to_py"):
        options = options.to_py()
    if isinstance(options, dict):
        return dict(options)
    return {}


def _to_float(value: Any, default: float) -> float:
    try:
        out = float(value)
    except (TypeError, ValueError):
        return default
    return out if np.isfinite(out) else default


def _to_bool(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        v = value.strip().lower()
        if v in {"1", "true", "yes", "on"}:
            return True
        if v in {"0", "false", "no", "off"}:
            return False
    return default


def _highpass_config(options: Mapping[str, Any] | None) -> Tuple[bool, float, float]:
    cfg = options or {}
    enabled = _to_bool(cfg.get("highpassEnabled", True), True)
    cutoff = max(0.0, _to_float(cfg.get("highpassCutoffHz"), DEFAULT_HIGHPASS_CUTOFF_HZ))
    transition = max(0.0, _to_float(cfg.get("highpassTransitionHz"), DEFAULT_HIGHPASS_TRANSITION_HZ))
    return enabled, cutoff, transition


def _layer_sort_key(name: str) -> Tuple[int, str]:
    match = re.search(r"(\d+)$", name)
    if match:
        return (int(match.group(1)), name)
    return (10**9, name)


def _list_layer_sheets(xl: pd.ExcelFile) -> List[str]:
    sheets = [name for name in xl.sheet_names if name.startswith("Layer")]
    return sorted(sheets, key=_layer_sort_key)


def _ensure_common_layers(x_xl: pd.ExcelFile, y_xl: pd.ExcelFile) -> List[str]:
    x_layers = _list_layer_sheets(x_xl)
    y_layers = _list_layer_sheets(y_xl)
    if not x_layers or not y_layers:
        raise ValueError("Missing Layer sheets in one or both files.")
    if x_layers != y_layers:
        raise ValueError("Layer sheet sets are not identical between X and Y files.")
    return x_layers


def _read_layer_column(xl: pd.ExcelFile, layer_name: str, value_column: str) -> Tuple[np.ndarray, np.ndarray]:
    df = xl.parse(layer_name)
    if "Time (s)" not in df.columns or value_column not in df.columns:
        raise ValueError(f"Sheet '{layer_name}' is missing 'Time (s)' or '{value_column}'.")

    data = df[["Time (s)", value_column]].copy()
    data["Time (s)"] = pd.to_numeric(data["Time (s)"], errors="coerce")
    data[value_column] = pd.to_numeric(data[value_column], errors="coerce")
    data = data.dropna(subset=["Time (s)", value_column]).sort_values("Time (s)")

    if data.empty:
        raise ValueError(f"Sheet '{layer_name}' has no numeric rows for '{value_column}'.")

    return data["Time (s)"].to_numpy(dtype=float), data[value_column].to_numpy(dtype=float)


def _align_two_series(
    time_x: np.ndarray,
    value_x: np.ndarray,
    time_y: np.ndarray,
    value_y: np.ndarray,
) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    if time_x.size == time_y.size and np.allclose(time_x, time_y):
        return time_x, value_x, value_y

    start = max(float(time_x[0]), float(time_y[0]))
    end = min(float(time_x[-1]), float(time_y[-1]))
    if end <= start:
        raise ValueError("X and Y time windows do not overlap.")

    dx = np.median(np.diff(time_x)) if time_x.size > 1 else 0.01
    dy = np.median(np.diff(time_y)) if time_y.size > 1 else 0.01
    dt = min(dx, dy)
    if not np.isfinite(dt) or dt <= 0:
        dt = 0.01

    count = int(math.floor((end - start) / dt)) + 1
    common_time = start + np.arange(count, dtype=float) * dt

    x_interp = np.interp(common_time, time_x, value_x)
    y_interp = np.interp(common_time, time_y, value_y)
    return common_time, x_interp, y_interp


def parse_profile_thickness(xl: pd.ExcelFile) -> Tuple[np.ndarray, np.ndarray]:
    if "Profile" not in xl.sheet_names:
        raise ValueError("Missing 'Profile' sheet.")

    profile = xl.parse("Profile")
    data = profile.iloc[1:].reset_index(drop=True)
    if data.empty:
        raise ValueError("Profile sheet has no numeric content.")

    mid_col = "Effective Stress" if "Effective Stress" in data.columns else data.columns[0]
    depth_col = "Maximum Displacement" if "Maximum Displacement" in data.columns else data.columns[6]

    mid_depths = _to_number_series(data[mid_col])
    out_depths = _to_number_series(data[depth_col])

    n = min(mid_depths.size, out_depths.size)
    if n == 0:
        raise ValueError("Unable to parse profile depths from Profile sheet.")

    mid_depths = mid_depths[:n]
    out_depths = out_depths[:n]

    thickness = np.zeros(n, dtype=float)
    cumulative = 0.0
    for i, depth_mid in enumerate(mid_depths):
        if i == 0:
            h = 2.0 * depth_mid
        else:
            h = 2.0 * (depth_mid - cumulative)
        if h <= 0:
            raise ValueError("Non-positive layer thickness detected while parsing profile.")
        thickness[i] = h
        cumulative += h

    predicted_depths = np.concatenate(([0.0], np.cumsum(thickness)))[:-1]
    if predicted_depths.size == out_depths.size and np.max(np.abs(predicted_depths - out_depths)) > 1e-3:
        raise ValueError("Profile depth and inferred thickness are inconsistent.")

    return out_depths, thickness


def _parse_profile_displacement_max(xl: pd.ExcelFile) -> Tuple[np.ndarray, np.ndarray]:
    if "Profile" not in xl.sheet_names:
        raise ValueError("Missing 'Profile' sheet.")

    profile = xl.parse("Profile")
    data = profile.iloc[1:].reset_index(drop=True)
    if data.empty:
        raise ValueError("Profile sheet has no numeric content.")

    depth_col = "Maximum Displacement" if "Maximum Displacement" in data.columns else data.columns[6]
    depth_idx = data.columns.get_loc(depth_col)
    disp_idx = min(depth_idx + 1, len(data.columns) - 1)

    depths = _to_number_series(data.iloc[:, depth_idx])
    max_disp = _to_number_series(data.iloc[:, disp_idx])

    n = min(depths.size, max_disp.size)
    if n == 0:
        raise ValueError("Unable to parse Profile maximum displacement columns.")

    return depths[:n], max_disp[:n]


def _read_input_motion(xl: pd.ExcelFile) -> Tuple[np.ndarray, np.ndarray]:
    if "Input Motion" not in xl.sheet_names:
        raise ValueError("Missing 'Input Motion' sheet.")

    motion = xl.parse("Input Motion")
    if motion.shape[1] < 2:
        raise ValueError("Input Motion sheet has fewer than two columns.")

    subset = motion.iloc[:, :2].copy()
    subset.columns = ["Time (s)", "Acceleration (g)"]
    subset["Time (s)"] = pd.to_numeric(subset["Time (s)"], errors="coerce")
    subset["Acceleration (g)"] = pd.to_numeric(subset["Acceleration (g)"], errors="coerce")
    subset = subset.dropna(subset=["Time (s)", "Acceleration (g)"]).sort_values("Time (s)")

    if subset.empty:
        raise ValueError("Input Motion sheet does not contain numeric time/acceleration rows.")

    return (
        subset["Time (s)"].to_numpy(dtype=float),
        subset["Acceleration (g)"].to_numpy(dtype=float),
    )


def _compute_strain_bundle(
    x_xl: pd.ExcelFile,
    y_xl: pd.ExcelFile,
    options: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    layer_names = _ensure_common_layers(x_xl, y_xl)

    x_depths, thickness = parse_profile_thickness(x_xl)
    y_depths, _ = parse_profile_thickness(y_xl)
    if x_depths.size != y_depths.size or not np.allclose(x_depths, y_depths):
        raise ValueError("Profile depths are inconsistent between X and Y files.")

    n_layers = min(len(layer_names), x_depths.size, thickness.size)
    layer_names = layer_names[:n_layers]
    depths = x_depths[:n_layers]
    thickness = thickness[:n_layers]

    layer_payload: List[Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]] = []
    t_start = -np.inf
    t_end = np.inf
    dt_min = np.inf

    for layer_name in layer_names:
        tx, strain_x_pct = _read_layer_column(x_xl, layer_name, "Strain (%)")
        ty, strain_y_pct = _read_layer_column(y_xl, layer_name, "Strain (%)")
        gamma_x = strain_x_pct / 100.0
        gamma_y = strain_y_pct / 100.0

        t_start = max(t_start, tx[0], ty[0])
        t_end = min(t_end, tx[-1], ty[-1])
        if tx.size > 1:
            dt_min = min(dt_min, float(np.median(np.diff(tx))))
        if ty.size > 1:
            dt_min = min(dt_min, float(np.median(np.diff(ty))))

        layer_payload.append((tx, gamma_x, ty, gamma_y))

    if t_end <= t_start:
        raise ValueError("No overlapping time window found across layer strain records.")

    if not np.isfinite(dt_min) or dt_min <= 0:
        dt_min = 0.01

    sample_count = int(math.floor((t_end - t_start) / dt_min)) + 1
    time = t_start + np.arange(sample_count, dtype=float) * dt_min

    gamma_x_matrix = np.zeros((n_layers, sample_count), dtype=float)
    gamma_y_matrix = np.zeros((n_layers, sample_count), dtype=float)

    for i, (tx, gx, ty, gy) in enumerate(layer_payload):
        gamma_x_matrix[i, :] = np.interp(time, tx, gx)
        gamma_y_matrix[i, :] = np.interp(time, ty, gy)

    du_x = gamma_x_matrix * thickness[:, None]
    du_y = gamma_y_matrix * thickness[:, None]

    u_rel_base_x = np.flip(np.cumsum(np.flip(du_x, axis=0), axis=0), axis=0)
    u_rel_base_y = np.flip(np.cumsum(np.flip(du_y, axis=0), axis=0), axis=0)

    t_input_x, a_input_x = _read_input_motion(x_xl)
    t_input_y, a_input_y = _read_input_motion(y_xl)

    a_input_x_i = np.interp(time, t_input_x, a_input_x)
    a_input_y_i = np.interp(time, t_input_y, a_input_y)

    hp_enabled, hp_cutoff, hp_transition = _highpass_config(options)
    u_input_proxy_x = _acc_to_disp(
        time,
        a_input_x_i,
        highpass_enabled=hp_enabled,
        highpass_cutoff_hz=hp_cutoff,
        highpass_transition_hz=hp_transition,
    )
    u_input_proxy_y = _acc_to_disp(
        time,
        a_input_y_i,
        highpass_enabled=hp_enabled,
        highpass_cutoff_hz=hp_cutoff,
        highpass_transition_hz=hp_transition,
    )

    u_rel_input_x = u_rel_base_x - u_input_proxy_x[None, :]
    u_rel_input_y = u_rel_base_y - u_input_proxy_y[None, :]
    u_tbdy_total_x = u_rel_base_x + u_input_proxy_x[None, :]
    u_tbdy_total_y = u_rel_base_y + u_input_proxy_y[None, :]

    x_base = np.max(np.abs(u_rel_base_x), axis=1)
    y_base = np.max(np.abs(u_rel_base_y), axis=1)
    total_base = np.max(np.sqrt(u_rel_base_x**2 + u_rel_base_y**2), axis=1)

    x_tbdy_total = np.max(np.abs(u_tbdy_total_x), axis=1)
    y_tbdy_total = np.max(np.abs(u_tbdy_total_y), axis=1)
    total_tbdy_total = np.max(np.sqrt(u_tbdy_total_x**2 + u_tbdy_total_y**2), axis=1)

    x_input = np.max(np.abs(u_rel_input_x), axis=1)
    y_input = np.max(np.abs(u_rel_input_y), axis=1)
    total_input = np.max(np.sqrt(u_rel_input_x**2 + u_rel_input_y**2), axis=1)

    summary_df = pd.DataFrame(
        {
            "Layer_Index": np.arange(1, n_layers + 1, dtype=int),
            "Depth_m": depths,
            "Thickness_m": thickness,
            "X_base_rel_max_m": x_base,
            "Y_base_rel_max_m": y_base,
            "Total_base_rel_max_m": total_base,
            "X_tbdy_total_max_m": x_tbdy_total,
            "Y_tbdy_total_max_m": y_tbdy_total,
            "Total_tbdy_total_max_m": total_tbdy_total,
            "X_input_proxy_rel_max_m": x_input,
            "Y_input_proxy_rel_max_m": y_input,
            "Total_input_proxy_rel_max_m": total_input,
        }
    )

    return {
        "layer_names": layer_names,
        "depths": depths,
        "thickness": thickness,
        "time": time,
        "u_rel_base_x": u_rel_base_x,
        "u_rel_base_y": u_rel_base_y,
        "u_input_proxy_x": u_input_proxy_x,
        "u_input_proxy_y": u_input_proxy_y,
        "u_rel_input_x": u_rel_input_x,
        "u_rel_input_y": u_rel_input_y,
        "u_tbdy_total_x": u_tbdy_total_x,
        "u_tbdy_total_y": u_tbdy_total_y,
        "summary_df": summary_df,
    }


def compute_strain_relative(
    x_xl: pd.ExcelFile,
    y_xl: pd.ExcelFile,
    options: Mapping[str, Any] | None = None,
) -> pd.DataFrame:
    bundle = _compute_strain_bundle(x_xl, y_xl, options)
    return bundle["summary_df"].copy()


def _build_layer_time_df(
    time: np.ndarray,
    depths: np.ndarray,
    matrix: np.ndarray,
    value_suffix: str,
) -> pd.DataFrame:
    n_layers = min(int(matrix.shape[0]), int(depths.size))
    data: Dict[str, np.ndarray] = {"Time_s": time}
    for i in range(n_layers):
        depth = float(depths[i])
        data[f"L{i + 1:02d}_z{depth:.3f}m_{value_suffix}"] = matrix[i]
    return pd.DataFrame(data)


def _compute_single_strain_bundle(
    xl: pd.ExcelFile,
    axis_label: str,
    options: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    layer_names = _list_layer_sheets(xl)
    if not layer_names:
        raise ValueError(f"Missing Layer sheets in {axis_label} file.")

    depths, thickness = parse_profile_thickness(xl)
    n_layers = min(len(layer_names), depths.size, thickness.size)
    layer_names = layer_names[:n_layers]
    depths = depths[:n_layers]
    thickness = thickness[:n_layers]

    payload: List[Tuple[np.ndarray, np.ndarray]] = []
    t_start = -np.inf
    t_end = np.inf
    dt_min = np.inf

    for layer_name in layer_names:
        t, strain_pct = _read_layer_column(xl, layer_name, "Strain (%)")
        gamma = strain_pct / 100.0
        payload.append((t, gamma))

        t_start = max(t_start, t[0])
        t_end = min(t_end, t[-1])
        if t.size > 1:
            dt_min = min(dt_min, float(np.median(np.diff(t))))

    if t_end <= t_start:
        raise ValueError(f"No overlapping time window found across strain records in {axis_label} file.")

    if not np.isfinite(dt_min) or dt_min <= 0:
        dt_min = 0.01

    sample_count = int(math.floor((t_end - t_start) / dt_min)) + 1
    time = t_start + np.arange(sample_count, dtype=float) * dt_min

    gamma_matrix = np.zeros((n_layers, sample_count), dtype=float)
    for i, (t, gamma) in enumerate(payload):
        gamma_matrix[i, :] = np.interp(time, t, gamma)

    du = gamma_matrix * thickness[:, None]
    u_rel_base = np.flip(np.cumsum(np.flip(du, axis=0), axis=0), axis=0)

    t_input, a_input = _read_input_motion(xl)
    a_input_i = np.interp(time, t_input, a_input)
    hp_enabled, hp_cutoff, hp_transition = _highpass_config(options)
    u_input_proxy = _acc_to_disp(
        time,
        a_input_i,
        highpass_enabled=hp_enabled,
        highpass_cutoff_hz=hp_cutoff,
        highpass_transition_hz=hp_transition,
    )

    u_rel_input = u_rel_base - u_input_proxy[None, :]
    u_tbdy_total = u_rel_base + u_input_proxy[None, :]

    base_rel_max = np.max(np.abs(u_rel_base), axis=1)
    tbdy_total_max = np.max(np.abs(u_tbdy_total), axis=1)
    input_proxy_rel_max = np.max(np.abs(u_rel_input), axis=1)

    summary_df = pd.DataFrame(
        {
            "Layer_Index": np.arange(1, n_layers + 1, dtype=int),
            "Depth_m": depths,
            "Thickness_m": thickness,
            "Axis": axis_label,
            "Base_rel_max_m": base_rel_max,
            "TBDY_total_max_m": tbdy_total_max,
            "Input_proxy_rel_max_m": input_proxy_rel_max,
        }
    )

    return {
        "layer_names": layer_names,
        "depths": depths,
        "thickness": thickness,
        "time": time,
        "u_rel_base": u_rel_base,
        "u_rel_input": u_rel_input,
        "u_tbdy_total": u_tbdy_total,
        "summary_df": summary_df,
    }


def _compute_single_direction_disp_bundle(
    xl: pd.ExcelFile,
    axis_label: str,
    options: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    layer_names = _list_layer_sheets(xl)
    if not layer_names:
        raise ValueError(f"Missing Layer sheets in {axis_label} file.")

    depths, _ = parse_profile_thickness(xl)
    n_layers = min(len(layer_names), depths.size)
    layer_names = layer_names[:n_layers]
    depths = depths[:n_layers]

    payload: List[Tuple[np.ndarray, np.ndarray]] = []
    t_start = -np.inf
    t_end = np.inf
    dt_min = np.inf
    hp_enabled, hp_cutoff, hp_transition = _highpass_config(options)

    for layer_name in layer_names:
        t, a = _read_layer_column(xl, layer_name, "Acceleration (g)")
        d = _acc_to_disp(
            t,
            a,
            highpass_enabled=hp_enabled,
            highpass_cutoff_hz=hp_cutoff,
            highpass_transition_hz=hp_transition,
        )
        payload.append((t, d))

        t_start = max(t_start, t[0])
        t_end = min(t_end, t[-1])
        if t.size > 1:
            dt_min = min(dt_min, float(np.median(np.diff(t))))

    if t_end <= t_start:
        common_time = payload[0][0]
    else:
        if not np.isfinite(dt_min) or dt_min <= 0:
            dt_min = 0.01
        sample_count = int(math.floor((t_end - t_start) / dt_min)) + 1
        common_time = t_start + np.arange(sample_count, dtype=float) * dt_min

    disp_matrix = np.zeros((n_layers, common_time.size), dtype=float)
    for i, (t, d) in enumerate(payload):
        disp_matrix[i, :] = np.interp(common_time, t, d)

    table_df = _build_layer_time_df(common_time, depths, disp_matrix, "disp_m")

    return {
        "axis": axis_label,
        "layer_names": layer_names,
        "depths": depths,
        "time": common_time,
        "disp_matrix": disp_matrix,
        "table_df": table_df,
    }


def _build_resultant_time_df(
    x_bundle: Mapping[str, Any],
    y_bundle: Mapping[str, Any],
) -> pd.DataFrame:
    n_layers = min(
        int(x_bundle["disp_matrix"].shape[0]),
        int(y_bundle["disp_matrix"].shape[0]),
        int(x_bundle["depths"].size),
        int(y_bundle["depths"].size),
    )
    if n_layers <= 0:
        return pd.DataFrame({"Time_s": []})

    t = _get_common_time_for_layer(x_bundle["time"], y_bundle["time"])
    data: Dict[str, np.ndarray] = {"Time_s": t}

    for i in range(n_layers):
        x_i = np.interp(t, x_bundle["time"], x_bundle["disp_matrix"][i])
        y_i = np.interp(t, y_bundle["time"], y_bundle["disp_matrix"][i])
        total = np.sqrt(x_i**2 + y_i**2)
        depth = float(x_bundle["depths"][i])
        data[f"L{i + 1:02d}_z{depth:.3f}m_resultant_m"] = total

    return pd.DataFrame(data)


def _compute_legacy_bundle(
    x_xl: pd.ExcelFile,
    y_xl: pd.ExcelFile,
    options: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    layer_names = _ensure_common_layers(x_xl, y_xl)

    depth_x, profile_x = _parse_profile_displacement_max(x_xl)
    depth_y, profile_y = _parse_profile_displacement_max(y_xl)

    if depth_x.size != depth_y.size or not np.allclose(depth_x, depth_y):
        raise ValueError("Profile maximum displacement depths mismatch between X and Y files.")

    n_layers = min(len(layer_names), depth_x.size, profile_x.size, profile_y.size)
    layer_names = layer_names[:n_layers]
    depths = depth_x[:n_layers]
    profile_x = np.abs(profile_x[:n_layers])
    profile_y = np.abs(profile_y[:n_layers])

    time_hist_x = np.zeros(n_layers, dtype=float)
    time_hist_y = np.zeros(n_layers, dtype=float)
    time_hist_resultant = np.zeros(n_layers, dtype=float)

    hp_enabled, hp_cutoff, hp_transition = _highpass_config(options)

    for i, layer_name in enumerate(layer_names):
        tx, ax = _read_layer_column(x_xl, layer_name, "Acceleration (g)")
        ty, ay = _read_layer_column(y_xl, layer_name, "Acceleration (g)")
        t, ax_i, ay_i = _align_two_series(tx, ax, ty, ay)

        dx = _acc_to_disp(
            t,
            ax_i,
            highpass_enabled=hp_enabled,
            highpass_cutoff_hz=hp_cutoff,
            highpass_transition_hz=hp_transition,
        )
        dy = _acc_to_disp(
            t,
            ay_i,
            highpass_enabled=hp_enabled,
            highpass_cutoff_hz=hp_cutoff,
            highpass_transition_hz=hp_transition,
        )
        total = np.sqrt(dx**2 + dy**2)

        time_hist_x[i] = float(np.max(np.abs(dx)))
        time_hist_y[i] = float(np.max(np.abs(dy)))
        time_hist_resultant[i] = float(np.max(total))

    profile_rss = np.sqrt(profile_x**2 + profile_y**2)

    summary_df = pd.DataFrame(
        {
            "Layer_Index": np.arange(1, n_layers + 1, dtype=int),
            "Depth_m": depths,
            "Profile_X_max_m": profile_x,
            "Profile_Y_max_m": profile_y,
            "Profile_RSS_total_m": profile_rss,
            "TimeHist_X_maxabs_m": time_hist_x,
            "TimeHist_Y_maxabs_m": time_hist_y,
            "TimeHist_Resultant_total_m": time_hist_resultant,
        }
    )

    return {
        "layer_names": layer_names,
        "depths": depths,
        "summary_df": summary_df,
    }


def compute_legacy_methods(
    x_xl: pd.ExcelFile,
    y_xl: pd.ExcelFile,
    options: Mapping[str, Any] | None = None,
) -> pd.DataFrame:
    bundle = _compute_legacy_bundle(x_xl, y_xl, options)
    return bundle["summary_df"].copy()


def _build_comparison_df(strain_df: pd.DataFrame, legacy_df: pd.DataFrame) -> pd.DataFrame:
    merged = strain_df[
        [
            "Layer_Index",
            "Depth_m",
            "X_base_rel_max_m",
            "Y_base_rel_max_m",
            "Total_base_rel_max_m",
            "Total_tbdy_total_max_m",
            "Total_input_proxy_rel_max_m",
        ]
    ].merge(
        legacy_df[
            [
                "Layer_Index",
                "Depth_m",
                "Profile_X_max_m",
                "Profile_Y_max_m",
                "Profile_RSS_total_m",
                "TimeHist_Resultant_total_m",
            ]
        ],
        on=["Layer_Index", "Depth_m"],
        how="inner",
    )

    profile_x_bottom = float(merged["Profile_X_max_m"].iloc[-1])
    profile_y_bottom = float(merged["Profile_Y_max_m"].iloc[-1])
    profile_rss_bottom = float(merged["Profile_RSS_total_m"].iloc[-1])

    merged["Profile_X_minus_bottom_m"] = merged["Profile_X_max_m"] - profile_x_bottom
    merged["Profile_Y_minus_bottom_m"] = merged["Profile_Y_max_m"] - profile_y_bottom
    merged["Profile_RSS_minus_bottom_m"] = merged["Profile_RSS_total_m"] - profile_rss_bottom
    merged["Delta_Xbase_vs_ProfileXminusbottom_m"] = (
        merged["X_base_rel_max_m"] - merged["Profile_X_minus_bottom_m"]
    )
    merged["Delta_Ybase_vs_ProfileYminusbottom_m"] = (
        merged["Y_base_rel_max_m"] - merged["Profile_Y_minus_bottom_m"]
    )

    merged["Delta_base_vs_profile_m"] = merged["Total_base_rel_max_m"] - merged["Profile_RSS_total_m"]
    merged["Delta_tbdy_vs_profile_m"] = merged["Total_tbdy_total_max_m"] - merged["Profile_RSS_total_m"]
    merged["Delta_base_vs_timehist_m"] = merged["Total_base_rel_max_m"] - merged["TimeHist_Resultant_total_m"]
    merged["Delta_inputproxy_vs_profile_m"] = (
        merged["Total_input_proxy_rel_max_m"] - merged["Profile_RSS_total_m"]
    )

    with np.errstate(divide="ignore", invalid="ignore"):
        merged["Ratio_base_to_profile"] = np.where(
            merged["Profile_RSS_total_m"] != 0,
            merged["Total_base_rel_max_m"] / merged["Profile_RSS_total_m"],
            np.nan,
        )
        merged["Ratio_tbdy_to_profile"] = np.where(
            merged["Profile_RSS_total_m"] != 0,
            merged["Total_tbdy_total_max_m"] / merged["Profile_RSS_total_m"],
            np.nan,
        )
        merged["Ratio_base_to_timehist"] = np.where(
            merged["TimeHist_Resultant_total_m"] != 0,
            merged["Total_base_rel_max_m"] / merged["TimeHist_Resultant_total_m"],
            np.nan,
        )

    return merged


def _build_depth_profiles_df(comparison_df: pd.DataFrame) -> pd.DataFrame:
    return comparison_df[
        [
            "Layer_Index",
            "Depth_m",
            "Total_base_rel_max_m",
            "Total_tbdy_total_max_m",
            "Total_input_proxy_rel_max_m",
            "Profile_RSS_total_m",
            "TimeHist_Resultant_total_m",
        ]
    ].copy()


def _build_base_corrected_profiles_df(comparison_df: pd.DataFrame) -> pd.DataFrame:
    return comparison_df[
        [
            "Layer_Index",
            "Depth_m",
            "X_base_rel_max_m",
            "Profile_X_minus_bottom_m",
            "Delta_Xbase_vs_ProfileXminusbottom_m",
            "Y_base_rel_max_m",
            "Profile_Y_minus_bottom_m",
            "Delta_Ybase_vs_ProfileYminusbottom_m",
            "Total_base_rel_max_m",
            "Profile_RSS_minus_bottom_m",
        ]
    ].copy()


def _get_common_time_for_layer(
    time_a: np.ndarray,
    time_b: np.ndarray,
) -> np.ndarray:
    start = max(float(time_a[0]), float(time_b[0]))
    end = min(float(time_a[-1]), float(time_b[-1]))
    if end <= start:
        return time_a.copy()

    dta = np.median(np.diff(time_a)) if time_a.size > 1 else 0.01
    dtb = np.median(np.diff(time_b)) if time_b.size > 1 else 0.01
    dt = min(dta, dtb)
    if not np.isfinite(dt) or dt <= 0:
        dt = 0.01

    count = int(math.floor((end - start) / dt)) + 1
    return start + np.arange(count, dtype=float) * dt


def _configure_chart_axes(chart) -> None:
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.tickLblPos = "nextTo"
    chart.y_axis.tickLblPos = "nextTo"
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"


def _add_depth_profile_chart(worksheet, n_rows: int) -> None:
    if n_rows < 2:
        return

    chart = ScatterChart()
    chart.title = "Depth-Dependent Total Displacement Profiles"
    chart.x_axis.title = "Displacement (m)"
    chart.y_axis.title = "Depth (m)"
    chart.legend.position = "r"
    chart.y_axis.scaling.orientation = "maxMin"
    chart.height = 9.5
    chart.width = 15.0
    _configure_chart_axes(chart)

    y_values = Reference(worksheet, min_col=2, min_row=2, max_row=n_rows + 1)
    for col in range(3, worksheet.max_column + 1):
        x_values = Reference(worksheet, min_col=col, min_row=2, max_row=n_rows + 1)
        series = Series(x_values, y_values, title=worksheet.cell(row=1, column=col).value)
        chart.series.append(series)

    worksheet.add_chart(chart, "H2")


def _add_base_corrected_chart(worksheet, n_rows: int) -> None:
    if n_rows < 2:
        return

    chart_x = ScatterChart()
    chart_x.title = "X Profile: Strain Base-Relative vs Deepsoil(Base-Corrected)"
    chart_x.x_axis.title = "Displacement (m)"
    chart_x.y_axis.title = "Depth (m)"
    chart_x.legend.position = "r"
    chart_x.y_axis.scaling.orientation = "maxMin"
    chart_x.height = 8.5
    chart_x.width = 14.0
    _configure_chart_axes(chart_x)

    y_values = Reference(worksheet, min_col=2, min_row=2, max_row=n_rows + 1)
    for col in (3, 4):
        x_values = Reference(worksheet, min_col=col, min_row=2, max_row=n_rows + 1)
        series = Series(x_values, y_values, title=worksheet.cell(row=1, column=col).value)
        chart_x.series.append(series)

    chart_y = ScatterChart()
    chart_y.title = "Y Profile: Strain Base-Relative vs Deepsoil(Base-Corrected)"
    chart_y.x_axis.title = "Displacement (m)"
    chart_y.y_axis.title = "Depth (m)"
    chart_y.legend.position = "r"
    chart_y.y_axis.scaling.orientation = "maxMin"
    chart_y.height = 8.5
    chart_y.width = 14.0
    _configure_chart_axes(chart_y)

    for col in (6, 7):
        x_values = Reference(worksheet, min_col=col, min_row=2, max_row=n_rows + 1)
        series = Series(x_values, y_values, title=worksheet.cell(row=1, column=col).value)
        chart_y.series.append(series)

    worksheet.add_chart(chart_x, "L2")
    worksheet.add_chart(chart_y, "L22")


def _add_all_layers_chart(
    worksheet,
    n_rows: int,
    n_series: int,
    title: str,
) -> None:
    if n_rows < 2 or n_series < 1:
        return

    chart = LineChart()
    chart.title = title
    chart.y_axis.title = "Displacement (m)"
    chart.x_axis.title = "Time (s)"
    chart.height = 10.0
    chart.width = 18.0
    chart.legend.position = "r"
    _configure_chart_axes(chart)

    categories = Reference(worksheet, min_col=1, min_row=2, max_row=n_rows + 1)
    values = Reference(worksheet, min_col=2, max_col=n_series + 1, min_row=1, max_row=n_rows + 1)
    chart.add_data(values, titles_from_data=True)
    chart.set_categories(categories)
    worksheet.add_chart(chart, "B2")


def build_output_workbook(
    strain_df: pd.DataFrame,
    legacy_df: pd.DataFrame,
    comparison_df: pd.DataFrame,
    x_time_df: pd.DataFrame | None = None,
    y_time_df: pd.DataFrame | None = None,
    resultant_time_df: pd.DataFrame | None = None,
    tbdy_total_x_time_df: pd.DataFrame | None = None,
    tbdy_total_y_time_df: pd.DataFrame | None = None,
    tbdy_total_resultant_time_df: pd.DataFrame | None = None,
) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        strain_df.to_excel(writer, sheet_name="Strain_Relative", index=False)
        legacy_df.to_excel(writer, sheet_name="Legacy_Methods", index=False)
        comparison_df.to_excel(writer, sheet_name="Comparison", index=False)

        depth_profiles_df = _build_depth_profiles_df(comparison_df)
        depth_profiles_df.to_excel(writer, sheet_name="Depth_Profiles", index=False)

        base_corrected_df = _build_base_corrected_profiles_df(comparison_df)
        base_corrected_df.to_excel(writer, sheet_name="Profile_BaseCorrected", index=False)

        if x_time_df is not None and not x_time_df.empty:
            x_time_df.to_excel(writer, sheet_name="Direction_X_Time", index=False)
        if y_time_df is not None and not y_time_df.empty:
            y_time_df.to_excel(writer, sheet_name="Direction_Y_Time", index=False)
        if resultant_time_df is not None and not resultant_time_df.empty:
            resultant_time_df.to_excel(writer, sheet_name="Resultant_Time", index=False)
        if tbdy_total_x_time_df is not None and not tbdy_total_x_time_df.empty:
            tbdy_total_x_time_df.to_excel(writer, sheet_name="TBDY_Total_X_Time", index=False)
        if tbdy_total_y_time_df is not None and not tbdy_total_y_time_df.empty:
            tbdy_total_y_time_df.to_excel(writer, sheet_name="TBDY_Total_Y_Time", index=False)
        if tbdy_total_resultant_time_df is not None and not tbdy_total_resultant_time_df.empty:
            tbdy_total_resultant_time_df.to_excel(writer, sheet_name="TBDY_Total_Resultant_Time", index=False)

        workbook = writer.book

        if "Depth_Profiles" in writer.sheets:
            ws_depth = writer.sheets["Depth_Profiles"]
            _add_depth_profile_chart(ws_depth, len(depth_profiles_df))

        if "Profile_BaseCorrected" in writer.sheets:
            ws_bc = writer.sheets["Profile_BaseCorrected"]
            _add_base_corrected_chart(ws_bc, len(base_corrected_df))

        if "Direction_X_Time" in writer.sheets:
            ws_x = writer.sheets["Direction_X_Time"]
            _add_all_layers_chart(
                ws_x,
                ws_x.max_row - 1,
                ws_x.max_column - 1,
                "Direction X: All Layers Signed Displacement-Time",
            )

        if "Direction_Y_Time" in writer.sheets:
            ws_y = writer.sheets["Direction_Y_Time"]
            _add_all_layers_chart(
                ws_y,
                ws_y.max_row - 1,
                ws_y.max_column - 1,
                "Direction Y: All Layers Signed Displacement-Time",
            )

        if "Resultant_Time" in writer.sheets:
            ws_r = writer.sheets["Resultant_Time"]
            _add_all_layers_chart(
                ws_r,
                ws_r.max_row - 1,
                ws_r.max_column - 1,
                "Resultant: All Layers Displacement-Time",
            )

        if "TBDY_Total_X_Time" in writer.sheets:
            ws_tx = writer.sheets["TBDY_Total_X_Time"]
            _add_all_layers_chart(
                ws_tx,
                ws_tx.max_row - 1,
                ws_tx.max_column - 1,
                "TBDY Total X: u(base)+u(rel)",
            )

        if "TBDY_Total_Y_Time" in writer.sheets:
            ws_ty = writer.sheets["TBDY_Total_Y_Time"]
            _add_all_layers_chart(
                ws_ty,
                ws_ty.max_row - 1,
                ws_ty.max_column - 1,
                "TBDY Total Y: u(base)+u(rel)",
            )

        if "TBDY_Total_Resultant_Time" in writer.sheets:
            ws_tr = writer.sheets["TBDY_Total_Resultant_Time"]
            _add_all_layers_chart(
                ws_tr,
                ws_tr.max_row - 1,
                ws_tr.max_column - 1,
                "TBDY Total Resultant: All Layers",
            )

        _ = workbook

    return buffer.getvalue()


def build_single_output_workbook(
    summary_df: pd.DataFrame,
    direction_time_df: pd.DataFrame,
    strain_rel_time_df: pd.DataFrame | None = None,
    tbdy_total_time_df: pd.DataFrame | None = None,
    input_proxy_rel_time_df: pd.DataFrame | None = None,
) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Single_Direction_Summary", index=False)
        direction_time_df.to_excel(writer, sheet_name="Direction_Time", index=False)

        if strain_rel_time_df is not None and not strain_rel_time_df.empty:
            strain_rel_time_df.to_excel(writer, sheet_name="Strain_Relative_Time", index=False)
        if tbdy_total_time_df is not None and not tbdy_total_time_df.empty:
            tbdy_total_time_df.to_excel(writer, sheet_name="TBDY_Total_Time", index=False)
        if input_proxy_rel_time_df is not None and not input_proxy_rel_time_df.empty:
            input_proxy_rel_time_df.to_excel(writer, sheet_name="InputProxy_Relative_Time", index=False)

        if "Direction_Time" in writer.sheets:
            ws_dir = writer.sheets["Direction_Time"]
            _add_all_layers_chart(
                ws_dir,
                ws_dir.max_row - 1,
                ws_dir.max_column - 1,
                "Single Direction: All Layers Displacement-Time",
            )

        if "Strain_Relative_Time" in writer.sheets:
            ws_sr = writer.sheets["Strain_Relative_Time"]
            _add_all_layers_chart(
                ws_sr,
                ws_sr.max_row - 1,
                ws_sr.max_column - 1,
                "Single Direction: Strain Base-Relative Time",
            )

        if "TBDY_Total_Time" in writer.sheets:
            ws_tb = writer.sheets["TBDY_Total_Time"]
            _add_all_layers_chart(
                ws_tb,
                ws_tb.max_row - 1,
                ws_tb.max_column - 1,
                "Single Direction: TBDY Total Time (u_base + u_rel)",
            )

        if "InputProxy_Relative_Time" in writer.sheets:
            ws_ip = writer.sheets["InputProxy_Relative_Time"]
            _add_all_layers_chart(
                ws_ip,
                ws_ip.max_row - 1,
                ws_ip.max_column - 1,
                "Single Direction: Input-Proxy Relative Time",
            )

    return buffer.getvalue()


def _infer_axis_label(file_name: str) -> str:
    upper_name = file_name.upper()
    if "_X_" in upper_name:
        return "X"
    if "_Y_" in upper_name:
        return "Y"
    return "SINGLE"


def _build_pair_key(x_name: str, y_name: str) -> str:
    x_stem = Path(x_name).stem
    y_stem = Path(y_name).stem
    base = x_stem.replace("_X_", "_").replace("_H1", "")
    return f"{base}|{y_stem}"


def process_single_file(
    file_bytes: bytes,
    file_name: str,
    options: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    normalized_options = _normalize_options(options)
    axis_label = _infer_axis_label(file_name)

    with pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl") as xl:
        strain_bundle = _compute_single_strain_bundle(xl, axis_label, normalized_options)
        direction_bundle = _compute_single_direction_disp_bundle(xl, axis_label, normalized_options)
        profile_depths, profile_max = _parse_profile_displacement_max(xl)

        summary_df = strain_bundle["summary_df"].copy()
        n_layers = min(
            len(summary_df),
            int(direction_bundle["disp_matrix"].shape[0]),
            int(profile_max.size),
            int(profile_depths.size),
        )
        summary_df = summary_df.iloc[:n_layers].copy()
        summary_df["Profile_max_m"] = np.abs(profile_max[:n_layers])
        summary_df["TimeHist_maxabs_m"] = np.max(
            np.abs(direction_bundle["disp_matrix"][:n_layers, :]),
            axis=1,
        )

        strain_rel_time_df = _build_layer_time_df(
            strain_bundle["time"],
            strain_bundle["depths"],
            strain_bundle["u_rel_base"],
            "base_rel_m",
        )
        tbdy_total_time_df = _build_layer_time_df(
            strain_bundle["time"],
            strain_bundle["depths"],
            strain_bundle["u_tbdy_total"],
            "tbdy_total_m",
        )
        input_proxy_rel_time_df = _build_layer_time_df(
            strain_bundle["time"],
            strain_bundle["depths"],
            strain_bundle["u_rel_input"],
            "input_proxy_rel_m",
        )

        output_bytes = build_single_output_workbook(
            summary_df=summary_df,
            direction_time_df=direction_bundle["table_df"],
            strain_rel_time_df=strain_rel_time_df,
            tbdy_total_time_df=tbdy_total_time_df,
            input_proxy_rel_time_df=input_proxy_rel_time_df,
        )

    output_file_name = f"output_single_{Path(file_name).stem}.xlsx"
    return {
        "pairKey": f"SINGLE|{Path(file_name).stem}",
        "xFileName": file_name,
        "yFileName": "",
        "outputFileName": output_file_name,
        "outputBytes": output_bytes,
        "metrics": {
            "mode": "single",
            "axis": axis_label,
            "layerCount": int(len(summary_df)),
            "timeSeriesSheets": 4,
            "timeSheets": [
                "Direction_Time",
                "Strain_Relative_Time",
                "TBDY_Total_Time",
                "InputProxy_Relative_Time",
            ],
            "surfaceBaseTotal_m": float(summary_df["Base_rel_max_m"].iloc[0]),
            "surfaceTBDYTotal_m": float(summary_df["TBDY_total_max_m"].iloc[0]),
            "surfaceProfileRSS_m": float(summary_df["Profile_max_m"].iloc[0]),
        },
    }


def process_xy_pair(
    x_bytes: bytes,
    y_bytes: bytes,
    x_name: str,
    y_name: str,
    options: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    normalized_options = _normalize_options(options)

    with pd.ExcelFile(io.BytesIO(x_bytes), engine="openpyxl") as x_xl, pd.ExcelFile(
        io.BytesIO(y_bytes), engine="openpyxl"
    ) as y_xl:
        strain_bundle = _compute_strain_bundle(x_xl, y_xl, normalized_options)
        legacy_bundle = _compute_legacy_bundle(x_xl, y_xl, normalized_options)
        x_direction_bundle = _compute_single_direction_disp_bundle(x_xl, "X", normalized_options)
        y_direction_bundle = _compute_single_direction_disp_bundle(y_xl, "Y", normalized_options)

        strain_df = strain_bundle["summary_df"].copy()
        legacy_df = legacy_bundle["summary_df"].copy()
        comparison_df = _build_comparison_df(strain_df, legacy_df)
        resultant_time_df = _build_resultant_time_df(x_direction_bundle, y_direction_bundle)
        tbdy_total_x_time_df = _build_layer_time_df(
            strain_bundle["time"],
            strain_bundle["depths"],
            strain_bundle["u_tbdy_total_x"],
            "tbdy_total_x_m",
        )
        tbdy_total_y_time_df = _build_layer_time_df(
            strain_bundle["time"],
            strain_bundle["depths"],
            strain_bundle["u_tbdy_total_y"],
            "tbdy_total_y_m",
        )
        tbdy_total_resultant_matrix = np.sqrt(
            strain_bundle["u_tbdy_total_x"] ** 2 + strain_bundle["u_tbdy_total_y"] ** 2
        )
        tbdy_total_resultant_time_df = _build_layer_time_df(
            strain_bundle["time"],
            strain_bundle["depths"],
            tbdy_total_resultant_matrix,
            "tbdy_total_resultant_m",
        )

    output_bytes = build_output_workbook(
        strain_df,
        legacy_df,
        comparison_df,
        x_time_df=x_direction_bundle["table_df"],
        y_time_df=y_direction_bundle["table_df"],
        resultant_time_df=resultant_time_df,
        tbdy_total_x_time_df=tbdy_total_x_time_df,
        tbdy_total_y_time_df=tbdy_total_y_time_df,
        tbdy_total_resultant_time_df=tbdy_total_resultant_time_df,
    )
    output_file_name = f"output_total_{Path(x_name).stem}.xlsx"

    return {
        "pairKey": _build_pair_key(x_name, y_name),
        "xFileName": x_name,
        "yFileName": y_name,
        "outputFileName": output_file_name,
        "outputBytes": output_bytes,
        "metrics": {
            "layerCount": int(len(strain_df)),
            "timeSeriesSheets": 6,
            "timeSheets": [
                "Direction_X_Time",
                "Direction_Y_Time",
                "Resultant_Time",
                "TBDY_Total_X_Time",
                "TBDY_Total_Y_Time",
                "TBDY_Total_Resultant_Time",
            ],
            "surfaceBaseTotal_m": float(strain_df["Total_base_rel_max_m"].iloc[0]),
            "surfaceTBDYTotal_m": float(strain_df["Total_tbdy_total_max_m"].iloc[0]),
            "surfaceProfileRSS_m": float(legacy_df["Profile_RSS_total_m"].iloc[0]),
        },
    }


def _is_candidate_file(name: str, include_manip: bool) -> bool:
    lower_name = name.lower()
    if not lower_name.endswith(".xlsx"):
        return False
    if any(lower_name.startswith(prefix) for prefix in EXCLUDE_PREFIXES):
        return False
    if not include_manip and any(lower_name.endswith(suffix) for suffix in EXCLUDE_SUFFIXES):
        return False
    return True


def _derive_y_name(x_name: str) -> str:
    replaced = x_name.replace("_X_", "_Y_", 1)
    replaced = re.sub(r"_H1(?=\.xlsx$)", "_H2", replaced, flags=re.IGNORECASE)
    return replaced


def find_xy_pairs(file_names: Sequence[str], include_manip: bool = False) -> Tuple[List[Tuple[str, str]], List[str]]:
    candidates = {name for name in file_names if _is_candidate_file(name, include_manip)}

    x_files = sorted(
        [name for name in candidates if "_X_" in name and re.search(r"_H1(?=\.xlsx$)", name, flags=re.IGNORECASE)]
    )

    pairs: List[Tuple[str, str]] = []
    missing: List[str] = []

    for x_name in x_files:
        y_name = _derive_y_name(x_name)
        if y_name in candidates:
            pairs.append((x_name, y_name))
        else:
            missing.append(x_name)

    return pairs, missing


def process_batch_files(file_map: Mapping[str, bytes], options: Mapping[str, Any] | None = None) -> Dict[str, Any]:
    normalized_options = _normalize_options(options)
    include_manip = bool(normalized_options.get("includeManip", False))
    fail_fast = bool(normalized_options.get("failFast", False))
    hp_enabled, hp_cutoff, hp_transition = _highpass_config(normalized_options)

    logs: List[Dict[str, str]] = []
    errors: List[Dict[str, str]] = []
    results: List[Dict[str, Any]] = []

    file_names = sorted(file_map.keys())
    candidates = sorted([name for name in file_names if _is_candidate_file(name, include_manip)])
    pairs, missing = find_xy_pairs(file_names, include_manip=include_manip)

    used_in_pairs = set()
    for x_name, y_name in pairs:
        used_in_pairs.add(x_name)
        used_in_pairs.add(y_name)

    singles = sorted([name for name in candidates if name not in used_in_pairs])

    _log(logs, "info", f"Candidate files: {len(file_names)}")
    _log(logs, "info", f"Detected X/Y pairs: {len(pairs)}")
    _log(logs, "info", f"Detected single files: {len(singles)}")
    _log(
        logs,
        "info",
        f"High-pass: {'on' if hp_enabled else 'off'} | cutoff={hp_cutoff:.4f} Hz | transition={hp_transition:.4f} Hz",
    )

    for missing_x in missing:
        if missing_x in singles:
            _log(logs, "warning", f"No Y match for X file; processing single: {missing_x}")
        else:
            _log(logs, "warning", f"No Y match for X file: {missing_x}")

    pair_processed = 0
    pair_failed = 0
    single_processed = 0
    single_failed = 0

    for x_name, y_name in pairs:
        try:
            result = process_xy_pair(
                file_map[x_name],
                file_map[y_name],
                x_name,
                y_name,
                normalized_options,
            )
            results.append(result)
            pair_processed += 1
            _log(logs, "info", f"Processed pair: {x_name} + {y_name}")
        except Exception as exc:  # noqa: BLE001
            pair_failed += 1
            errors.append({"pairKey": f"{x_name}|{y_name}", "reason": str(exc)})
            _log(logs, "error", f"Failed pair {x_name} + {y_name}: {exc}")
            if fail_fast:
                break

    if not fail_fast or not errors:
        for name in singles:
            try:
                result = process_single_file(
                    file_map[name],
                    name,
                    normalized_options,
                )
                results.append(result)
                single_processed += 1
                _log(logs, "info", f"Processed single: {name}")
            except Exception as exc:  # noqa: BLE001
                single_failed += 1
                errors.append({"pairKey": f"SINGLE|{name}", "reason": str(exc)})
                _log(logs, "error", f"Failed single {name}: {exc}")
                if fail_fast:
                    break

    processed_total = pair_processed + single_processed
    failed_total = pair_failed + single_failed

    return {
        "results": results,
        "logs": logs,
        "errors": errors,
        "metrics": {
            "pairsDetected": len(pairs),
            "pairsProcessed": pair_processed,
            "pairsFailed": pair_failed,
            "pairsMissing": len(missing),
            "singlesDetected": len(singles),
            "singlesProcessed": single_processed,
            "singlesFailed": single_failed,
            "processedTotal": processed_total,
            "failedTotal": failed_total,
        },
    }


def process_batch_directory(
    input_dir: str | os.PathLike[str],
    output_dir: str | os.PathLike[str],
    options: Mapping[str, Any] | None = None,
) -> Dict[str, Any]:
    in_path = Path(input_dir)
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    file_map: Dict[str, bytes] = {}
    for item in in_path.iterdir():
        if not item.is_file() or item.suffix.lower() != ".xlsx":
            continue
        if item.name.startswith("~$"):
            continue
        try:
            file_map[item.name] = item.read_bytes()
        except PermissionError:
            continue

    summary = process_batch_files(file_map, options)

    for result in summary["results"]:
        target = out_path / result["outputFileName"]
        target.write_bytes(result["outputBytes"])
        result["writtenPath"] = str(target)

    return summary


__all__ = [
    "parse_profile_thickness",
    "compute_strain_relative",
    "compute_legacy_methods",
    "build_output_workbook",
    "build_single_output_workbook",
    "process_xy_pair",
    "process_single_file",
    "find_xy_pairs",
    "process_batch_files",
    "process_batch_directory",
]
