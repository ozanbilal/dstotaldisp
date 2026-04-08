from pathlib import Path

from disp_core import process_batch_files


ROOT = Path(__file__).resolve().parents[1]


def _load_inputs(folder: str) -> dict[str, bytes]:
    base = ROOT / folder
    return {item.name: item.read_bytes() for item in base.iterdir() if item.is_file()}


def test_batch_progress_callback_reports_percentages():
    events: list[tuple[str, str, float, bool]] = []

    def progress_callback(message: str, phase: str, progress: float, indeterminate: bool) -> None:
        events.append((message, phase, float(progress), bool(indeterminate)))

    process_batch_files(
        _load_inputs("_tmp_m23_case_pair_in"),
        {
            "method2Enabled": True,
            "method3Enabled": True,
            "_progress_callback": progress_callback,
        },
    )

    assert events
    numeric_progress = [progress for _, _, progress, indeterminate in events if not indeterminate]
    assert numeric_progress
    assert numeric_progress == sorted(numeric_progress)
    assert numeric_progress[0] >= 55.0
    assert numeric_progress[-1] >= 90.0
    assert any("Method-3 aggregate workbook" in message for message, _, _, _ in events)
