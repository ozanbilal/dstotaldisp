import base64
import io
import os
import zipfile
from pathlib import Path
from typing import Any, Dict, List

from disp_core import process_batch_files


def _normalize_options(options: Any) -> Dict[str, Any]:
    if options is None:
        return {}
    if hasattr(options, "to_py"):
        options = options.to_py()
    if isinstance(options, dict):
        return dict(options)
    return {}


def run_batch_from_fs(input_dir: str, options: Any = None) -> Dict[str, Any]:
    path = Path(input_dir)
    if not path.exists() or not path.is_dir():
        raise ValueError(f"Input directory does not exist in FS: {input_dir}")

    file_map: Dict[str, bytes] = {}
    for item in path.iterdir():
        if item.is_file() and item.suffix.lower() == ".xlsx":
            file_map[item.name] = item.read_bytes()

    summary = process_batch_files(file_map, _normalize_options(options))

    web_results: List[Dict[str, Any]] = []
    for result in summary["results"]:
        encoded = base64.b64encode(result["outputBytes"]).decode("ascii")
        web_result = {
            "pairKey": result["pairKey"],
            "xFileName": result["xFileName"],
            "yFileName": result["yFileName"],
            "outputFileName": result["outputFileName"],
            "outputBytesB64": encoded,
            "metrics": result["metrics"],
        }
        web_results.append(web_result)

    return {
        "results": web_results,
        "logs": summary["logs"],
        "errors": summary["errors"],
        "metrics": summary["metrics"],
    }


def build_zip_from_results(results: Any) -> str:
    if hasattr(results, "to_py"):
        results = results.to_py()

    if not isinstance(results, list):
        raise ValueError("results must be a list.")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for item in results:
            if not isinstance(item, dict):
                continue
            name = item.get("outputFileName")
            payload_b64 = item.get("outputBytesB64")
            if not name or not payload_b64:
                continue
            content = base64.b64decode(payload_b64)
            zf.writestr(name, content)

    return base64.b64encode(buffer.getvalue()).decode("ascii")
