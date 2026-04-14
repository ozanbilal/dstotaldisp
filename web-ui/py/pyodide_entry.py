import base64
from pathlib import Path
from typing import Any, Dict, List, Mapping

from disp_core import process_batch_files


def _normalize_options(options: Any) -> Dict[str, Any]:
    if options is None:
        return {}
    if hasattr(options, "to_py"):
        options = options.to_py()
    if isinstance(options, dict):
        return dict(options)
    return {}


class _LazyFsFileMap(Mapping[str, bytes]):
    def __init__(self, directory: Path) -> None:
        self._items: Dict[str, Path] = {}
        for item in directory.iterdir():
            if not item.is_file() or item.suffix.lower() not in {".xlsx", ".db", ".db3"}:
                continue
            if item.name.startswith("~$"):
                continue
            self._items[item.name] = item

    def __getitem__(self, key: str) -> bytes:
        return self._items[key].read_bytes()

    def __iter__(self):
        return iter(self._items)

    def __len__(self) -> int:
        return len(self._items)


def run_batch_from_fs(input_dir: str, options: Any = None, progress_callback: Any = None) -> Dict[str, Any]:
    path = Path(input_dir)
    if not path.exists() or not path.is_dir():
        raise ValueError(f"Input directory does not exist in FS: {input_dir}")

    file_map: Mapping[str, bytes] = _LazyFsFileMap(path)
    normalized_options = _normalize_options(options)
    normalized_options["_returnWebResults"] = True
    if progress_callback is not None:
        normalized_options["_progress_callback"] = progress_callback

    summary = process_batch_files(file_map, normalized_options)

    web_results: List[Dict[str, Any]] = []
    for result in summary["results"]:
        if isinstance(result, dict) and "outputBytesB64" in result:
            web_results.append(result)
            continue
        encoded = base64.b64encode(result["outputBytes"]).decode("ascii")
        web_result = {
            "pairKey": result["pairKey"],
            "xFileName": result["xFileName"],
            "yFileName": result["yFileName"],
            "outputFileName": result["outputFileName"],
            "outputBytesB64": encoded,
            "previewCharts": result.get("previewCharts", []),
            "metrics": result["metrics"],
        }
        web_results.append(web_result)

    return {
        "results": web_results,
        "sourceCatalog": summary.get("sourceCatalog", []),
        "summaryCatalog": summary.get("summaryCatalog", []),
        "logs": summary["logs"],
        "errors": summary["errors"],
        "metrics": summary["metrics"],
    }
