import argparse
import sys
from pathlib import Path

from disp_core import process_batch_directory
from report_alignment import generate_alignment_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Process DEEPSOIL result Excel files and build displacement outputs. "
            "Supports both X/Y paired mode and single-file mode."
        )
    )
    parser.add_argument(
        "--input-dir",
        default=".",
        help="Directory containing DEEPSOIL result Excel files (default: current directory).",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Directory to write output files (default: same as input-dir).",
    )
    parser.add_argument(
        "--include-manip",
        action="store_true",
        help="Include *-manip.xlsx files while searching pairs.",
    )
    parser.add_argument(
        "--fail-fast",
        action="store_true",
        help="Stop batch processing on first pair failure.",
    )
    parser.add_argument(
        "--with-report",
        action="store_true",
        help="Generate alignment markdown+plot report files.",
    )
    parser.add_argument(
        "--no-method23",
        action="store_true",
        help="Disable Method-2/Method-3 extra export files.",
    )
    parser.add_argument(
        "--no-method2",
        action="store_true",
        help="Disable Method-2 per-file output files.",
    )
    parser.add_argument(
        "--no-method3",
        action="store_true",
        help="Disable Method-3 aggregate output file.",
    )
    parser.add_argument(
        "--baseline-on",
        action="store_true",
        help="Enable baseline correction (default: off).",
    )
    parser.add_argument(
        "--filter-on",
        action="store_true",
        help="Enable filtering (default: off).",
    )
    parser.add_argument(
        "--base-reference",
        choices=("input", "deepest_layer"),
        default="input",
        help="Base displacement reference for TBDY total: input (default) or deepest_layer.",
    )
    parser.add_argument(
        "--integration-compare",
        action="store_true",
        help="Enable alternative FFT-regularized integration and output comparison sheets/columns.",
    )
    parser.add_argument(
        "--hide-resultant-profiles",
        action="store_true",
        help="Hide resultant/total series from Depth_Profiles sheet and chart.",
    )
    parser.add_argument(
        "--alt-integration-method",
        choices=("fft_regularized",),
        default="fft_regularized",
        help="Alternative integration method when --integration-compare is enabled.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve() if args.output_dir else input_dir

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input directory not found: {input_dir}", file=sys.stderr)
        return 2

    method23_enabled = not bool(args.no_method23)
    method2_enabled = method23_enabled and (not bool(args.no_method2))
    method3_enabled = method23_enabled and (not bool(args.no_method3))

    options = {
        "includeManip": bool(args.include_manip),
        "failFast": bool(args.fail_fast),
        "method2Enabled": method2_enabled,
        "method3Enabled": method3_enabled,
        "baselineOn": bool(args.baseline_on),
        "filterOn": bool(args.filter_on),
        "baseReference": str(args.base_reference),
        "integrationCompareEnabled": bool(args.integration_compare),
        "includeResultantProfiles": not bool(args.hide_resultant_profiles),
        "altIntegrationMethod": str(args.alt_integration_method),
    }

    summary = process_batch_directory(input_dir, output_dir, options)

    for log in summary["logs"]:
        print(f"[{log['level'].upper()}] {log['message']}")

    for result in summary["results"]:
        print(f"[OUTPUT] {result['writtenPath']}")
        mode = str(result.get("metrics", {}).get("mode", "pair")).lower()
        if args.with_report and mode == "pair":
            try:
                artifacts = generate_alignment_report(result["writtenPath"])
                print(f"[REPORT] {artifacts.markdown_path}")
                print(f"[PLOT] {artifacts.profile_plot_path}")
                print(f"[PLOT] {artifacts.delta_plot_path}")
            except Exception as report_exc:  # noqa: BLE001
                print(f"[WARN] Report generation failed for {result['writtenPath']}: {report_exc}")
        elif args.with_report:
            print(f"[INFO] Skipping alignment report for mode={mode}: {result['writtenPath']}")

    if summary["errors"]:
        for err in summary["errors"]:
            print(f"[ERROR] {err['pairKey']}: {err['reason']}", file=sys.stderr)
        if args.fail_fast:
            return 1

    metrics = summary["metrics"]
    print(
        "[SUMMARY] "
        f"pairs_detected={metrics['pairsDetected']} pairs_processed={metrics['pairsProcessed']} "
        f"pairs_failed={metrics['pairsFailed']} pairs_missing={metrics['pairsMissing']} "
        f"singles_detected={metrics.get('singlesDetected', 0)} "
        f"singles_processed={metrics.get('singlesProcessed', 0)} "
        f"singles_failed={metrics.get('singlesFailed', 0)} "
        f"method2_enabled={metrics.get('method2Enabled', False)} "
        f"method3_enabled={metrics.get('method3Enabled', False)} "
        f"include_resultant_profiles={metrics.get('includeResultantProfiles', True)} "
        f"base_reference={metrics.get('baseReference', 'input')} "
        f"integration_compare={metrics.get('integrationCompareEnabled', False)} "
        f"alt_method={metrics.get('altIntegrationMethod', '-')} "
        f"method2_detected={metrics.get('method2Detected', 0)} "
        f"method2_processed={metrics.get('method2Processed', 0)} "
        f"method2_failed={metrics.get('method2Failed', 0)} "
        f"method3_produced={metrics.get('method3Produced', 0)} "
        f"processed_total={metrics.get('processedTotal', metrics['pairsProcessed'])} "
        f"failed_total={metrics.get('failedTotal', metrics['pairsFailed'])}"
    )
    return 0 if not summary["errors"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
