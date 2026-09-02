import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from lib.paths import (
    ALL_ASIN_DETAILS_DEDUPED_JSON,
    ALL_ASIN_DETAILS_JSON,
    ALL_ASIN_DETAILS_XLSX,
    ALL_ASINS_CSV,
    ALREADY_REGISTERED_CSV,
    INPUT_DIR,
    JSON_DIR,
    LISTINGS_JSON,
    OUTPUT_DIR,
    PRODUCT_DETAILS_JSON,
    PRODUCT_DETAILS_XLSX,
    SELLERBOARD_HTML,
    ensure_data_dirs,
)
from scripts.export_xlsx import main as export_main
from scripts.extract_links import main as extract_main
from scripts.scrape_asins import main as scrape_asins_main
from scripts.scrape_products import main as scrape_main

PIPELINE_STEPS = ("extract", "scrape", "export")


def count_json_records(path: Path) -> int | None:
    if not path.exists():
        return None
    data = json.loads(path.read_text(encoding="utf-8"))
    return len(data) if isinstance(data, list) else None


def print_status() -> None:
    ensure_data_dirs()

    print("Amazon Gating Pipeline")
    print("=" * 40)
    print()
    print("Data folders:")
    print(f"  input:  {INPUT_DIR}")
    print(f"  json:   {JSON_DIR}")
    print(f"  output: {OUTPUT_DIR}")
    print()
    print("Files:")

    files = [
        ("Sellerboard HTML", SELLERBOARD_HTML, None),
        ("All ASINs CSV", ALL_ASINS_CSV, None),
        ("Already registered CSV", ALREADY_REGISTERED_CSV, None),
        ("Listings JSON", LISTINGS_JSON, count_json_records),
        ("Product details JSON", PRODUCT_DETAILS_JSON, count_json_records),
        ("All ASIN details JSON", ALL_ASIN_DETAILS_JSON, count_json_records),
        ("All ASIN details deduped JSON", ALL_ASIN_DETAILS_DEDUPED_JSON, count_json_records),
        ("Product details XLSX", PRODUCT_DETAILS_XLSX, None),
        ("All ASIN details XLSX", ALL_ASIN_DETAILS_XLSX, None),
    ]

    for label, path, counter in files:
        if path.exists():
            extra = ""
            if counter is not None:
                count = counter(path)
                if count is not None:
                    extra = f" ({count} records)"
            elif path.suffix == ".xlsx":
                extra = f" ({path.stat().st_size:,} bytes)"
            print(f"  [ok] {label}: {path}{extra}")
        else:
            print(f"  [ ] {label}: {path} (missing)")

    print()
    print("Quick start:")
    print("  1. Save Sellerboard export to data/input/sellerboard.html")
    print("  2. python run.py extract")
    print("  3. python run.py scrape")
    print("  4. python run.py export")
    print("  Or run everything: python run.py all")
    print()
    print("ASIN CSV workflow:")
    print("  1. Save ASIN export to data/input/all-asins.csv")
    print("  2. Save registered products to data/input/already registered.csv")
    print("  3. python run.py scrape-asins")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Amazon gating pipeline: extract links, scrape details, export Excel.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  python run.py status\n"
            "  python run.py extract\n"
            "  python run.py scrape --limit 2\n"
            "  python run.py export\n"
            "  python run.py all\n"
            "  python run.py scrape-asins --limit 2"
        ),
    )
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("status", help="Show pipeline file locations and readiness")

    subparsers.add_parser("extract", help="Step 1: extract Amazon links from Sellerboard HTML")

    scrape_parser = subparsers.add_parser("scrape", help="Step 2: scrape brand/category from Amazon")
    scrape_parser.add_argument(
        "-n",
        "--limit",
        type=int,
        default=None,
        help="Only scrape the first N listings",
    )

    export_parser = subparsers.add_parser("export", help="Step 3: export product details to Excel")
    export_parser.add_argument(
        "--no-deduplicate",
        action="store_true",
        help="Keep all rows instead of one row per brand/category pair",
    )

    all_parser = subparsers.add_parser("all", help="Run extract, scrape, and export in sequence")
    all_parser.add_argument(
        "-n",
        "--limit",
        type=int,
        default=None,
        help="Only scrape the first N listings",
    )
    all_parser.add_argument(
        "--no-deduplicate",
        action="store_true",
        help="Keep all rows in the Excel export",
    )

    scrape_asins_parser = subparsers.add_parser(
        "scrape-asins",
        help="Scrape brand/category for every ASIN in all-asins.csv",
    )
    scrape_asins_parser.add_argument(
        "-n",
        "--limit",
        type=int,
        default=None,
        help="Only scrape the first N ASINs",
    )

    return parser


def run_all(limit: int | None, deduplicate: bool) -> int:
    steps = [
        (
            "extract",
            lambda: extract_main(input_path=SELLERBOARD_HTML, output_path=LISTINGS_JSON),
        ),
        (
            "scrape",
            lambda: scrape_main(
                input_path=LISTINGS_JSON,
                output_path=PRODUCT_DETAILS_JSON,
                limit=limit,
            ),
        ),
        (
            "export",
            lambda: export_main(
                input_path=PRODUCT_DETAILS_JSON,
                output_path=PRODUCT_DETAILS_XLSX,
                deduplicate=deduplicate,
            ),
        ),
    ]

    for step_name, step_fn in steps:
        print()
        print(f"=== {step_name.upper()} ===")
        print()
        code = step_fn()
        if code != 0:
            print(f"\nPipeline stopped at step: {step_name}")
            return code

    print()
    print("Pipeline complete.")
    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.command is None:
        print_status()
        return 0

    if args.command == "status":
        print_status()
        return 0

    if args.command == "extract":
        return extract_main(input_path=SELLERBOARD_HTML, output_path=LISTINGS_JSON)

    if args.command == "scrape":
        return scrape_main(
            input_path=LISTINGS_JSON,
            output_path=PRODUCT_DETAILS_JSON,
            limit=args.limit,
        )

    if args.command == "export":
        return export_main(
            input_path=PRODUCT_DETAILS_JSON,
            output_path=PRODUCT_DETAILS_XLSX,
            deduplicate=not args.no_deduplicate,
        )

    if args.command == "all":
        return run_all(limit=args.limit, deduplicate=not args.no_deduplicate)

    if args.command == "scrape-asins":
        return scrape_asins_main(limit=args.limit)

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
