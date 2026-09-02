"""Optional CLI. To scrape, edit SETTINGS in main.py and run that file."""

import argparse
import sys
from dataclasses import replace
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from lib.paths import (
    COMBINED_XLSX,
    EBAY_COOKIES_FILE,
    INPUT_DIR,
    OUTPUT_DIR,
    WINNING_LISTINGS_XLSX,
    ensure_data_dirs,
)
from main import SETTINGS, main as scrape_from_settings
from scripts.enrich_amazon_playwright import DEFAULT_PROFILE_DIR, enrich_workbook
from scripts.enrich_amazon_requests import DEFAULT_OUTPUT as AMAZON_ENRICHED_XLSX


def print_status() -> None:
    ensure_data_dirs()
    print("eBay Listings Scraper")
    print("=" * 40)
    print()
    print("To scrape: edit SETTINGS in main.py, then run that file.")
    print()
    print(f"  input:  {INPUT_DIR}")
    print(f"  output: {OUTPUT_DIR}")
    print()
    for label, path in (
        ("Workbook", COMBINED_XLSX),
        ("Cookies", EBAY_COOKIES_FILE),
        ("Winners", WINNING_LISTINGS_XLSX),
    ):
        status = "ok" if path.exists() else "missing"
        print(f"  [{status}] {label}: {path}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="eBay listings scraper. Prefer editing ebay/main.py.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  python run.py status\n"
            "  python run.py scrape\n"
            "  python run.py scrape --start 2 --end 20\n"
            "  python run.py enrich-amazon"
        ),
    )
    subparsers = parser.add_subparsers(dest="command")
    subparsers.add_parser("status", help="Show file locations")

    scrape_parser = subparsers.add_parser(
        "scrape",
        help="Scrape using SETTINGS from main.py",
    )
    scrape_parser.add_argument("--start", type=int, default=None)
    scrape_parser.add_argument("--end", type=int, default=None)
    scrape_parser.add_argument("-n", "--limit", type=int, default=None)
    scrape_parser.add_argument("--html", type=Path, default=None)

    enrich_parser = subparsers.add_parser(
        "enrich-amazon",
        help="Fetch Amazon details with Playwright into a workbook copy",
    )
    enrich_parser.add_argument("-i", "--input", type=Path, default=COMBINED_XLSX)
    enrich_parser.add_argument("-o", "--output", type=Path, default=AMAZON_ENRICHED_XLSX)
    enrich_parser.add_argument("--timeout-seconds", type=float, default=10.0)
    enrich_parser.add_argument("--retries", type=int, default=2)
    enrich_parser.add_argument("--delay-seconds", type=float, default=0.0)
    enrich_parser.add_argument("--min-buybox-price", type=float, default=40.0)
    enrich_parser.add_argument("--save-every", type=int, default=10)
    enrich_parser.add_argument("--max-consecutive-blocks", type=int, default=3)
    enrich_parser.add_argument("--skip-recent-days", type=int, default=3)
    enrich_parser.add_argument("--fresh-copy", action="store_true")
    enrich_parser.add_argument(
        "--browser-profile",
        type=Path,
        default=DEFAULT_PROFILE_DIR,
    )
    enrich_parser.add_argument("--headless", action="store_true")
    enrich_parser.add_argument("--start-row", type=int, default=None)
    enrich_parser.add_argument("--end-row", type=int, default=None)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.command in (None, "status"):
        print_status()
        return 0

    if args.command == "scrape":
        updates = {}
        if args.start is not None:
            updates["start_row"] = args.start
        if args.end is not None:
            updates["end_row"] = args.end
        if args.limit is not None:
            updates["limit"] = args.limit
        if args.html is not None:
            updates["html_path"] = args.html
        settings = replace(SETTINGS, **updates) if updates else SETTINGS
        return scrape_from_settings(settings)

    if args.command == "enrich-amazon":
        return enrich_workbook(
            input_path=args.input,
            output_path=args.output,
            timeout_seconds=args.timeout_seconds,
            retries=args.retries,
            delay_seconds=args.delay_seconds,
            min_buybox_price=args.min_buybox_price,
            save_every=args.save_every,
            max_consecutive_blocks=args.max_consecutive_blocks,
            fresh_copy=args.fresh_copy,
            user_data_dir=args.browser_profile,
            headless=args.headless,
            start=args.start_row,
            end=args.end_row,
            skip_recent_days=args.skip_recent_days,
        )

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
