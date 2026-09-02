import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from playwright.sync_api import sync_playwright

from lib.amazon_scraper import REQUEST_DELAY_SECONDS, create_browser_context, scrape_product, warm_up_session
from lib.paths import LISTINGS_JSON, PRODUCT_DETAILS_JSON, ensure_data_dirs


def load_listings(path: Path) -> list[dict[str, str]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError(f"{path} must contain a JSON array")

    listings: list[dict[str, str]] = []
    for item in data:
        if not isinstance(item, dict) or "url" not in item:
            continue
        listings.append(item)

    return listings


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape brand and category from Amazon product pages."
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=LISTINGS_JSON,
        help=f"Listings JSON file (default: {LISTINGS_JSON})",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=PRODUCT_DETAILS_JSON,
        help=f"Output JSON file (default: {PRODUCT_DETAILS_JSON})",
    )
    parser.add_argument(
        "-n",
        "--limit",
        type=int,
        default=None,
        help="Only scrape the first N listings (useful for testing)",
    )
    return parser.parse_args()


def main(
    input_path: Path | None = None,
    output_path: Path | None = None,
    limit: int | None = None,
) -> int:
    input_file = input_path or LISTINGS_JSON
    output_file = output_path or PRODUCT_DETAILS_JSON

    ensure_data_dirs()

    if not input_file.exists():
        print(f"Error: listings file not found: {input_file}")
        print("Run 'python run.py extract' first.")
        return 1

    listings = load_listings(input_file)
    if limit is not None:
        listings = listings[:limit]

    results: list[dict[str, str]] = []

    print(f"Loaded {len(listings)} product URLs from {input_file}")
    print()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = create_browser_context(browser)
        page = context.new_page()
        warm_up_session(page)

        try:
            for index, listing in enumerate(listings, start=1):
                url = listing["url"]
                name = listing.get("name", "")

                print(f"[{index}/{len(listings)}] {name or url}")
                result = scrape_product(
                    page,
                    url,
                    title=name,
                )
                results.append(result)

                if index < len(listings):
                    time.sleep(REQUEST_DELAY_SECONDS)
        finally:
            context.close()
            browser.close()

    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(json.dumps(results, indent=2), encoding="utf-8")
    print()
    print(f"Wrote {len(results)} product details to {output_file}")
    return 0


if __name__ == "__main__":
    args = parse_args()
    raise SystemExit(main(input_path=args.input, output_path=args.output, limit=args.limit))
