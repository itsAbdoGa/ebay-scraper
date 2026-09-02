import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlparse

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from lib.paths import LISTINGS_JSON, SELLERBOARD_HTML, ensure_data_dirs

SKIP_NAMES = {
    "AMAZON",
    "APPLY TO SELL",
    "Click to copy ASIN",
    "PROJ PROFIT",
    "PROJ ROI",
    "REPLEN",
    "ZOOM",
    "Sellerboard Snapshots",
}
ASIN_TEXT_RE = re.compile(r"^B[A-Z0-9]{9}$")
PRICE_RE = re.compile(r"^\$[\d,.]+$")
PERCENT_RE = re.compile(r"^[\d.]+%$")
MIN_TITLE_LENGTH = 20


@dataclass
class AmazonListing:
    name: str
    url: str
    asin: str


class BodyLinkParser(HTMLParser):
    def __init__(self, parse_entire_document: bool = False) -> None:
        super().__init__(convert_charrefs=True)
        self._body_depth = 1 if parse_entire_document else 0
        self._recent_text: list[str] = []
        self.listings: list[AmazonListing] = []

    def handle_data(self, data: str) -> None:
        if self._body_depth == 0:
            return

        text = data.strip()
        if text:
            self._recent_text.append(text)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() == "body":
            self._body_depth += 1
            return

        if self._body_depth == 0 or tag.lower() != "a":
            return

        for name, value in attrs:
            if name.lower() != "href" or not value:
                continue

            url = unquote(unescape(value)).strip()
            asin = extract_asin(url)
            if asin is None:
                return

            listing_name = self._listing_name_from_recent_text()
            self.listings.append(AmazonListing(name=listing_name, url=url, asin=asin))
            self._recent_text.clear()
            return

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "body" and self._body_depth > 0:
            self._body_depth -= 1

    def _listing_name_from_recent_text(self) -> str:
        fallback = ""

        for text in reversed(self._recent_text):
            if not text or text in SKIP_NAMES:
                continue
            if ASIN_TEXT_RE.fullmatch(text):
                continue
            if text.startswith("eBay Order"):
                continue
            if PRICE_RE.match(text) or PERCENT_RE.match(text):
                continue
            if len(text) >= MIN_TITLE_LENGTH:
                return text
            if not fallback:
                fallback = text

        return fallback


def extract_asin(link: str) -> str | None:
    decoded_link = unquote(unescape(link)).strip()
    parsed = urlparse(decoded_link)
    host = parsed.netloc.lower()

    if "@" in host:
        host = host.rsplit("@", 1)[1]
    if ":" in host:
        host = host.split(":", 1)[0]
    if host.startswith("www."):
        host = host[4:]

    if host != "amazon.com":
        return None

    parts = parsed.path.strip("/").split("/")
    if len(parts) < 2 or parts[0].lower() != "dp":
        return None

    asin = parts[1]
    if len(asin) == 10 and asin.isalnum():
        return asin.upper()
    return None


def extract_amazon_listings_from_html(html: str) -> list[AmazonListing]:
    parser = BodyLinkParser(parse_entire_document="<body" not in html.lower())
    parser.feed(html)
    parser.close()
    return parser.listings


def find_duplicate_asins(listings: list[AmazonListing]) -> dict[str, list[AmazonListing]]:
    by_asin: dict[str, list[AmazonListing]] = {}
    for listing in listings:
        by_asin.setdefault(listing.asin, []).append(listing)
    return {asin: items for asin, items in by_asin.items() if len(items) > 1}


def deduplicate_listings(listings: list[AmazonListing]) -> list[AmazonListing]:
    seen: set[str] = set()
    unique: list[AmazonListing] = []

    for listing in listings:
        if listing.asin in seen:
            continue
        seen.add(listing.asin)
        unique.append(listing)

    return unique


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract Amazon /dp/ links from a Sellerboard HTML export."
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=SELLERBOARD_HTML,
        help=f"Sellerboard HTML file (default: {SELLERBOARD_HTML})",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=LISTINGS_JSON,
        help=f"Output JSON file (default: {LISTINGS_JSON})",
    )
    return parser.parse_args()


def main(input_path: Path | None = None, output_path: Path | None = None) -> int:
    html_file = input_path or SELLERBOARD_HTML
    output_file = output_path or LISTINGS_JSON

    ensure_data_dirs()

    if not html_file.exists():
        print(f"Error: input file not found: {html_file}")
        print(f"Place your Sellerboard HTML export at: {SELLERBOARD_HTML}")
        return 1

    html = html_file.read_text(encoding="utf-8")
    listings = extract_amazon_listings_from_html(html)
    duplicates = find_duplicate_asins(listings)
    unique_listings = deduplicate_listings(listings)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(
        json.dumps([asdict(listing) for listing in unique_listings], indent=2),
        encoding="utf-8",
    )

    print(
        f"Found {len(listings)} Amazon /dp/ links "
        f"({len(duplicates)} duplicate ASINs, {len(unique_listings)} unique)."
    )
    print(f"Wrote {len(unique_listings)} unique listings to {output_file}")
    print()

    for index, listing in enumerate(unique_listings, start=1):
        print(f"{index:2}. {listing.name}")
        print(f"    {listing.url}")

    if duplicates:
        print()
        print("Skipped duplicate ASINs:")
        for asin, items in duplicates.items():
            print(f"  {asin} appeared {len(items)} times, kept: {items[0].name}")

    return 0


if __name__ == "__main__":
    args = parse_args()
    raise SystemExit(main(input_path=args.input, output_path=args.output))
