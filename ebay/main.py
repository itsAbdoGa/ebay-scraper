"""
eBay scraper — edit SETTINGS below, then run this file.

You do not need the command line. Change a number, press Run.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from scripts.scrape_listings import ScrapeSettings, main as scrape_main


# ---------------------------------------------------------------------------
# EDIT HERE
# ---------------------------------------------------------------------------

SETTINGS = ScrapeSettings(
    # Excel rows (row 1 is the header). None = don't limit.
    start_row=None,
    end_row=None,
    limit=None,

    # Only these brands, or skip these brands. Empty = every brand.
    include_brands=(),          # e.g. ("LEGO", "Hasbro")
    exclude_brands=(),

    # A listing wins when it is Brand New, US, recent, and in this ROI range.
    min_roi_percent=80,
    max_roi_percent=300,
    max_listing_age_days=2,
    min_seller_reviews=50,
    max_sales_rank=500_000,
    min_buybox_price=40.00,

    # Search EAN, and also UPC when it is different.
    upc_as_well=True,
)


def main(settings: ScrapeSettings | None = None) -> int:
    return scrape_main(settings or SETTINGS)


if __name__ == "__main__":
    raise SystemExit(main())
