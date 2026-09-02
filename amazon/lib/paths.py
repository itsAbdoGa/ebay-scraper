from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

DATA_DIR = ROOT / "data"
INPUT_DIR = DATA_DIR / "input"
JSON_DIR = DATA_DIR / "json"
OUTPUT_DIR = DATA_DIR / "output"

SELLERBOARD_HTML = INPUT_DIR / "sellerboard.html"
ALL_ASINS_CSV = INPUT_DIR / "all-asins.csv"
ALREADY_REGISTERED_CSV = INPUT_DIR / "already registered.csv"
LISTINGS_JSON = JSON_DIR / "listings.json"
PRODUCT_DETAILS_JSON = JSON_DIR / "product_details.json"
ALL_ASIN_DETAILS_JSON = JSON_DIR / "all_asin_details.json"
ALL_ASIN_DETAILS_DEDUPED_JSON = JSON_DIR / "all_asin_details_deduped.json"
PRODUCT_DETAILS_XLSX = OUTPUT_DIR / "product_details.xlsx"
ALL_ASIN_DETAILS_XLSX = OUTPUT_DIR / "all_asin_details.xlsx"


def ensure_data_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    JSON_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
