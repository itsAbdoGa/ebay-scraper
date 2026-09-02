from datetime import datetime
from pathlib import Path
import json
import shutil

ROOT = Path(__file__).resolve().parent.parent
PROJECT_ROOT = ROOT.parent

DATA_DIR = ROOT / "data"
INPUT_DIR = DATA_DIR / "input"
JSON_DIR = DATA_DIR / "json"
OUTPUT_DIR = DATA_DIR / "output"

COMBINED_XLSX = PROJECT_ROOT / "clean" / "10krows with List 1 6-25.xlsx"
EBAY_SOURCE_HTML = INPUT_DIR / "ebay_source.html"
EBAY_COOKIES_FILE = INPUT_DIR / "ebay_cookies.txt"
WINNING_LISTINGS_XLSX = OUTPUT_DIR / "winning_listings.xlsx"
WINNING_LISTINGS_JSON = JSON_DIR / "winning_listings_live.json"
WINNING_LISTINGS_HISTORY_JSON = JSON_DIR / "winning_listings_history.json"
IDENTIFIER_NO_MATCH_HISTORY_JSON = JSON_DIR / "identifier_no_match_history.json"
LAST_STOP_JSON = JSON_DIR / "last_stop.json"
WORKBOOK_BACKUP_KEEP = 2


def ensure_data_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    JSON_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def log_workbook_stop(
    *,
    workbook_path: Path,
    excel_row: object,
    reason: str,
    script: str,
    extra: dict | None = None,
) -> None:
    ensure_data_dirs()
    try:
        row_number = int(str(excel_row).strip())
        row_label = str(row_number)
    except (TypeError, ValueError):
        row_number = None
        row_label = str(excel_row or "unknown")

    print(f"Stopped at Excel row {row_label} in {workbook_path}")
    print(f"Reason: {reason}")
    if row_number is not None:
        print(f"Resume with start_row={row_number}")

    payload = {
        "stopped_at": datetime.now().isoformat(timespec="seconds"),
        "script": script,
        "file": str(workbook_path),
        "excel_row": row_number if row_number is not None else row_label,
        "reason": reason,
    }
    if extra:
        payload.update(extra)
    try:
        LAST_STOP_JSON.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print(f"Stop location saved to {LAST_STOP_JSON}")
    except Exception as error:
        print(f"Could not save stop location: {error}")


def workbook_backup_glob(path: Path) -> str:
    return f"{path.stem}.backup-*{path.suffix}"


def prune_workbook_backups(
    path: Path,
    *,
    keep: int = WORKBOOK_BACKUP_KEEP,
) -> list[Path]:
    backups = sorted(
        (
            candidate
            for candidate in path.parent.glob(workbook_backup_glob(path))
            if candidate.is_file()
        ),
        key=lambda candidate: candidate.stat().st_mtime,
        reverse=True,
    )
    deleted: list[Path] = []
    for extra in backups[max(keep, 0) :]:
        try:
            extra.unlink()
        except OSError as error:
            print(f"Could not delete backup {extra.name}: {error}")
            continue
        deleted.append(extra)
    return deleted


def prune_directory_workbook_backups(
    directory: Path,
    *,
    keep: int = WORKBOOK_BACKUP_KEEP,
) -> list[Path]:
    deleted: list[Path] = []
    originals = [
        candidate
        for candidate in directory.glob("*.xlsx")
        if candidate.is_file()
        and ".backup-" not in candidate.name
        and not candidate.name.startswith(".")
    ]
    for original in originals:
        deleted.extend(prune_workbook_backups(original, keep=keep))
    return deleted


def create_workbook_backup(
    path: Path,
    *,
    keep: int = WORKBOOK_BACKUP_KEEP,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = path.with_name(f"{path.stem}.backup-{timestamp}{path.suffix}")
    shutil.copy2(path, backup_path)
    deleted = prune_workbook_backups(path, keep=keep)
    if deleted:
        print(f"Deleted {len(deleted)} old workbook backup(s)")
    return backup_path
