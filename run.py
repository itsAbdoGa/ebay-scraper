import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
AMAZON_DIR = ROOT / "amazon"
EBAY_DIR = ROOT / "ebay"


def run_tool(tool_dir: Path, args: list[str]) -> int:
    command = [sys.executable, str(tool_dir / "run.py"), *args]
    completed = subprocess.run(command, cwd=tool_dir)
    return completed.returncode


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gating project launcher for Amazon and eBay tools.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  python run.py\n"
            "  python run.py amazon status\n"
            "  python run.py amazon scrape --limit 2\n"
            "  python run.py ebay status\n"
            "  python run.py ebay scrape --start 2 --end 20"
        ),
    )
    parser.add_argument(
        "tool",
        nargs="?",
        choices=("amazon", "ebay"),
        help="Tool to run: amazon (gating scraper) or ebay (listings scraper)",
    )
    parser.add_argument(
        "tool_args",
        nargs=argparse.REMAINDER,
        help="Arguments passed to the selected tool",
    )
    return parser


def print_overview() -> None:
    print("Gating Project")
    print("=" * 40)
    print()
    print("Two standalone tools:")
    print(f"  Amazon gating scraper: {AMAZON_DIR / 'run.py'}")
    print(f"  eBay listings scraper: {EBAY_DIR / 'run.py'}")
    print()
    print("Examples:")
    print("  python run.py amazon status")
    print("  python run.py amazon all")
    print("  python run.py ebay status")
    print("  python run.py ebay scrape")
    print()
    print("You can also run each tool directly from its folder:")
    print("  cd amazon && python run.py status")
    print("  cd ebay && python main.py")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.tool is None:
        print_overview()
        return 0

    tool_dir = AMAZON_DIR if args.tool == "amazon" else EBAY_DIR
    tool_args = args.tool_args
    if tool_args and tool_args[0] == "--":
        tool_args = tool_args[1:]

    return run_tool(tool_dir, tool_args)


if __name__ == "__main__":
    raise SystemExit(main())
