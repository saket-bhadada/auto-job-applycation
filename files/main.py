"""CLI entry point.

  python main.py --discover "Find 5 Summer 2027 Backend or AI Internships in India"
  python main.py --watch
  python main.py --all ["optional discovery request"]
"""
from __future__ import annotations

import argparse
import logging
import sys

import config
from auto_applier import AutoApplier
from excel_manager import ExcelLockedError, ExcelManager

DEFAULT_QUERY = "Find 5 Summer 2027 Backend or AI Internships in India"


def main() -> int:
    ap = argparse.ArgumentParser(description="Internship discovery + Excel-approved auto-apply pipeline")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--discover", metavar="REQUEST", help="Run discovery with this request")
    g.add_argument("--watch", action="store_true", help="Watch Excel for YES/NO and auto-apply")
    g.add_argument("--all", nargs="?", const=DEFAULT_QUERY, metavar="REQUEST",
                   help="Discovery, then watch mode (optional custom request)")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
                        datefmt="%H:%M:%S")
    needs_apply = bool(args.watch or args.all)
    problems = config.validate(for_apply=needs_apply)
    if args.watch:  # watch mode doesn't need API keys
        problems = [p for p in problems if "API_KEY" not in p]
    if problems:
        for p in problems:
            logging.error("Config: %s", p)
        return 1

    try:
        excel = ExcelManager()
        request = args.discover or args.all
        if request:
            from agent_discovery import DiscoveryAgent  # lazy: not needed for --watch
            DiscoveryAgent(excel).run(request)
        if needs_apply:
            AutoApplier(excel).watch()
    except ExcelLockedError as exc:
        logging.error("%s", exc)
        return 2
    except KeyboardInterrupt:
        logging.info("Interrupted.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
