from __future__ import annotations

import argparse

from app import app, cleanup_false_positive_candidates


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean likely non-person candidate-review PIP records.")
    parser.add_argument("--delete", action="store_true", help="Delete suspect records instead of marking them as not a person.")
    args = parser.parse_args()

    with app.app_context():
        count = cleanup_false_positive_candidates(delete=args.delete)
        action = "deleted" if args.delete else "marked as not a person"
        print(f"{count} likely non-person candidate record(s) {action}.")


if __name__ == "__main__":
    main()
