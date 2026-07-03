#!/usr/bin/env python3
"""Build a sample weekly TV buy from station rate-card XML files.

Usage:
    python3 build_sample_buy.py rate_cards/*.xml --target-grps 750 -o output/sample_buy.xlsx
"""

import argparse
import glob

from broadcast_buy.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("rate_cards", nargs="+", help="Rate card XML file(s) or glob(s)")
    parser.add_argument("--target-grps", type=float, default=750.0)
    parser.add_argument("--target-demo-group", default="Adults")
    parser.add_argument("--target-demo-age", type=int, default=35)
    parser.add_argument(
        "--earliest-time", default="7:00", help="No spots before this clock time (default 7:00 AM)"
    )
    parser.add_argument(
        "--latest-time", default="23:00", help="No spots after this clock time (default 11:00 PM)"
    )
    parser.add_argument("-o", "--output", default="output/sample_buy.xlsx")
    parser.add_argument(
        "--market-name", default=None, help="Also write a Strata-importable .sbx order for this market"
    )
    parser.add_argument("--flight-start", default=None, help="Flight start date, YYYY-MM-DD (requires --market-name)")
    parser.add_argument("--flight-end", default=None, help="Flight end date, YYYY-MM-DD (requires --market-name)")
    parser.add_argument("--campaign-name", default=None, help="Campaign name for the Strata order (default: Sample Buy)")
    args = parser.parse_args()

    paths = []
    for pattern in args.rate_cards:
        matches = sorted(glob.glob(pattern))
        paths.extend(matches if matches else [pattern])

    result, log_lines, strata_path = run_pipeline(
        paths,
        args.output,
        target_grps=args.target_grps,
        target_demo_group=args.target_demo_group,
        target_demo_age=args.target_demo_age,
        earliest_time=args.earliest_time,
        latest_time=args.latest_time,
        market_name=args.market_name,
        flight_start=args.flight_start,
        flight_end=args.flight_end,
        campaign_name=args.campaign_name,
    )

    for line in log_lines:
        print(line)

    print()
    print(f"Target weekly GRPs: {args.target_grps:.0f}")
    print(f"Achieved weekly GRPs: {result.achieved_grps:.1f}")
    print(f"Total weekly cost: ${result.total_cost:,.0f}")
    if result.achieved_grps:
        print(f"Blended market CPP: ${result.total_cost / result.achieved_grps:,.2f}")
    print(f"Total spots/week: {len(result.spots)}")
    for w in result.warnings:
        print(f"NOTE: {w}")
    print(f"\nWrote {args.output}")
    if strata_path:
        print(f"Wrote {strata_path}")


if __name__ == "__main__":
    main()
