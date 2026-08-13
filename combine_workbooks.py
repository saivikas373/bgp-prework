#!/usr/bin/env python3
"""
Combines the separate IPv4, IPv6, and aggregates-review workbooks into one
file with three sheets, instead of juggling multiple files.

Usage:
    python combine_workbooks.py --ipv4 all_devices.xlsx --ipv6 all_devices_v6.xlsx \
        --aggregates aggregates_review.xlsx --out combined.xlsx
"""

import argparse

import pandas as pd


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ipv4", required=True, help="bgp_prework.py IPv4 output workbook")
    ap.add_argument("--ipv6", required=True, help="bgp_prework.py IPv6 output workbook")
    ap.add_argument("--aggregates", required=True, help="build_aggregate_review.py output workbook")
    ap.add_argument("--out", default="combined.xlsx", help="Output Excel path")
    args = ap.parse_args()

    ipv4_df = pd.read_excel(args.ipv4, engine="openpyxl")
    ipv6_df = pd.read_excel(args.ipv6, engine="openpyxl")
    agg_df = pd.read_excel(args.aggregates, engine="openpyxl")

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        ipv4_df.to_excel(writer, sheet_name="ipv4", index=False)
        ipv6_df.to_excel(writer, sheet_name="ipv6", index=False)
        agg_df.to_excel(writer, sheet_name="aggregates", index=False)

    print(f"Wrote {args.out}: ipv4 ({len(ipv4_df)} rows), ipv6 ({len(ipv6_df)} rows), "
          f"aggregates ({len(agg_df)} rows).")


if __name__ == "__main__":
    main()
