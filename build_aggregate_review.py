#!/usr/bin/env python3
"""
Segregates a bgp_prework.py output workbook by the "Agg" column - one row
per unique aggregate block your team actually owns/originates, pulled
straight from real run data, with a blank classification column for you to
mark Public vs Private GUA. This is the review step before finalizing
aggregates.csv: instead of guessing which blocks exist, see exactly which
ones showed up across all your devices/prefixes in this run.

Usage:
    python build_aggregate_review.py --input all_devices.xlsx --out aggregates_review.xlsx
"""

import argparse

import pandas as pd


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", required=True, help="bgp_prework.py output workbook (e.g. all_devices.xlsx)")
    ap.add_argument("--out", default="aggregates_review.xlsx", help="Output Excel path")
    args = ap.parse_args()

    df = pd.read_excel(args.input, engine="openpyxl")
    df = df[df["Agg"].notna() & (df["Agg"] != "")]

    rows = []
    for agg, group in df.groupby("Agg"):
        prefixes = sorted(group["prefix"].dropna().unique())
        hostnames = sorted(group["hostname"].dropna().unique())
        directions = sorted(group["direction"].dropna().unique())
        classifications_seen = sorted(group["IP Classification"].dropna().unique())
        communities_seen = sorted({c for c in group["communities"].dropna().unique() if c})

        rows.append({
            "aggregate_cidr": agg,
            "classification": "TBD",  # fill in: Public / Private GUA
            "prefix_count": len(prefixes),
            "prefixes": ", ".join(prefixes),
            "hostnames": ", ".join(hostnames),
            "directions_seen": ", ".join(directions),
            "current_classification_in_data": ", ".join(classifications_seen),
            "communities_seen": "; ".join(communities_seen),
        })

    out_df = pd.DataFrame(rows, columns=[
        "aggregate_cidr", "classification", "prefix_count", "prefixes",
        "hostnames", "directions_seen", "current_classification_in_data", "communities_seen",
    ]).sort_values("aggregate_cidr")

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="aggregates", index=False)

    print(f"Wrote {len(out_df)} unique aggregate blocks to {args.out}. "
          f"Fill in the 'classification' column (TBD -> Public or Private GUA), "
          f"then feed it back into aggregates.csv for future runs.")


if __name__ == "__main__":
    main()
