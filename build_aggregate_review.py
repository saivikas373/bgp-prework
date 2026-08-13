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
    df = df[df["prefix"].notna() & (df["prefix"] != "")]

    RESERVED_LABELS = {"RFC 1918", "CGNAT (RFC 6598)", "ULA (RFC 4193)", "Link-Local"}

    matched = df[df["Agg"].notna() & (df["Agg"] != "")]

    # Every prefix we advertise that has no Agg match - shown regardless of
    # classification, so nothing real gets silently hidden from this list.
    unmatched = df[
        (df["Agg"].isna() | (df["Agg"] == ""))
        & (df["direction"] == "advertised (out)")
        & (~df["prefix"].isin(["0.0.0.0/0", "::/0"]))
    ].copy()
    unmatched["Agg"] = unmatched["prefix"]  # candidate: the prefix is its own (untracked) aggregate

    rows = []
    for agg, group in pd.concat([matched, unmatched]).groupby("Agg"):
        prefixes = sorted(group["prefix"].dropna().unique())
        hostnames = sorted(group["hostname"].dropna().unique())
        directions = sorted(group["direction"].dropna().unique())
        classifications_seen = sorted(group["IP Classification"].dropna().unique())
        communities_seen = sorted({c for c in group["communities"].dropna().unique() if c})
        in_known_aggregates = agg not in set(unmatched["Agg"])

        # RFC1918/CGNAT/ULA/link-local are auto-classified and never
        # internet-routable regardless - pre-fill rather than mark TBD, so
        # you're not asked to make a Public/Private GUA call that doesn't
        # apply to them, but they're still visible in the list.
        reserved_labels_here = [c for c in classifications_seen if c in RESERVED_LABELS]
        classification = reserved_labels_here[0] if reserved_labels_here else "TBD"

        rows.append({
            "aggregate_cidr": agg,
            "in_aggregates_csv": "yes" if in_known_aggregates else "NO - add this",
            "classification": classification,  # TBD -> fill in Public / Private GUA
            "prefix_count": len(prefixes),
            "prefixes": ", ".join(prefixes),
            "hostnames": ", ".join(hostnames),
            "directions_seen": ", ".join(directions),
            "current_classification_in_data": ", ".join(classifications_seen),
            "communities_seen": "; ".join(communities_seen),
        })

    out_df = pd.DataFrame(rows, columns=[
        "aggregate_cidr", "in_aggregates_csv", "classification", "prefix_count", "prefixes",
        "hostnames", "directions_seen", "current_classification_in_data", "communities_seen",
    ]).sort_values(["in_aggregates_csv", "aggregate_cidr"])

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="aggregates", index=False)

    untracked_count = (out_df["in_aggregates_csv"] != "yes").sum()
    needs_decision = (out_df["classification"] == "TBD").sum()
    print(f"Wrote {len(out_df)} unique aggregate blocks to {args.out} "
          f"({untracked_count} not yet in aggregates.csv - marked 'NO - add this'; "
          f"{needs_decision} still need a real Public/Private GUA decision - marked TBD; "
          f"the rest are RFC1918/CGNAT/etc, pre-filled since those never apply). "
          f"Fill in the remaining TBDs, then feed it back into aggregates.csv for future runs.")


if __name__ == "__main__":
    main()
