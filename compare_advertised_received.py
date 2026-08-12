#!/usr/bin/env python3
"""
Cross-checks what an ELF-side device advertises to a neighbor against what
that neighbor actually shows as received - the sanity check this whole
prework exists for: confirm nothing gets silently dropped (or unexpectedly
appears) between "what we send" and "what the other side gets", both before
the neighbor-group rename and again after, to prove the rename didn't change
behavior.

Usage:
    python compare_advertised_received.py --pre bgppre.xlsx --rcr bgprcr.xlsx \
        --pairs pairs.csv --out reconciliation.xlsx

--pre  : workbook from the ELF side (bgp_prework.py output), advertised (out) rows
--rcr  : workbook from the upstream/RCR side, received (in) rows
--pairs: CSV mapping which ELF host/neighbor pairs with which RCR host/neighbor
         (see pairs.example.csv) - copy it to pairs.csv and fill in your real
         hostnames/neighbor IPs.

Both --pre and --rcr can point to the same file if you ran one combined pull
that covers both directions/devices - the script just filters by
hostname+neighbor+direction, it doesn't care how many files that data is
split across.
"""

import argparse
import csv
import sys

import pandas as pd


def load_pairs(path):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def prefix_set(df, hostname, neighbor, direction):
    subset = df[(df["hostname"] == hostname) & (df["neighbor"].astype(str) == neighbor)
                & (df["direction"] == direction)]
    return subset, set(subset["prefix"].dropna())


def row_for_prefix(df, prefix):
    """First matching row for a prefix, as a dict - used to pull as_path/communities/etc
    for the reconciliation sheet regardless of which side (adv or rec) has it."""
    match = df[df["prefix"] == prefix]
    if match.empty:
        return {}
    return match.iloc[0].to_dict()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pre", required=True, help="Workbook with ELF-side advertised (out) rows")
    ap.add_argument("--rcr", required=True, help="Workbook with upstream-side received (in) rows")
    ap.add_argument("--pairs", default="pairs.csv",
                    help="CSV: elf_hostname,elf_neighbor,rcr_hostname,rcr_neighbor,label")
    ap.add_argument("--out", default="reconciliation.xlsx", help="Output Excel path")
    args = ap.parse_args()

    pre = pd.read_excel(args.pre)
    rcr = pd.read_excel(args.rcr)
    pairs = load_pairs(args.pairs)

    detail_rows = []
    summary_rows = []

    for pair in pairs:
        elf_host, elf_nb = pair["elf_hostname"], pair["elf_neighbor"]
        rcr_host, rcr_nb = pair["rcr_hostname"], pair["rcr_neighbor"]
        label = pair.get("label") or f"{elf_host} -> {rcr_host}"

        adv_df, adv_prefixes = prefix_set(pre, elf_host, elf_nb, "advertised (out)")
        rec_df, rec_prefixes = prefix_set(rcr, rcr_host, rcr_nb, "received (in)")

        matched = adv_prefixes & rec_prefixes
        missing = adv_prefixes - rec_prefixes   # advertised but not confirmed received - investigate
        unexpected = rec_prefixes - adv_prefixes  # received but not in the advertised list - investigate

        print(f"[{label}] advertised={len(adv_prefixes)} received={len(rec_prefixes)} "
              f"matched={len(matched)} missing={len(missing)} unexpected={len(unexpected)}",
              file=sys.stderr)

        summary_rows.append({
            "pair": label, "elf_hostname": elf_host, "elf_neighbor": elf_nb,
            "rcr_hostname": rcr_host, "rcr_neighbor": rcr_nb,
            "advertised_count": len(adv_prefixes), "received_count": len(rec_prefixes),
            "matched_count": len(matched), "missing_count": len(missing),
            "unexpected_count": len(unexpected),
        })

        for prefix in sorted(missing):
            src = row_for_prefix(adv_df, prefix)
            detail_rows.append({
                "pair": label, "prefix": prefix, "status": "MISSING - advertised but not received",
                "as_path": src.get("as_path", ""), "communities": src.get("communities", ""),
                "Agg": src.get("Agg", ""), "IP Classification": src.get("IP Classification", ""),
            })
        for prefix in sorted(unexpected):
            src = row_for_prefix(rec_df, prefix)
            detail_rows.append({
                "pair": label, "prefix": prefix, "status": "UNEXPECTED - received but not advertised",
                "as_path": src.get("as_path", ""), "communities": src.get("communities", ""),
                "Agg": src.get("Agg", ""), "IP Classification": src.get("IP Classification", ""),
            })
        for prefix in sorted(matched):
            src = row_for_prefix(adv_df, prefix) or row_for_prefix(rec_df, prefix)
            detail_rows.append({
                "pair": label, "prefix": prefix, "status": "matched",
                "as_path": src.get("as_path", ""), "communities": src.get("communities", ""),
                "Agg": src.get("Agg", ""), "IP Classification": src.get("IP Classification", ""),
            })

    summary_df = pd.DataFrame(summary_rows)
    detail_df = pd.DataFrame(detail_rows, columns=[
        "pair", "prefix", "status", "as_path", "communities", "Agg", "IP Classification",
    ])

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        detail_df.to_excel(writer, sheet_name="detail", index=False)
        mismatches_only = detail_df[detail_df["status"] != "matched"]
        mismatches_only.to_excel(writer, sheet_name="mismatches_only", index=False)

    total_missing = summary_df["missing_count"].sum() if not summary_df.empty else 0
    total_unexpected = summary_df["unexpected_count"].sum() if not summary_df.empty else 0
    print(f"\nWrote {args.out} - {total_missing} missing, {total_unexpected} unexpected "
          f"across {len(pairs)} pair(s). Check the 'mismatches_only' sheet first.",
          file=sys.stderr)


if __name__ == "__main__":
    main()
