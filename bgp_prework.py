#!/usr/bin/env python3
"""
BGP prework data collector.

Pulls per-neighbor BGP state (neighbor-group, remote-AS, advertised/received
prefixes, communities, aggregate match, public/private) from Cisco IOS-XR/IOS-XE
routers and writes it to an Excel workbook. Built for the prework before
standardizing neighbor-group naming on the remaining 2 locations, so you can
diff "what was advertised before" vs "after" and confirm the NOC team's
route manipulation still works post-rename.

Usage:
    python bgp_prework.py --devices devices.csv --aggregates aggregates.csv --out prework.xlsx
    python bgp_prework.py --sample-dir ./samples --aggregates aggregates.csv --out prework.xlsx

Credentials: set NETOPS_USERNAME / NETOPS_PASSWORD env vars, or you'll be
prompted (password via getpass, never echoed or stored to disk).
"""

import argparse
import csv
import getpass
import ipaddress
import os
import re
import sys
from pathlib import Path

import pandas as pd

try:
    from netmiko import ConnectHandler
except ImportError:
    ConnectHandler = None


RESERVED_NETWORKS = [
    (ipaddress.ip_network("10.0.0.0/8"), "RFC 1918"),
    (ipaddress.ip_network("172.16.0.0/12"), "RFC 1918"),
    (ipaddress.ip_network("192.168.0.0/16"), "RFC 1918"),
    (ipaddress.ip_network("100.64.0.0/10"), "CGNAT (RFC 6598)"),
    (ipaddress.ip_network("169.254.0.0/16"), "Link-Local"),
    (ipaddress.ip_network("fc00::/7"), "ULA (RFC 4193)"),
    (ipaddress.ip_network("fe80::/10"), "Link-Local"),
]


def _matching_aggregates(prefix_str, aggregates):
    try:
        net = ipaddress.ip_network(prefix_str, strict=False)
    except ValueError:
        return []
    matches = [a for a in aggregates if a["net"].version == net.version and net.subnet_of(a["net"])]
    matches.sort(key=lambda a: a["net"].prefixlen, reverse=True)
    return matches


def find_aggregate(prefix_str, aggregates):
    matches = _matching_aggregates(prefix_str, aggregates)
    return str(matches[0]["net"]) if matches else ""


def classify_ip(prefix_str, aggregates):
    """
    IP Classification per your team's taxonomy:
      - RFC1918 / CGNAT / ULA / link-local -> labeled by which reserved range
        it actually falls in (auto-detected, no input needed).
      - Matches a block in aggregates.csv -> whatever classification you set
        there (e.g. "Private GUA" for your own public-IP space marked
        internal-only, "Public" for freely-advertised space) - this has to
        come from your team, since "don't leak to internet" is a business
        decision, not something a WHOIS/registry lookup can tell you.
      - No match -> "Public" (assumed external/internet-routable)
    """
    try:
        net = ipaddress.ip_network(prefix_str, strict=False)
    except ValueError:
        return "unknown"
    for reserved, label in RESERVED_NETWORKS:
        if reserved.version == net.version and net.subnet_of(reserved):
            return label
    matches = _matching_aggregates(prefix_str, aggregates)
    if matches and matches[0].get("classification"):
        return matches[0]["classification"]
    return "Public"


def load_aggregates(path):
    aggregates = []
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            try:
                net = ipaddress.ip_network(row["aggregate_cidr"].strip(), strict=False)
            except ValueError:
                continue
            aggregates.append({
                "location": row["location"], "net": net,
                "classification": (row.get("classification") or "").strip(),
                "description": row.get("description", ""),
            })
    return aggregates


def load_devices(path):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------------------
# IOS-XR output parsing
# NOTE: these regexes match typical IOS-XR formatting. Real output varies by
# release/platform - validate against a captured sample (--sample-dir) before
# trusting results from a live run.
# ---------------------------------------------------------------------------

def _parse_blocks(config_text, block_start_re):
    """
    Generic parser for `neighbor <ip> { ... }` or `neighbor-group <name> { ... }`
    blocks, pulling out use-neighbor-group / remote-as / description /
    route-policy in / route-policy out. Returns {key: {...}}.
    """
    blocks = {}
    current_key = None
    for raw_line in config_text.splitlines():
        line = raw_line.rstrip()
        m = block_start_re.match(line)
        if m:
            current_key = m.group(1)
            blocks[current_key] = {
                "neighbor_group": "", "remote_as": "", "description": "",
                "policy_in": "", "policy_out": "",
            }
            continue
        if current_key is None:
            continue
        m = re.search(r"use neighbor-group (\S+)", line)
        if m:
            blocks[current_key]["neighbor_group"] = m.group(1)
        m = re.search(r"remote-as (\S+)", line)
        if m:
            blocks[current_key]["remote_as"] = m.group(1)
        m = re.search(r"description (.+)", line)
        if m:
            blocks[current_key]["description"] = m.group(1).strip()
        m = re.search(r"route-policy (\S+) in", line)
        if m:
            blocks[current_key]["policy_in"] = m.group(1)
        m = re.search(r"route-policy (\S+) out", line)
        if m:
            blocks[current_key]["policy_out"] = m.group(1)
        # a bare top-level line that isn't part of this block ends it
        if raw_line and not raw_line[0].isspace():
            current_key = None
    return blocks


def parse_bgp_running_config(config_text):
    """
    Parses `show running-config router bgp` into {neighbor_ip: {...}}.

    IOS-XR neighbors typically inherit remote-as/route-policy from a
    neighbor-group via `use neighbor-group X` rather than repeating them:

        neighbor-group IPV4-CORE
         remote-as 65001
         address-family ipv4 unicast
          route-policy RP-IN in
          route-policy RP-OUT out
        !
        neighbor 192.0.2.1
         use neighbor-group IPV4-CORE
         description PEER-XYZ

    So this parses neighbor-group blocks and neighbor blocks separately, then
    fills in any field left blank on the neighbor from its neighbor-group.
    """
    neighbor_group_re = re.compile(r"^\s*neighbor-group\s+(\S+)\s*$")
    neighbor_re = re.compile(r"^\s*neighbor\s+(\S+)\s*$")

    groups = _parse_blocks(config_text, neighbor_group_re)
    neighbors = _parse_blocks(config_text, neighbor_re)

    for ip, cfg in neighbors.items():
        group_name = cfg.get("neighbor_group", "")
        group_cfg = groups.get(group_name, {})
        for field in ("remote_as", "policy_in", "policy_out"):
            if not cfg.get(field):
                cfg[field] = group_cfg.get(field, "")

    return neighbors


def parse_route_policy_communities(config_text):
    """
    Extracts `set community (...)` actions per route-policy name, e.g.:
        route-policy RP-OUT
         set community (65000:100) additive
        end-policy
    Returns {policy_name: [community strings]}.

    Caveat: this is a static read of the policy text, not per-prefix. If a
    policy sets different communities under conditional branches (if/elseif),
    all communities in the policy are listed against every neighbor using it -
    it won't tell you which prefix got which community. For that level of
    precision, add a per-prefix `show bgp ... <prefix> detail` pull.
    """
    policies = {}
    current_policy = None
    for line in config_text.splitlines():
        m = re.match(r"^\s*route-policy\s+(\S+)\s*$", line)
        if m:
            current_policy = m.group(1)
            policies[current_policy] = []
            continue
        if current_policy is None:
            continue
        if re.match(r"^\s*end-policy\s*$", line):
            current_policy = None
            continue
        m = re.search(r"set community\s+\(([^)]+)\)", line)
        if m:
            policies[current_policy].append(m.group(1).strip())
    return policies


ROUTE_LINE_RE = re.compile(
    r"^\s*(?P<prefix>\d+\.\d+\.\d+\.\d+/\d+|[0-9a-fA-F:]+/\d+)\s+(?P<nexthop>\S+)\s+(?P<rest>.+?)\s*$"
)


def parse_routes(show_routes_text):
    """
    Parses `advertised-routes` / `received-routes` output -> [{prefix, next_hop, as_path}].

    Real IOS-XR format (confirmed against live output, not the classic-IOS
    metric/locprf/weight table this originally assumed):
        Network            Next Hop        From            AS Path
        10.1.194.0/23      24.93.64.1      24.27.242.37    19115?
        24.27.242.0/23     24.93.64.1      Local Aggregate 19115i

    "From" can be 1 or 2 words (a peer IP, "Local", or "Local Aggregate"), so
    rather than guess its width, we take the last whitespace token as the AS
    path (with origin code attached) and ignore "From" - it isn't part of the
    target sheet schema.
    """
    routes = []
    for line in show_routes_text.splitlines():
        m = ROUTE_LINE_RE.match(line)
        if not m:
            continue
        rest_tokens = m.group("rest").split()
        as_path = rest_tokens[-1] if rest_tokens else ""
        routes.append({"prefix": m.group("prefix"), "next_hop": m.group("nexthop"), "as_path": as_path})
    return routes


def parse_prefix_detail_communities(detail_text):
    """
    Parses `show bgp ipv4/ipv6 unicast <prefix>` output into the community
    string of the *best* path - the one actually being advertised - as a
    semicolon-joined string (matching your example sheet's format, e.g.
    "20115:3101;20115:64080").

    The output has one `Path #N:` block per candidate route; only one is
    marked "best" in its status line (e.g. "valid, external, best,
    group-best, multipath"). Falls back to the first path with a Community
    line if no path is explicitly marked best.
    """
    blocks = re.split(r"(?=^Path #\d+:)", detail_text, flags=re.MULTILINE)
    best_communities = None
    first_communities = None
    for block in blocks:
        if not block.strip().startswith("Path #"):
            continue
        comm_match = re.search(r"^Community:\s*(.+)$", block, flags=re.MULTILINE)
        communities = comm_match.group(1).strip() if comm_match else ""
        if first_communities is None:
            first_communities = communities
        if re.search(r"\bbest\b", block):
            best_communities = communities
            break
    chosen = best_communities if best_communities is not None else (first_communities or "")
    return chosen.replace(" ", ";")


def parse_bgp_summary(show_summary_text):
    """Parses `show bgp ipv4/ipv6 unicast summary` -> list of neighbor IPs."""
    ips = []
    started = False
    for line in show_summary_text.splitlines():
        if line.strip().startswith("Neighbor"):
            started = True
            continue
        if not started:
            continue
        parts = line.split()
        if not parts:
            continue
        try:
            ipaddress.ip_address(parts[0])
        except ValueError:
            continue
        ips.append(parts[0])
    return ips


# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------

def collect_from_device(conn_or_dir, hostname, use_sample_dir=False):
    outputs = {}
    if use_sample_dir:
        base = Path(conn_or_dir) / hostname
        for key, fname in [("summary_v4", "summary_v4.txt"),
                            ("summary_v6", "summary_v6.txt"),
                            ("running_bgp", "running_bgp.txt")]:
            fp = base / fname
            outputs[key] = fp.read_text() if fp.exists() else ""
    else:
        conn = conn_or_dir
        outputs["summary_v4"] = conn.send_command("show bgp ipv4 unicast summary")
        outputs["summary_v6"] = conn.send_command("show bgp ipv6 unicast summary")
        outputs["running_bgp"] = conn.send_command("show running-config router bgp")
    return outputs


def collect_routes_for_neighbor(conn_or_dir, hostname, ip, afi, use_sample_dir=False):
    if use_sample_dir:
        base = Path(conn_or_dir) / hostname
        adv_file = base / f"advertised_{ip.replace(':', '_')}.txt"
        rec_file = base / f"received_{ip.replace(':', '_')}.txt"
        adv = adv_file.read_text() if adv_file.exists() else ""
        rec = rec_file.read_text() if rec_file.exists() else ""
    else:
        conn = conn_or_dir
        adv = conn.send_command(f"show bgp {afi} unicast neighbors {ip} advertised-routes")
        try:
            rec = conn.send_command(f"show bgp {afi} unicast neighbors {ip} received-routes")
        except Exception:
            # needs `bgp neighbor soft-reconfiguration inbound` or route-refresh capability
            rec = ""
    return adv, rec


def collect_prefix_communities(conn_or_dir, hostname, prefix, afi, use_sample_dir=False):
    """
    Per-prefix community lookup via `show bgp <afi> unicast <prefix>`. Only
    called when --with-communities is set, since it's one extra command per
    prefix (can be hundreds across a full neighbor table - slow but exact).
    """
    if use_sample_dir:
        base = Path(conn_or_dir) / hostname
        fname = f"prefix_detail_{prefix.replace('/', '_').replace(':', '_')}.txt"
        fp = base / fname
        text = fp.read_text() if fp.exists() else ""
    else:
        conn = conn_or_dir
        text = conn.send_command(f"show bgp {afi} unicast {prefix}")
    return parse_prefix_detail_communities(text)


def build_rows(location, hostname, neighbor_cfg, communities_by_policy, aggregates,
               summary_v4_ips, summary_v6_ips, conn_or_dir, use_sample_dir, with_communities):
    rows = []
    all_ips = [(ip, "ipv4") for ip in summary_v4_ips] + [(ip, "ipv6") for ip in summary_v6_ips]

    for ip, afi in all_ips:
        cfg = neighbor_cfg.get(ip, {})
        adv_text, rec_text = collect_routes_for_neighbor(conn_or_dir, hostname, ip, afi, use_sample_dir)
        adv_routes = parse_routes(adv_text)
        rec_routes = parse_routes(rec_text)

        # Static fallback: communities as configured in the route-policy text.
        # Overridden per-prefix below when --with-communities pulls the exact
        # best-path community from `show bgp ... <prefix>`.
        static_comm_out = "; ".join(communities_by_policy.get(cfg.get("policy_out", ""), []))

        base_row = {
            "location": location, "hostname": hostname, "neighbor": ip,
            "IPv4/IPv6": "IPv4" if afi == "ipv4" else "IPv6",
            "neighbor_group": cfg.get("neighbor_group", ""),
            "remote_as": cfg.get("remote_as", ""), "description": cfg.get("description", ""),
            "policy_out": cfg.get("policy_out", ""), "policy_in": cfg.get("policy_in", ""),
            "COMMENT": cfg.get("policy_out", ""),
        }

        if not adv_routes and not rec_routes:
            rows.append({**base_row, "direction": "", "prefix": "", "nexthop": "", "as_path": "",
                         "Agg": "", "IP Classification": "", "communities": ""})
            continue

        for r in adv_routes:
            communities = (collect_prefix_communities(conn_or_dir, hostname, r["prefix"], afi, use_sample_dir)
                           if with_communities else static_comm_out)
            rows.append({**base_row, "direction": "advertised (out)", "prefix": r["prefix"],
                         "nexthop": r["next_hop"], "as_path": r["as_path"],
                         "Agg": find_aggregate(r["prefix"], aggregates),
                         "IP Classification": classify_ip(r["prefix"], aggregates),
                         "communities": communities})
        for r in rec_routes:
            rows.append({**base_row, "direction": "received (in)", "prefix": r["prefix"],
                         "nexthop": r["next_hop"], "as_path": r["as_path"],
                         "Agg": find_aggregate(r["prefix"], aggregates),
                         "IP Classification": classify_ip(r["prefix"], aggregates),
                         "communities": ""})
    return rows


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--devices", default="devices.csv", help="CSV: location,hostname,mgmt_ip,device_type,port")
    ap.add_argument("--aggregates", default="aggregates.csv",
                    help="CSV: location,aggregate_cidr,classification,description")
    ap.add_argument("--out", default="bgp_prework.xlsx", help="Output Excel path")
    ap.add_argument("--sample-dir", default=None,
                    help="Read saved show-command text files from this dir instead of SSH (for tuning the parser)")
    ap.add_argument("--with-communities", action="store_true",
                    help="Pull exact best-path community per advertised prefix via "
                         "`show bgp ... <prefix>` (one extra command per prefix - slow "
                         "on large tables). Without this flag, communities column falls "
                         "back to the static route-policy-level community list.")
    args = ap.parse_args()

    aggregates = load_aggregates(args.aggregates)
    devices = load_devices(args.devices)
    use_sample_dir = args.sample_dir is not None

    username = password = None
    if not use_sample_dir:
        if ConnectHandler is None:
            sys.exit("netmiko is not installed. Run: pip install -r requirements.txt")
        username = os.environ.get("NETOPS_USERNAME") or input("Username: ")
        password = os.environ.get("NETOPS_PASSWORD") or getpass.getpass("Password: ")

    all_rows = []

    for dev in devices:
        location, hostname = dev["location"], dev["hostname"]
        print(f"[{location}] {hostname}: collecting...", file=sys.stderr)

        if use_sample_dir:
            conn_or_dir = args.sample_dir
            outputs = collect_from_device(conn_or_dir, hostname, use_sample_dir=True)
        else:
            conn = ConnectHandler(
                device_type=dev.get("device_type", "cisco_xr"),
                host=dev["mgmt_ip"], username=username, password=password,
                port=int(dev.get("port") or 22),
            )
            conn_or_dir = conn
            outputs = collect_from_device(conn, hostname, use_sample_dir=False)

        neighbor_cfg = parse_bgp_running_config(outputs["running_bgp"])
        communities_by_policy = parse_route_policy_communities(outputs["running_bgp"])
        summary_v4_ips = parse_bgp_summary(outputs["summary_v4"])
        summary_v6_ips = parse_bgp_summary(outputs["summary_v6"])

        rows = build_rows(location, hostname, neighbor_cfg, communities_by_policy, aggregates,
                          summary_v4_ips, summary_v6_ips, conn_or_dir, use_sample_dir,
                          args.with_communities)
        all_rows.extend(rows)

        if not use_sample_dir:
            conn_or_dir.disconnect()

    df = pd.DataFrame(all_rows, columns=[
        "location", "hostname", "neighbor", "prefix", "nexthop", "as_path", "communities",
        "COMMENT", "IPv4/IPv6", "Agg", "IP Classification",
        "direction", "neighbor_group", "remote_as", "description", "policy_out", "policy_in",
    ])

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="bgp_prework", index=False)
        if not df.empty:
            grp = df.groupby(["location", "neighbor_group", "neighbor"]).size().reset_index(name="route_count")
            grp.to_excel(writer, sheet_name="neighbor_group_summary", index=False)

    print(f"Wrote {len(df)} rows to {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
