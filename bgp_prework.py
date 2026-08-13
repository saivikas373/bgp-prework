#!/usr/bin/env python3
"""
BGP prework data collector.

Pulls per-neighbor BGP state (neighbor-group, remote-AS, advertised/received
prefixes, communities, aggregate match, public/private) from Cisco IOS-XR/IOS-XE
routers, or from Junos neighbors (device_type containing "juniper"/"junos" in
devices.csv), and writes it to an Excel workbook. Built for the prework before
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


def _is_subnet_of(inner, outer):
    """
    inner.subnet_of(outer) equivalent, without relying on
    ipaddress.subnet_of() (added in Python 3.7 - jump server runs 3.6).
    """
    if inner.version != outer.version:
        return False
    return (int(outer.network_address) <= int(inner.network_address)
            and int(inner.broadcast_address) <= int(outer.broadcast_address))


def _matching_aggregates(prefix_str, aggregates):
    try:
        net = ipaddress.ip_network(prefix_str, strict=False)
    except ValueError:
        return []
    matches = [a for a in aggregates if _is_subnet_of(net, a["net"])]
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
        if _is_subnet_of(net, reserved):
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


def device_platform(dev):
    """netmiko device_type -> which parser family to use ('junos' or 'iosxr')."""
    dt = (dev.get("device_type") or "").lower()
    if "juniper" in dt or "junos" in dt:
        return "junos"
    return "iosxr"


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


def parse_bgp_originations(config_text):
    """
    Extracts what this router actually originates/aggregates, straight from
    the top-level (not per-neighbor) address-family blocks:

        address-family ipv4 unicast
         network 22.80.64.0/19 route-policy LOCAL_PREF_50_POLICY
         network 24.27.242.36/31
         aggregate-address 24.27.242.0/23
         aggregate-address 66.56.248.0/22 summary-only

    Returns a list of {cidr, kind, route_policy, summary_only} - this is the
    candidate list for aggregates.csv. Classification (Public vs Private GUA)
    still has to come from you/NOC - this only finds *what* is originated,
    not whether it should be hidden from the internet.

    Only matches inside a bare `address-family` block (i.e. not indented
    under a `neighbor` or `neighbor-group` block) so per-neighbor policy
    lines aren't mistaken for global originations.
    """
    originations = []
    in_global_af = False
    indent_stack = 0
    for raw_line in config_text.splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()
        indent = len(line) - len(line.lstrip())

        if re.match(r"^\s*(neighbor|neighbor-group)\s+\S+\s*$", line):
            in_global_af = False
            continue
        if re.match(r"^\s*address-family\s+\S+\s+\S+\s*$", line):
            # only "global" if it's at the router bgp's own indent level (1),
            # not nested deeper inside a neighbor/neighbor-group block
            in_global_af = indent <= 1
            continue
        if not stripped or stripped == "!":
            if indent <= 1:
                in_global_af = False
            continue
        if not in_global_af:
            continue

        m = re.match(r"network\s+(\S+)(?:\s+route-policy\s+(\S+))?", stripped)
        if m:
            originations.append({
                "cidr": m.group(1), "kind": "network",
                "route_policy": m.group(2) or "", "summary_only": False,
            })
            continue
        m = re.match(r"aggregate-address\s+(\S+)(\s+summary-only)?", stripped)
        if m:
            originations.append({
                "cidr": m.group(1), "kind": "aggregate-address",
                "route_policy": "", "summary_only": bool(m.group(2)),
            })
    return originations


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
    r"^\s*(?:[*>sdhirSN]+\s+)?(?P<prefix>\d+\.\d+\.\d+\.\d+/\d+|[0-9a-fA-F:]+/\d+)"
    r"\s+(?P<nexthop>\S+)\s+(?P<rest>.+?)\s*$"
)


def parse_routes(show_routes_text):
    """
    Parses `advertised-routes` / `received routes` output -> [{prefix, next_hop, as_path}].

    IOS-XR uses TWO DIFFERENT table formats depending on the command
    (confirmed against real output from both) - `advertised-routes` uses a
    simple table with no status markers and a concatenated AS+origin:
        Network            Next Hop        From            AS Path
        10.1.194.0/23      24.93.64.1      24.27.242.37    19115?
        24.27.242.0/23     24.93.64.1      Local Aggregate 19115i

    but `received routes` (note: two words, not hyphenated - see
    collect_routes_for_neighbor) uses the classic full BGP table format,
    with leading status markers (*, >, s, d, h, i, r, S, N - combinable,
    e.g. "*>" for valid+best) and a separate Metric/LocPrf/Weight/AS-path
    with the origin code space-separated instead of concatenated:
        Status codes: s suppressed, d damped, h history, * valid, > best
        Origin codes: i - IGP, e - EGP, ? - incomplete
           Network            Next Hop            Metric LocPrf Weight Path
        *> 10.1.194.0/23      24.93.64.3                             0 19115 ?
        *> 22.80.64.0/19      24.93.64.3               0             0 19115 i

    The optional leading status-marker group handles the second format; for
    the AS-path, if the last token is a bare origin code (space-separated
    from the AS number), the trailing "<asn> <origin>" pair is taken
    together - otherwise (concatenated, or no origin code present) just the
    last token is used. "From"/Metric/LocPrf/Weight columns are dropped
    either way - not part of the target sheet schema.
    """
    routes = []
    for line in show_routes_text.splitlines():
        m = ROUTE_LINE_RE.match(line)
        if not m:
            continue
        rest_tokens = m.group("rest").split()
        if len(rest_tokens) >= 2 and rest_tokens[-1] in ("I", "E", "?", "i", "e"):
            as_path = " ".join(rest_tokens[-2:])
        elif rest_tokens:
            as_path = rest_tokens[-1]
        else:
            as_path = ""
        routes.append({"prefix": m.group("prefix"), "next_hop": m.group("nexthop"), "as_path": as_path})
    return routes


def parse_prefix_detail_communities(detail_text):
    """
    Parses `show bgp ipv4/ipv6 unicast <prefix>` output into the community
    string of the *best* path - the one actually being advertised - as a
    semicolon-joined string (matching your example sheet's format, e.g.
    "20115:3101;20115:64080").

    Multi-path output has one `Path #N:` block per candidate route; only one
    is marked "best" in its status line (e.g. "valid, external, best,
    group-best, multipath"). Falls back to the first path with a Community
    line if no path is explicitly marked best.

    Single-path entries (e.g. a locally-originated `network`/`aggregate-
    address` route with only one path) may not use "Path #N:" labeling at
    all - just one straightforward "Community:" line. If no "Path #N:"
    blocks are found, fall back to the first "Community:" line anywhere in
    the text rather than returning nothing.
    """
    # re.split() on a zero-width lookahead needs Python 3.7+; the jump
    # server runs 3.6, so slice on match start positions manually instead.
    # "Path #N:" and "Community:" lines are NOT reliably flush-left - real
    # captured output showed some prefixes indented (2 spaces) and others
    # not, on the same device/command. Match regardless of leading
    # whitespace rather than assuming column 0.
    starts = [m.start() for m in re.finditer(r"^\s*Path #\d+:", detail_text, flags=re.MULTILINE)]
    if not starts:
        comm_match = re.search(r"^\s*Community:\s*(.+)$", detail_text, flags=re.MULTILINE)
        return comm_match.group(1).strip().replace(" ", ";") if comm_match else ""

    blocks = [detail_text[start:(starts[i + 1] if i + 1 < len(starts) else len(detail_text))]
              for i, start in enumerate(starts)]
    best_communities = None
    first_communities = None
    for block in blocks:
        if not block.strip().startswith("Path #"):
            continue
        comm_match = re.search(r"^\s*Community:\s*(.+)$", block, flags=re.MULTILINE)
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
# Junos output parsing
# NOTE: written from Junos config/show-output conventions, not yet validated
# against real captured output from a live Junos box (the IOS-XR parsers
# above went through exactly this and needed fixing once real output didn't
# match what was assumed - same risk applies here). Run with --sample-dir
# against captured real `show`/`show configuration` output before trusting
# this against live neighbors.
# ---------------------------------------------------------------------------

def parse_junos_bgp_summary(show_summary_text):
    """
    Parses `show bgp summary` -> (v4_ips, v6_ips).

    Junos mixes both address families in one table (unlike XR's separate
    `summary`/`summary` per AFI commands), e.g.:

        Groups: 2 Peers: 4 Down peers: 0
        Table          Tot Paths  Act Paths Suppressed    History Damp State    Pending
        inet.0                57         57          0          0          0          0
        Peer                     AS      InPkt     OutPkt    OutQ   Flaps Last Up/Dwn State|...
        24.93.64.1            11426      12345      12340       0       0     3w2d6h Establ
          inet.0: 57/57/57/0
        2606:a000:0:4::74     11426        123        120       0       0     3w2d6h Establ
          inet6.0: 1/1/1/0

    So instead of an XR-style header gate, every line's first token is
    tried as an IP address and split into v4/v6 by version.
    """
    v4_ips, v6_ips = [], []
    for line in show_summary_text.splitlines():
        parts = line.split()
        if not parts:
            continue
        try:
            ip = ipaddress.ip_address(parts[0])
        except ValueError:
            continue
        (v4_ips if ip.version == 4 else v6_ips).append(parts[0])
    return v4_ips, v6_ips


JUNOS_ROUTE_LINE_RE = re.compile(
    r"^\*?\s*(?P<prefix>\d+\.\d+\.\d+\.\d+/\d+|[0-9a-fA-F:]+/\d+)\s+(?P<nexthop>\S+)\s+(?P<rest>.+?)\s*$"
)


def parse_junos_routes(show_routes_text):
    """
    Parses `show route advertising-protocol bgp <ip>` / `show route
    receive-protocol bgp <ip>` output -> [{prefix, next_hop, as_path}].

    Typical Junos format:
        inet.0: 25 destinations, 25 routes (25 active, 0 holddown, 0 hidden)
          Prefix		  Nexthop	      MED     Lclpref    AS path
        * 10.1.194.0/23      Self                                    19115 11426 I
        * 22.80.64.0/19      24.93.64.1                               19115 11426 I

    Leading `*` marks the active route; nexthop is "Self" on advertised
    routes toward the peer. MED/Lclpref are frequently blank (not
    zero-padded) but not always - verified against real output where one
    row out of fifteen had a MED of "0", which leaked into as_path when this
    just kept everything after nexthop verbatim. Column position isn't
    reliable either (MED/Lclpref being blank shifts everything left), so
    instead: split the tail into tokens and take the trailing "<asn>
    <origin-code>" (or just "<asn>" if no origin code is present on that
    line) as the AS path, discarding any MED/Lclpref tokens in between.
    """
    routes = []
    for line in show_routes_text.splitlines():
        m = JUNOS_ROUTE_LINE_RE.match(line)
        if not m:
            continue
        rest_tokens = m.group("rest").split()
        if len(rest_tokens) >= 2 and rest_tokens[-1] in ("I", "E", "?", "i", "e"):
            as_path = " ".join(rest_tokens[-2:])
        elif rest_tokens:
            as_path = rest_tokens[-1]
        else:
            as_path = ""
        routes.append({"prefix": m.group("prefix"), "next_hop": m.group("nexthop"), "as_path": as_path})
    return routes


def parse_junos_bgp_config(config_text):
    """
    Parses `show configuration protocols bgp` (Junos curly-brace format)
    into {neighbor_ip: {neighbor_group, remote_as, description, policy_in,
    policy_out}}.

    Junos nests neighbors directly inside their group block (unlike XR's
    `use neighbor-group` reference), and a neighbor sub-block overrides
    whatever it repeats from the group:

        group IPV4-CORE {
            type external;
            description "CHRCNCTR01R";
            peer-as 11426;
            import IPV4-CORE-PEER-IN;
            export IPV4-CORE-PEER-OUT;
            neighbor 24.93.64.1;
            neighbor 24.93.64.2 {
                description "override";
            }
        }

    Walks brace depth generically (any "... {" pushes a context inherited
    from its parent, any "}" pops it) so unrelated nested stanzas (family,
    bfd-liveness-detection, etc.) are skipped harmlessly rather than
    breaking the group/neighbor tracking.
    """
    base_fields = {"neighbor_group": "", "remote_as": "", "description": "", "policy_in": "", "policy_out": ""}
    stack = [dict(base_fields)]
    tag_stack = [None]
    neighbors = {}

    for raw_line in config_text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        m = re.match(r"^neighbor\s+(\S+)\s*;$", line)
        if m:
            neighbors[m.group(1)] = dict(stack[-1])
            continue

        m = re.match(r"^group\s+(\S+)\s*\{$", line)
        if m:
            fields = dict(stack[-1])
            fields["neighbor_group"] = m.group(1)
            stack.append(fields)
            tag_stack.append(None)
            continue

        m = re.match(r"^neighbor\s+(\S+)\s*\{$", line)
        if m:
            stack.append(dict(stack[-1]))
            tag_stack.append(("neighbor", m.group(1)))
            continue

        if line.endswith("{"):
            stack.append(dict(stack[-1]))
            tag_stack.append(None)
            continue

        if line == "}":
            fields = stack.pop() if len(stack) > 1 else stack[-1]
            tag = tag_stack.pop() if len(tag_stack) > 1 else None
            if tag and tag[0] == "neighbor":
                neighbors[tag[1]] = fields
            continue

        m = re.match(r'^description\s+"([^"]*)"\s*;$', line) or re.match(r"^description\s+(\S+)\s*;$", line)
        if m:
            stack[-1]["description"] = m.group(1).strip()
            continue
        m = re.match(r"^peer-as\s+(\S+)\s*;$", line)
        if m:
            stack[-1]["remote_as"] = m.group(1)
            continue
        m = re.match(r"^import\s+(\S+)\s*;$", line)
        if m:
            stack[-1]["policy_in"] = m.group(1)
            continue
        m = re.match(r"^export\s+(\S+)\s*;$", line)
        if m:
            stack[-1]["policy_out"] = m.group(1)
            continue

    return neighbors


def parse_junos_policy_communities(policy_options_text):
    """
    Static fallback: resolves `then community add NAME;` inside
    `policy-options policy-statement ...` blocks against `policy-options
    community NAME members [ ... ];` definitions. Equivalent in spirit to
    the XR `route-policy ... set community (...)` static parser.

    Expects `show configuration policy-options` text. Same caveat as the
    XR version: a static read of the policy, not per-prefix - if a policy
    sets different communities per term, all are attributed to every
    neighbor using this policy. Brace depth is tracked with a running
    counter (rather than a full stack) since only "am I still inside this
    policy-statement" matters here.
    """
    community_defs = {}
    for line in policy_options_text.splitlines():
        m = re.match(r"^\s*community\s+(\S+)\s+members\s+\[([^\]]+)\]\s*;", line)
        if m:
            community_defs[m.group(1)] = m.group(2).split()
            continue
        m = re.match(r"^\s*community\s+(\S+)\s+members\s+(\S+)\s*;", line)
        if m:
            community_defs.setdefault(m.group(1), [m.group(2)])

    policies = {}
    current_policy = None
    depth = 0
    depth_at_policy_start = 0
    for raw_line in policy_options_text.splitlines():
        line = raw_line.strip()
        m = re.match(r"^policy-statement\s+(\S+)\s*\{$", line)
        if m and current_policy is None:
            current_policy = m.group(1)
            policies.setdefault(current_policy, [])
            depth_at_policy_start = depth
        if current_policy:
            cm = re.search(r"community\s+add\s+(\S+)\s*;", line)
            if cm:
                policies[current_policy].extend(community_defs.get(cm.group(1), []))
        depth += line.count("{") - line.count("}")
        if current_policy is not None and depth <= depth_at_policy_start:
            current_policy = None
    return policies


def parse_junos_prefix_communities(detail_text):
    """
    Parses `show route <prefix> extensive` -> best-path community string,
    semicolon-joined to match the sheet's existing format (same shape as
    parse_prefix_detail_communities for XR).

    Junos marks the active path's protocol line with a leading `*`
    (e.g. "*BGP    Preference: 170/-101") and lists communities as:
        Communities: 20115:3101 20115:64080
    Falls back to the first path with a Communities line if none is
    explicitly marked active.
    """
    best_communities = None
    first_communities = None
    current_is_best = False
    for line in detail_text.splitlines():
        if re.match(r"^\s*\*?(BGP|Local|Static|OSPF|IS-IS)\b", line):
            current_is_best = line.strip().startswith("*")
        m = re.search(r"Communities:\s*(.+)$", line)
        if m:
            communities = m.group(1).strip()
            if first_communities is None:
                first_communities = communities
            if current_is_best:
                best_communities = communities
    chosen = best_communities if best_communities is not None else (first_communities or "")
    return chosen.replace(" ", ";")


# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------

def collect_from_device(conn_or_dir, hostname, platform="iosxr", use_sample_dir=False):
    outputs = {}
    if platform == "junos":
        files = [("summary", "summary.txt"),
                 ("config_bgp", "config_bgp.txt"),
                 ("policy_options", "policy_options.txt")]
    else:
        files = [("summary_v4", "summary_v4.txt"),
                 ("summary_v6", "summary_v6.txt"),
                 ("running_bgp", "running_bgp.txt")]

    if use_sample_dir:
        base = Path(conn_or_dir) / hostname
        for key, fname in files:
            fp = base / fname
            outputs[key] = fp.read_text() if fp.exists() else ""
    else:
        conn = conn_or_dir
        if platform == "junos":
            cmds = [("summary", "show bgp summary"),
                    ("config_bgp", "show configuration protocols bgp"),
                    ("policy_options", "show configuration policy-options")]
        else:
            cmds = [("summary_v4", "show bgp ipv4 unicast summary"),
                    ("summary_v6", "show bgp ipv6 unicast summary"),
                    ("running_bgp", "show running-config router bgp")]
        for key, cmd in cmds:
            try:
                # A full config/policy dump on a core router with many
                # neighbor-groups (seen: 20+ on rcr01drhmncev) can be large -
                # same slow-command risk the route pulls already needed
                # protection from, so this gets the same delay_factor bump
                # and recovery-on-failure instead of crashing the whole
                # device and losing everything collected from it.
                outputs[key] = _run(conn, cmd, delay_factor=4)
            except Exception as e:
                print(f"    ! `{cmd}` failed/timed out ({e}) - treating as empty, "
                      f"recovering the connection, and continuing. Neighbor "
                      f"discovery/config context for this device may be incomplete.",
                      file=sys.stderr)
                outputs[key] = ""
                _recover_connection(conn)
    return outputs


def _run(conn, cmd, delay_factor=1):
    """Runs a show command on the device, echoing it to stderr first so you
    can see exactly what's about to execute (and Ctrl+C before it if needed).
    Every command this script ever sends is a read-only `show` - safe to
    interrupt at any point, nothing is left half-configured.

    delay_factor scales netmiko's read timeout for slow commands (e.g. a
    full route-table dump on a large peer) - default 1 is netmiko's normal
    ~10s-ish budget, bump it for advertised/received-routes pulls since
    those are the commands most likely to be slow on a busy neighbor."""
    print(f"    $ {cmd}", file=sys.stderr)
    return conn.send_command(cmd, delay_factor=delay_factor)


def _recover_connection(conn):
    """
    Best-effort recovery after a command that never returned its expected
    prompt (typically a huge route dump on an unrelated large peer, e.g. a
    full-table transit session). netmiko's internal prompt-tracking can be
    left pointed at leftover buffered route data instead of the real
    device prompt after a stall like that - which then makes *every
    subsequent* command on the same SSH session fail too, even small ones
    on completely unrelated neighbors. Drop and re-establish the session
    rather than keep reusing a connection that may be in that state.
    """
    try:
        conn.disconnect()
    except Exception:
        pass
    try:
        conn.establish_connection()
        conn.session_preparation()
    except Exception as e:
        print(f"    ! could not recover connection after a stalled command ({e}) - "
              f"remaining pulls on this device will likely keep failing.", file=sys.stderr)


def collect_routes_for_neighbor(conn_or_dir, hostname, ip, afi, platform="iosxr", use_sample_dir=False):
    if use_sample_dir:
        base = Path(conn_or_dir) / hostname
        adv_file = base / f"advertised_{ip.replace(':', '_')}.txt"
        rec_file = base / f"received_{ip.replace(':', '_')}.txt"
        adv = adv_file.read_text() if adv_file.exists() else ""
        rec = rec_file.read_text() if rec_file.exists() else ""
    else:
        conn = conn_or_dir
        if platform == "junos":
            table_suffix = " table inet6.0" if afi == "ipv6" else ""
            try:
                adv = _run(conn, f"show route advertising-protocol bgp {ip}{table_suffix}", delay_factor=4)
            except Exception as e:
                print(f"    ! advertised-routes pull for {ip} failed/timed out ({e}) - "
                      f"treating as empty, recovering the connection, and continuing. Likely "
                      f"a large table on this peer; re-run scoped to just the neighbor-group "
                      f"you actually need (--neighbor-groups) to skip it entirely.",
                      file=sys.stderr)
                adv = ""
                _recover_connection(conn)
            try:
                rec = _run(conn, f"show route receive-protocol bgp {ip}{table_suffix}", delay_factor=4)
            except Exception as e:
                print(f"    ! received-routes pull for {ip} failed/timed out ({e}) - "
                      f"treating as empty, recovering the connection, and continuing.", file=sys.stderr)
                rec = ""
                _recover_connection(conn)
        else:
            adv = _run(conn, f"show bgp {afi} unicast neighbors {ip} advertised-routes")
            try:
                # "received routes" (two words) - confirmed against real
                # IOS-XR CLI, not hyphenated like "advertised-routes" is.
                rec = _run(conn, f"show bgp {afi} unicast neighbors {ip} received routes")
                first_line = rec.splitlines()[0] if rec.strip() else ""
                if not rec.strip() or first_line.startswith("%"):
                    print(f"    ! received-routes for {ip} came back empty/errored - if this "
                          f"neighbor should be receiving prefixes, check "
                          f"`soft-reconfiguration inbound` is set for its group (needed for "
                          f"IOS-XR to show the pre-policy Adj-RIB-In).", file=sys.stderr)
            except Exception:
                print(f"    ! received-routes pull for {ip} failed - needs "
                      f"`soft-reconfiguration inbound` configured on this neighbor/group, or "
                      f"route-refresh capability. Treating as empty.", file=sys.stderr)
                rec = ""
    return adv, rec


def collect_prefix_communities(conn_or_dir, hostname, prefix, afi, platform="iosxr", use_sample_dir=False):
    """
    Per-prefix community lookup. Only called when --with-communities is
    set, since it's one extra command per prefix (can be hundreds across a
    full neighbor table - slow but exact).
    """
    if use_sample_dir:
        base = Path(conn_or_dir) / hostname
        fname = f"prefix_detail_{prefix.replace('/', '_').replace(':', '_')}.txt"
        fp = base / fname
        text = fp.read_text() if fp.exists() else ""
    else:
        conn = conn_or_dir
        cmd = (f"show route {prefix} extensive" if platform == "junos"
               else f"show bgp {afi} unicast {prefix}")
        try:
            # A prefix like 0.0.0.0/0 can have a huge extensive/detail
            # output (many candidate paths) - same slow-peer risk as the
            # main route pulls, so it gets the same delay_factor bump and
            # recovery-on-failure instead of taking down the whole device.
            text = _run(conn, cmd, delay_factor=4)
        except Exception as e:
            print(f"    ! community pull for {prefix} failed/timed out ({e}) - "
                  f"treating as no communities found, recovering the connection, "
                  f"and continuing.", file=sys.stderr)
            _recover_connection(conn)
            text = ""
    parser = parse_junos_prefix_communities if platform == "junos" else parse_prefix_detail_communities
    result = parser(text)
    if not result and text.strip():
        # Command succeeded and returned something, but the parser found no
        # communities in it. Print the FULL output (not a truncated snippet -
        # a 12-line cutoff previously cut this off right before the
        # "Community:" line, which typically appears well into each
        # "Path #N:" block, making a real parser bug look like "no data").
        # These outputs are only 20-30 lines, no reason to truncate.
        print(f"    ! no communities found in output for {prefix} (command succeeded, "
              f"{len(text.splitlines())} lines returned) - full raw output:\n"
              f"------\n{text}\n------", file=sys.stderr)
    return result


def build_rows(location, hostname, neighbor_cfg, communities_by_policy, aggregates,
               summary_v4_ips, summary_v6_ips, conn_or_dir, use_sample_dir, with_communities,
               neighbor_group_filter=None, platform="iosxr", neighbor_ip_filter=None):
    rows = []
    all_ips = [(ip, "ipv4") for ip in summary_v4_ips] + [(ip, "ipv6") for ip in summary_v6_ips]
    route_parser = parse_junos_routes if platform == "junos" else parse_routes

    for ip, afi in all_ips:
        cfg = neighbor_cfg.get(ip, {})
        group = cfg.get("neighbor_group") or ""

        if neighbor_ip_filter and ip not in neighbor_ip_filter:
            print(f"  neighbor {ip} ({afi}, group={group or 'unknown'}): skipped, "
                  f"not in --neighbor-ips filter", file=sys.stderr)
            continue

        if neighbor_group_filter and group.upper() not in neighbor_group_filter:
            print(f"  neighbor {ip} ({afi}, group={group or 'unknown'}): skipped, "
                  f"not in --neighbor-groups filter", file=sys.stderr)
            continue

        print(f"  neighbor {ip} ({afi}, group={group or 'unknown'}):", file=sys.stderr)
        adv_text, rec_text = collect_routes_for_neighbor(conn_or_dir, hostname, ip, afi, platform, use_sample_dir)
        adv_routes = route_parser(adv_text)
        rec_routes = route_parser(rec_text)

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

        if with_communities and adv_routes:
            print(f"    {len(adv_routes)} advertised prefixes - pulling per-prefix community "
                  f"for each (one command per prefix)...", file=sys.stderr)

        for i, r in enumerate(adv_routes, 1):
            if with_communities:
                print(f"    [{i}/{len(adv_routes)}] {r['prefix']}", file=sys.stderr)
                communities = collect_prefix_communities(conn_or_dir, hostname, r["prefix"], afi, platform, use_sample_dir)
            else:
                communities = static_comm_out
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
    ap.add_argument("--devices", default="devices.csv",
                    help="CSV: location,hostname,mgmt_ip,device_type,port. device_type is a "
                         "netmiko value - cisco_xr/cisco_xe for IOS-XR/XE, or juniper_junos "
                         "for Junos (full neighbor discovery supported on both platforms). "
                         "Use --neighbor-ips to scope a run to specific neighbors.")
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
    ap.add_argument("--discover-aggregates", metavar="OUT_CSV", default=None,
                    help="Instead of the full prework pull, just parse each device's "
                         "running-config for `network`/`aggregate-address` statements "
                         "and write a candidate aggregates CSV to OUT_CSV for you to "
                         "review. Classification is left as TBD - that's your/NOC's "
                         "call, not something derivable from the config.")
    ap.add_argument("--neighbor-groups", default=None,
                    help="Comma-separated neighbor-group names to include (case-insensitive), "
                         "e.g. IPV4_CORE,IPV6_CORE. Neighbors on any other group (ELF, SPINE, "
                         "etc.) are skipped entirely - useful when the prework is only about "
                         "the groups you're actually renaming. Default: no filter, all neighbors. "
                         "A neighbor-group can contain multiple neighbor IPs sharing the same "
                         "policy - use --neighbor-ips instead if you need exactly one IP.")
    ap.add_argument("--neighbor-ips", default=None,
                    help="Comma-separated exact neighbor IPs to include, e.g. "
                         "24.93.64.1,24.93.64.3 - skips every other neighbor on the device "
                         "regardless of what neighbor-group it's in. No space after commas. "
                         "Combine with --neighbor-groups if you want both narrowed further "
                         "(a neighbor must pass both filters to be included).")
    args = ap.parse_args()

    neighbor_group_filter = (
        {g.strip().upper() for g in args.neighbor_groups.split(",")}
        if args.neighbor_groups else None
    )
    neighbor_ip_filter = (
        {ip.strip() for ip in args.neighbor_ips.split(",")}
        if args.neighbor_ips else None
    )

    if args.discover_aggregates:
        devices = load_devices(args.devices)
        use_sample_dir = args.sample_dir is not None
        username = password = None
        if not use_sample_dir:
            if ConnectHandler is None:
                sys.exit("netmiko is not installed. Run: pip install -r requirements.txt")
            username = os.environ.get("NETOPS_USERNAME") or input("Username: ")
            password = os.environ.get("NETOPS_PASSWORD") or getpass.getpass("Password: ")

        discovered = []
        for dev in devices:
            location, hostname = dev["location"], dev["hostname"]
            platform = device_platform(dev)
            if platform == "junos":
                print(f"[{location}] {hostname}: --discover-aggregates only supports IOS-XR "
                      f"`network`/`aggregate-address` parsing today - skipping Junos device "
                      f"(would need `show configuration policy-options` + routing-options "
                      f"parsing). Fill its aggregates.csv rows in by hand.", file=sys.stderr)
                continue
            print(f"[{location}] {hostname}: reading running-config...", file=sys.stderr)
            if use_sample_dir:
                running_bgp = (Path(args.sample_dir) / hostname / "running_bgp.txt").read_text()
            else:
                conn = ConnectHandler(
                    device_type=dev.get("device_type", "cisco_xr"),
                    host=dev["mgmt_ip"], username=username, password=password,
                    port=int(dev.get("port") or 22),
                )
                try:
                    running_bgp = _run(conn, "show running-config router bgp", delay_factor=4)
                except Exception as e:
                    print(f"    ! running-config pull failed/timed out ({e}) - "
                          f"skipping this device's aggregates.", file=sys.stderr)
                    running_bgp = ""
                conn.disconnect()
            for o in parse_bgp_originations(running_bgp):
                discovered.append({
                    "location": location, "aggregate_cidr": o["cidr"],
                    "classification": "TBD",
                    "description": f"{o['kind']}" + (" summary-only" if o["summary_only"] else "")
                                   + (f" via {o['route_policy']}" if o["route_policy"] else "")
                                   + f" (from {hostname} running-config)",
                })

        with open(args.discover_aggregates, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["location", "aggregate_cidr", "classification", "description"])
            writer.writeheader()
            writer.writerows(discovered)
        print(f"Wrote {len(discovered)} candidate aggregate rows to {args.discover_aggregates}. "
              f"Review/edit the 'classification' column (TBD -> Public or Private GUA) before using it.",
              file=sys.stderr)
        return

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
    failed_devices = []

    for dev in devices:
        location, hostname = dev["location"], dev["hostname"]
        platform = device_platform(dev)
        print(f"[{location}] {hostname}: collecting ({platform})...", file=sys.stderr)

        conn_or_dir = None
        try:
            if use_sample_dir:
                conn_or_dir = args.sample_dir
                outputs = collect_from_device(conn_or_dir, hostname, platform, use_sample_dir=True)
            else:
                conn_or_dir = ConnectHandler(
                    device_type=dev.get("device_type", "cisco_xr"),
                    host=dev["mgmt_ip"], username=username, password=password,
                    port=int(dev.get("port") or 22),
                    conn_timeout=20, banner_timeout=20, auth_timeout=20, timeout=20,
                )
                outputs = collect_from_device(conn_or_dir, hostname, platform, use_sample_dir=False)

            if platform == "junos":
                neighbor_cfg = parse_junos_bgp_config(outputs["config_bgp"])
                communities_by_policy = parse_junos_policy_communities(outputs["policy_options"])
                summary_v4_ips, summary_v6_ips = parse_junos_bgp_summary(outputs["summary"])
            else:
                neighbor_cfg = parse_bgp_running_config(outputs["running_bgp"])
                communities_by_policy = parse_route_policy_communities(outputs["running_bgp"])
                summary_v4_ips = parse_bgp_summary(outputs["summary_v4"])
                summary_v6_ips = parse_bgp_summary(outputs["summary_v6"])

            rows = build_rows(location, hostname, neighbor_cfg, communities_by_policy, aggregates,
                              summary_v4_ips, summary_v6_ips, conn_or_dir, use_sample_dir,
                              args.with_communities, neighbor_group_filter, platform, neighbor_ip_filter)
            all_rows.extend(rows)
        except Exception as e:
            print(f"[{location}] {hostname}: FAILED ({e}) - skipping this device, keeping "
                  f"rows already collected from earlier devices.", file=sys.stderr)
            failed_devices.append(hostname)
        finally:
            if conn_or_dir is not None and not use_sample_dir:
                try:
                    conn_or_dir.disconnect()
                except Exception:
                    pass

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
    if failed_devices:
        print(f"WARNING: {len(failed_devices)} device(s) failed and were skipped "
              f"(no rows collected from them): {', '.join(failed_devices)}", file=sys.stderr)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        # Every command this script runs is a read-only `show` - aborting
        # mid-run leaves nothing half-configured on any device, and no
        # output file is written for a run you interrupted.
        print("\nAborted (Ctrl+C) - no output written. Nothing was changed "
              "on any device (only read-only `show` commands are ever sent).",
              file=sys.stderr)
        sys.exit(130)
