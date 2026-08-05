# BGP prework: neighbor-group standardization

Prework tool for the remaining 2 locations before renaming neighbor groups
(e.g. `IPV4_CORE`/`IPV6_CORE` -> your new standard naming). It captures a
snapshot of current neighbor config, advertised/received prefixes, and
communities, so you can diff before-vs-after and confirm the NOC/outbound
team's route manipulation still works once the rename lands.

Validated against real ASR9K/IOS-XR output (running-config, `summary`,
`advertised-routes`, and per-prefix `show bgp ... <prefix>` detail) — parser
format assumptions were corrected against actual command output rather than
guessed.

## Files

- `devices.csv` — fill in the 2 remaining locations: hostname, mgmt IP,
  `device_type` (netmiko value — `cisco_xr` or `cisco_xe`), SSH port.
- `aggregates.csv` — fill in the CIDR block(s) your team owns per location,
  plus a `classification` per block: `Public` (freely advertised) or
  `Private GUA` (real public-IP space you own but the NOC team wants kept off
  the internet). This has to come from your team's allocation — it's a
  business decision, not something derivable from the router or a WHOIS
  lookup. RFC1918/CGNAT/ULA/link-local ranges are auto-detected, no input
  needed for those.
- `bgp_prework.py` — the collector script.
- `requirements.txt` — Python deps.

## Setup

```bash
cd ~/bgp-prework
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
```

## Recommended: validate the parser before touching live routers

Before running against production devices, capture real output once
(manually, over an existing SSH session) and test the parser against text
files instead of live SSH:

```
samples/
  <hostname>/
    summary_v4.txt          # show bgp ipv4 unicast summary
    summary_v6.txt          # show bgp ipv6 unicast summary
    running_bgp.txt         # show running-config router bgp
    advertised_<ip>.txt     # show bgp ipv4 unicast neighbors <ip> advertised-routes
    received_<ip>.txt       # show bgp ipv4 unicast neighbors <ip> received-routes
    prefix_detail_<prefix_with_slash_as_underscore>.txt
                            # show bgp ipv4 unicast <prefix>  (only needed with --with-communities)
```

```bash
python bgp_prework.py --sample-dir ./samples --devices devices.csv --aggregates aggregates.csv --out prework_test.xlsx
```

Open `prework_test.xlsx` and check the columns look right. If something's
off, paste me the actual `show` output that didn't parse and I'll fix the
regex — no point guessing at IOS-XR quirks blind (this already happened once:
the original `advertised-routes` parser assumed a classic-IOS metric/locprf
table, but real ASR9K/XR output is `Network / Next Hop / From / AS Path` —
fixed and re-verified against real output).

## Running against live devices

```bash
export NETOPS_USERNAME=your_username
python bgp_prework.py --devices devices.csv --aggregates aggregates.csv --out prework.xlsx
```

You'll be prompted for the password via `getpass` (not echoed, never
written to disk). Alternatively set `NETOPS_PASSWORD` as an env var for a
single run — don't put it in a file that gets committed anywhere.

Add `--with-communities` to pull the *exact* best-path community per
advertised prefix via `show bgp ... <prefix>` (one extra SSH command per
prefix — can be slow on a neighbor with hundreds of routes). Without it, the
`communities` column falls back to the static list of communities the
route-policy *can* set, not necessarily the one that applied to that specific
prefix.

## Output columns

`prework.xlsx`, sheet `bgp_prework` — one row per (neighbor, direction, prefix):

| column | meaning |
|---|---|
| `location`, `hostname` | which device/site this row is from |
| `neighbor` | neighbor IP |
| `prefix`, `nexthop`, `as_path` | from advertised/received-routes |
| `communities` | best-path community (exact, with `--with-communities`) or static route-policy list |
| `COMMENT` | outbound route-policy name applied (starting point — not a substitute for manual annotation referencing specific prefix-lists/other devices) |
| `IPv4/IPv6`, `Agg`, `IP Classification` | address family, matched aggregate, and classification (`RFC 1918` / `CGNAT (RFC 6598)` / `Private GUA` / `Public`, etc.) |
| `direction` | `advertised (out)` or `received (in)` |
| `neighbor_group`, `remote_as`, `description`, `policy_out`, `policy_in` | config context, useful for the before/after diff once neighbor-groups are renamed |

A second sheet, `neighbor_group_summary`, gives route counts per
neighbor-group/neighbor — a quick before/after sanity check once you rename
the groups.

## Known limitations

- **`COMMENT` isn't the full manual annotation your example sheet had**
  (e.g. "@dlr09spbgscfe/dlr10spbgscfe"). That level of detail references
  other devices and isn't derivable from a single box's config — it's
  populated with the outbound route-policy name as a starting point only.
- **`received-routes` needs `soft-reconfiguration inbound`** (or route-refresh
  capability) configured on the neighbor, or IOS-XR won't have the pre-policy
  Adj-RIB-In to show. If it comes back empty, that's likely why.
- Regex-based config/output parsing — validated against real ASR9K/IOS-XR
  6.6.2 output, but other releases/platforms may format slightly differently.
  Use the sample-dir workflow above to catch that before it matters.
