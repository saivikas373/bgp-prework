# BGP prework: neighbor-group standardization

Prework tool for the remaining 2 locations before renaming neighbor groups
(e.g. `IPV4_CORE`/`IPV6_CORE` -> your new standard naming). It captures a
snapshot of current neighbor config, advertised/received prefixes, and
communities, so you can diff before-vs-after and confirm the NOC/outbound
team's route manipulation still works once the rename lands.

Validated against real ASR9K/IOS-XR output (running-config, `summary`,
`advertised-routes`, and per-prefix `show bgp ... <prefix>` detail) — parser
format assumptions were corrected against actual command output rather than
guessed. Also supports Junos (`device_type=juniper_junos`) — see
"Multi-vendor support" below for exactly which Junos parsers are validated
against real output vs. synthetic-only so far.

## Files

- `devices.csv` — fill in the locations: hostname, mgmt IP, `device_type`
  (netmiko value — `cisco_xr`/`cisco_xe` for Cisco, `juniper_junos` for
  Junos), SSH port. Full neighbor discovery works on both platforms — no
  need to hand-specify a neighbor IP. Use `--neighbor-ips`/`--neighbor-groups`
  (see below) to scope a run to specific neighbors.
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

Add `--neighbor-groups IPV4_CORE,IPV6_CORE` (comma-separated, case-insensitive,
**no space after the comma** — the shell splits on it otherwise) to only
process neighbors on those specific neighbor-groups, skipping everything else
(ELF, SPINE, etc. entirely, not just filtering them out afterward). Use this
when the prework is scoped to the groups you're actually renaming.

Prefer `--neighbor-ips 24.93.64.1,24.93.64.3` (also comma-separated, no space)
when you want to scope to *exact* neighbor addresses instead — a real
neighbor-group was found bundling multiple unrelated neighbors under one
shared policy, so `--neighbor-groups` alone doesn't guarantee you're only
looking at the neighbor you actually meant. Both filters can be combined.

Add `--discover-aggregates discovered.csv` to skip the full prework pull and
just parse each device's running-config for `network`/`aggregate-address`
statements, writing a candidate `aggregates.csv` for you to review — saves
hand-transcribing CIDR blocks out of the config. The `classification` column
comes back `TBD` for every row; fill in `Public`/`Private GUA` yourself, that
part is NOC's call.

Every command the script ever sends is a read-only `show` (never
`configure terminal`, never `commit`) — safe to Ctrl+C at any point, nothing
is left half-done on the device. The script prints each command right before
it runs, so you can watch exactly what's happening live.

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

## Cross-checking advertised vs. received (compare_advertised_received.py)

Running `bgp_prework.py` on the ELF side gives you what it *advertises*;
running it on the upstream/RCR side gives you what that neighbor *receives*.
Those two should match exactly - any gap means something's silently dropping
or altering routes between the two, which is exactly what you want to catch
both before and after the neighbor-group rename.

```bash
cp pairs.example.csv pairs.csv   # edit with your real hostnames/neighbor IPs
python compare_advertised_received.py --pre bgppre.xlsx --rcr bgprcr.xlsx --pairs pairs.csv --out reconciliation.xlsx
```

- `--pre` — workbook with the ELF-side `advertised (out)` rows
- `--rcr` — workbook with the upstream-side `received (in)` rows (can be the
  same file as `--pre` if one combined run covered both directions)
- `pairs.csv` — maps which ELF hostname+neighbor corresponds to which
  RCR-side hostname+neighbor (a link has two ends; the script needs to know
  which two device/neighbor combos represent the same physical session)

Output (`reconciliation.xlsx`) has three sheets: `summary` (counts per pair),
`detail` (every prefix, status = matched/missing/unexpected), and
`mismatches_only` (just the ones worth investigating). `MISSING` means
advertised but not confirmed received on the other end; `UNEXPECTED` means
received but not in the advertised list — check that side's `advertised-routes`
pull, since it can also just mean two runs weren't from the same point in time.

Run this once now (baseline, before the rename) and again after standardizing
the neighbor-groups — if the mismatch counts are still zero both times, the
rename didn't change what's actually getting through.

## Multi-vendor support

The script was originally IOS-XR-only. Junos support (`device_type=
juniper_junos` in `devices.csv`) was added afterward for checking what an
*upstream* device actually receives from an ELF-side neighbor, as a
cross-check against what the ELF side shows itself advertising. It's now a
full parallel implementation, not just the receive-side: `show bgp summary`
(neighbor discovery), `show configuration protocols bgp` +
`show configuration policy-options` (neighbor-group/remote-AS/policy/
community context), and both `show route advertising-protocol bgp <ip>` /
`show route receive-protocol bgp <ip>` (route pulls) all have Junos parsers,
mirroring the IOS-XR code paths.

**Validation status matters here — don't treat all of this as equally
trustworthy yet:**
- The route-pull parser (`parse_junos_routes`) **is** verified against real
  device output (`rcr01chrcnctr-re0`, neighbor `24.93.64.1`) — and that
  verification caught a real bug (a MED value leaking into the `as_path`
  field on one row out of fifteen), now fixed.
- The summary/config/policy parsers (`parse_junos_bgp_summary`,
  `parse_junos_bgp_config`, `parse_junos_policy_communities`) were written
  from Junos conventions and validated only against synthetic sample-dir
  output, **not yet against real device output** — same situation the
  IOS-XR parsers were in before real output showed the `advertised-routes`
  format assumption was wrong. Run `--sample-dir` against real captured
  output for these commands before trusting a live run's `neighbor_group`/
  `remote_as`/`description`/`policy_out`/`policy_in`/`communities` columns
  on Junos rows.
- `--discover-aggregates` still only supports IOS-XR — Junos devices are
  skipped with a message telling you to fill in `aggregates.csv` by hand for
  those rows (would need `show configuration policy-options` +
  `routing-options` parsing to automate).

Other things the parallel Junos work found and fixed, worth knowing about
regardless of platform:
- **`--neighbor-ips <ip1>,<ip2>`** (exact BGP neighbor address match) is more
  precise than `--neighbor-groups` — a real neighbor-group was found to
  bundle multiple unrelated neighbor IPs under one shared policy, so group
  filtering alone doesn't guarantee you're only looking at the neighbor you
  meant.
- **Connection recovery**: one neighbor with a genuinely huge route table can
  stall badly enough that netmiko's prompt-tracking gets left pointed at
  leftover buffered data, silently breaking every subsequent command on that
  same SSH session. The script now detects a stalled route pull and drops/
  re-establishes the connection rather than letting one bad peer poison every
  result queued after it.
- **Per-device failure isolation**: if one device is unreachable or times out
  during connection setup, that device is skipped (and listed in a
  `failed_devices` summary at the end) instead of the whole run losing every
  row already collected from devices processed earlier.

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

## Session notes: environment gotchas already hit and fixed

This ran into a string of environment-specific problems on the actual
Charter jump server (`chrcnctr-aps-naas-jmp-01`) — recording them here so the
next session (possibly on a different machine, e.g. an office laptop) doesn't
rediscover them from scratch:

- **Access path**: jump server requires PID + password + a 30-second-expiry
  Symantec VIP code, then a second plain-password SSH hop to the actual
  device. This script is meant to run **from the jump server itself**, after
  you've already authenticated there manually — it only automates the second
  hop (jump server -> device), not the OTP step, and it shouldn't try to.
- **Home directory disk quota**: hit "Disk quota exceeded" cloning into `~`
  on one jump server instance. Fix was either cloning into `/tmp` instead, or
  (as happened here) just being on a different jump server host that had
  quota available. Check `quota -s` / try `/tmp` if `git clone` fails this
  way.
- **Python 3.6.8 + pip 9.0.3** on the jump server (EOL, 2018-era pip). This
  forced several fixes, already applied in this repo:
  - `requirements.txt` pinned to `netmiko==3.4.0`, `pandas==1.1.5`,
    `openpyxl==3.0.10` (newer majors need Python 3.8/3.9+).
  - `paramiko==2.11.0` + `bcrypt==3.2.2` pinned explicitly — netmiko's
    unbounded `paramiko>=2.6.0` otherwise resolves to a paramiko needing a
    bcrypt version that requires a Rust build toolchain, not available here.
  - `pip install --upgrade pip` inside the venv first — pip 9 doesn't
    understand modern wheel tags and falls back to building from source
    (which then fails without a compiler), even when a wheel exists.
  - Code had to avoid two Python 3.7+-only things: `re.split()` on a
    zero-width lookahead pattern, and `ipaddress.IPv4Network.subnet_of()`.
    Both fixed in the code itself (see `_is_subnet_of()` and the manual
    `Path #N:` block-splitting in `parse_prefix_detail_communities()`), so
    this shouldn't resurface — but worth knowing why those look the way they
    do if you're reading the source.
- **GitHub push auth**: password auth for git is disabled by GitHub; needed
  `gh auth setup-git` to wire up the credential helper from an already
  `gh auth login`-ed session.
- **Repo visibility**: kept **public** deliberately, since the tracked files
  never contain real device data (`devices.csv`/`aggregates.csv`/`samples/`/
  `*.xlsx` are all gitignored) — this let the jump server `git clone` without
  needing a personal access token sitting on a shared corporate machine. If
  you ever add real data to a tracked file, stop and reconsider this.
- **Getting the output file off the jump server**: unresolved as of the last
  session — a `scp` attempt hit `Not a directory` on the destination path.
  Likely `~/Downloads` doesn't exist on one end. Next attempt: `scp
  prework.xlsx <user>@<host>:~/` (home dir, not a subfolder) from whichever
  side actually has reachability, or check exactly which command was run.
- **Shell quoting**: `--neighbor-groups IPV4_CORE, IPV6_CORE` (space after
  comma) gets split into two shell arguments and argparse rejects the second
  one as unrecognized. No space after the comma.
