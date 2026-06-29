"""Compile EVERY network host this run will touch, so the user can allowlist them before
any data pull runs (sandboxed environments block egress to un-allowlisted domains, which is
the usual reason a run "doesn't work"). Reads the authored config.json and prints + writes
the domain list grouped by purpose. Run this right after authoring config.json.

Outputs (in the run dir):
  in/required_domains.txt   — one host per line (paste-ready for an allowlist)
  in/required_endpoints.json — full URLs + purpose (audit trail / Data Sources)
"""
import json
import sys
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402

# Fixed infrastructure the pipeline always uses.
GEOCODERS = [
    ("geocoding.geo.census.gov", "Subject geocoding (US Census, no key)"),
    ("geocode.arcgis.com", "Subject geocoding fallback (ArcGIS World Geocoder)"),
]
VIEWER = [
    ("unpkg.com", "Viewer: Leaflet + leaflet-image JS/CSS (loads when the HTML opens)"),
    ("server.arcgisonline.com", "Viewer: Esri satellite + street basemap tiles"),
]


def host(url):
    return urlparse(url).netloc.lower()


def wildcard(h):
    """Collapse a subdomained host to '*.<registrable>' so one allowlist entry covers
    sibling subdomains (e.g. servicesN.arcgis.com / geocode.arcgis.com -> *.arcgis.com).
    Two-label hosts (unpkg.com) are returned as-is. Heuristic: last two labels = the
    registrable domain (true for the .com/.gov/.org/.net/.us hosts these GIS layers use)."""
    parts = h.split(".")
    if len(parts) > 2:
        return "*." + ".".join(parts[-2:])
    return h


def main():
    cfg = C.load()
    IN = C.indir(cfg)

    pull, seen = [], set()
    for s in cfg.get("parcel_sources", []):
        h = host(s["url"])
        pull.append((h, f"Parcels - {s.get('name') or s.get('county', '')}", s["url"]))
    for s in cfg.get("zoning_sources", []):
        h = host(s["url"])
        pull.append((h, f"Zoning - {s['jurisdiction']}", s["url"]))

    groups = {
        "BUILD-TIME - Python data pulls (MUST be allowlisted in the run sandbox)":
            [(h, why, url) for h, why, url in pull],
        "SUBJECT GEOCODING - only if you resolve the point via Python (else N/A)":
            [(h, why, None) for h, why in GEOCODERS],
        "VIEW-TIME - loaded by the HTML viewer in the browser":
            [(h, why, None) for h, why in VIEWER],
    }

    all_hosts, lines, endpoints = [], [], []
    print("\n================ DOMAINS THIS RUN NEEDS ================")
    for title, items in groups.items():
        print(f"\n## {title}")
        for h, why, url in items:
            mark = "" if h in seen else " *new*"
            seen.add(h)
            all_hosts.append(h)
            print(f"  - {h:38} {why}{mark}")
            if url:
                print(f"      {url}")
            endpoints.append({"host": h, "purpose": why, "url": url, "group": title})
    uniq = sorted(set(all_hosts))
    wild = sorted(set(wildcard(h) for h in uniq))
    (IN / "required_domains.txt").write_text("\n".join(uniq) + "\n", encoding="utf-8")
    (IN / "required_domains_wildcard.txt").write_text("\n".join(wild) + "\n", encoding="utf-8")
    (IN / "required_endpoints.json").write_text(json.dumps(endpoints, indent=2), encoding="utf-8")

    print("\n========= ALLOWLIST (WILDCARD - RECOMMENDED, fewer entries) =========")
    print("  Covers rotating ArcGIS subdomains (servicesN.arcgis.com) so other cities'")
    print("  zoning layers work without re-editing the allowlist:")
    for h in wild:
        print(f"  {h}")
    print("\n========= ALLOWLIST (EXACT HOSTS - if you prefer the narrowest set) =========")
    for h in uniq:
        print(f"  {h}")

    # Next-session prompt — allowlist changes only take effect in a NEW session, so hand the
    # user a ready-to-paste prompt for a fresh chat that rebuilds this run deterministically.
    sub = cfg.get("subject", {})
    who = sub.get("name", "the subject")
    addr = f", {sub['address']}" if sub.get("address") else ""
    domains = ", ".join(wild)
    prompt = (
        f"/land-use-analysis for {who}{addr}\n\n"
        f"Network: egress is already ON with these domains allowlisted — {domains}.\n"
        f"(Allowlist changes only apply to a NEW session, which this is.)\n\n"
        f"Before pulling, run the reachability pre-flight on every parcel/zoning endpoint. Any host\n"
        f"that returns 503/upstream-timeout (NOT host_not_allowed) is allowed-but-unreachable —\n"
        f"substitute a reachable county/state layer (prefer *.arcgis.com / county GIS over a city's\n"
        f"self-hosted ArcGIS box), log the swap, then proceed pull -> classify -> reason -> deliver.\n")
    (IN / "next_session_prompt.txt").write_text(prompt, encoding="utf-8")

    print(f"\nwrote in/required_domains.txt ({len(uniq)}) + in/required_domains_wildcard.txt "
          f"({len(wild)}) + in/required_endpoints.json + in/next_session_prompt.txt")
    print("\nHAND THE USER TWO THINGS:")
    print(" 1) The domain list above — in claude.ai: Settings > Capabilities > 'Code execution and")
    print("    file creation' > turn 'Allow network egress' ON, then under 'Additional allowed domains'")
    print("    Add each entry (wildcards like *.arcgis.com supported).")
    print(" 2) A NEXT-SESSION PROMPT (in/next_session_prompt.txt) — the allowlist only takes effect in")
    print("    a session started AFTER it's added, so they must paste this into a FRESH chat:")
    print("\n----------------------------- paste into a new chat -----------------------------")
    print(prompt.rstrip())
    print("---------------------------------------------------------------------------------")
    print("\n(If egress is already live in THIS session: echo ok > in/allowlist_ok && run_all.py pull)")


if __name__ == "__main__":
    main()
