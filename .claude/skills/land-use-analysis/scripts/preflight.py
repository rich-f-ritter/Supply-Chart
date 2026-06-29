"""Reachability pre-flight — ALLOWLISTED != REACHABLE.

A host can pass the sandbox allowlist yet be unreachable from the sandbox's egress IPs
(you get HTTP 503 "upstream connect error ... connection timeout", or a hang — NOT
x-deny-reason=host_not_allowed). City-self-hosted ArcGIS servers are the most common
offender. This probes every parcel/zoning endpoint root once and classifies it, so the
pull doesn't silently fail. Run before pulling (run_all.py does this automatically).

Verdicts:
  reachable   HTTP 200 — good.
  responded   up but the layer root returned HTTP N!=200 (often a slightly-off URL) — check.
  blocked     x-deny-reason=host_not_allowed — ALLOWLIST/SESSION problem (allowlist not in
              effect; remember it only applies to a NEW session started after you added it).
  unreachable allowed but origin down: 503/502/504 upstream-timeout or a connect timeout —
              SUBSTITUTE a reachable county/state layer (prefer *.arcgis.com / county GIS over
              a single city's self-hosted ArcGIS box) and log the swap in the decisions log.

Exit code: 0 if all reachable/responded; 1 if any source blocked or unreachable.
"""
import json
import sys
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402

TIMEOUT = 20
H = {"User-Agent": "Mozilla/5.0"}


def root_of(url):
    u = url.rstrip("/")
    if u.lower().endswith("/query"):
        u = u[: -len("/query")]
    return u.rstrip("/")


def probe(url):
    root = root_of(url)
    purl = root + ("&" if "?" in root else "?") + "f=json"
    try:
        r = requests.get(purl, timeout=TIMEOUT, headers=H, verify=False)
    except requests.exceptions.Timeout:
        return "unreachable", f"timeout (no response in {TIMEOUT}s) — allowed but origin not responding"
    except requests.exceptions.ConnectionError as e:
        return "unreachable", f"connection error — {str(e)[:90]}"
    except Exception as e:  # noqa: BLE001
        return "error", str(e)[:90]
    deny = (r.headers.get("x-deny-reason") or "").lower()
    body = (r.text or "")[:500]
    if deny == "host_not_allowed" or "host_not_allowed" in body:
        return "blocked", "x-deny-reason=host_not_allowed (allowlist not in effect for this session)"
    if r.status_code == 200:
        return "reachable", "HTTP 200"
    if r.status_code in (502, 503, 504) and (
            "upstream connect" in body or "connection timeout" in body or not r.headers.get("server")):
        return "unreachable", f"HTTP {r.status_code} upstream connect/timeout (allowed but origin unreachable)"
    return "responded", f"HTTP {r.status_code}"


def main():
    import urllib3
    urllib3.disable_warnings()
    cfg = C.load()
    IN = C.indir(cfg)
    sources = ([("parcels", s.get("name") or s.get("county", ""), s["url"]) for s in cfg.get("parcel_sources", [])] +
               [("zoning", s["jurisdiction"], s["url"]) for s in cfg.get("zoning_sources", [])])
    if not sources:
        print("[preflight] no sources in config — nothing to probe.")
        return

    results, bad = [], 0
    print("\n================ REACHABILITY PRE-FLIGHT ================")
    for kind, label, url in sources:
        verdict, detail = probe(url)
        flag = "OK " if verdict in ("reachable", "responded") else "!! "
        if verdict in ("blocked", "unreachable", "error"):
            bad += 1
        print(f"  {flag}[{kind:7}] {label[:34]:34} {verdict:11} {detail}")
        print(f"        {url}")
        results.append({"kind": kind, "label": label, "url": url,
                        "verdict": verdict, "detail": detail})
    (IN / "preflight.json").write_text(json.dumps(results, indent=2), encoding="utf-8")

    if bad:
        print(f"\n*** {bad} source(s) not usable. Fix before pulling: ***")
        print("  - 'blocked'     -> allowlist not in effect. Add the domain, then START A NEW SESSION")
        print("                     (the sandbox reads the network policy once, at session start).")
        print("  - 'unreachable' -> host is allowed but down from the sandbox. SUBSTITUTE a reachable")
        print("                     county/state layer (prefer *.arcgis.com / county GIS over a city's")
        print("                     self-hosted ArcGIS box). Re-author config.json, re-run `endpoints`")
        print("                     if the host changed, log the swap in Tables/decisions_log.md, retry.")
        sys.exit(1)
    print("\nAll sources reachable — safe to pull.")


if __name__ == "__main__":
    main()
