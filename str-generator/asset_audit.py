"""
Read-only export for the client asset clean-up: every Halo asset plus every N-sight
workstation and server, saved as JSON for review. Nothing in Halo or N-sight is changed.

  ..\\.venv\\Scripts\\python.exe asset_audit.py

Takes a few minutes (N-sight allows about one call per second).
Writes output\\asset-audit\\halo_assets.json, halo_clients.json, nsight_devices.json.
"""

from __future__ import annotations

import json
import logging
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import str_generator as G  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
log = logging.getLogger("audit")

OUT = HERE / "output" / "asset-audit"


def halo_all(c, path: str, key: str, params: dict) -> list:
    out, page = [], 1
    while True:
        body = c.get(path, {**params, "pageinate": True, "page_size": 100, "page_no": page})
        batch = (body or {}).get(key) or []
        out.extend(batch)
        total = (body or {}).get("record_count")
        if not batch or (total and len(out) >= total) or (not total and len(batch) < 100) or page > 500:
            log.info("  %s: %s of %s", path, len(out), total)
            return out
        page += 1


def main() -> int:
    G.load_config()
    OUT.mkdir(parents=True, exist_ok=True)
    H, N = G.halo_mod(), G.nsight_mod()
    c = H.halo()

    log.info("Halo clients…")
    clients = halo_all(c, "Client", "clients", {"includeinactive": True})
    keep_c = ("id", "name", "inactive", "gficlientid", "accountsid")
    (OUT / "halo_clients.json").write_text(json.dumps([{k: x.get(k) for k in keep_c} for x in clients], indent=1), encoding="utf-8")

    log.info("Halo assets (active and inactive)…")
    assets = halo_all(c, "Asset", "assets", {"includeinactive": True})
    (OUT / "halo_assets.json").write_text(json.dumps(assets, indent=1, default=str), encoding="utf-8")
    log.info("  %s assets", len(assets))

    if "--halo-only" in sys.argv:
        log.info("Done (Halo only). Files in %s", OUT)
        return 0

    log.info("N-sight devices…")
    ns = N.ns()
    devices = []
    ncl = N._items(ns.call("list_clients"), "client")
    for i, cl in enumerate(ncl, 1):
        cid, cname = N._find(cl, "clientid"), N._find(cl, "name")
        try:
            for s in N._items(ns.call("list_sites", clientid=cid), "site"):
                sid, sname = N._find(s, "siteid"), N._find(s, "name")
                for kind, call, tag in (("workstation", "list_workstations", "workstation"), ("server", "list_servers", "server")):
                    for d in N._items(ns.call(call, siteid=sid), tag):
                        d = {k: v for k, v in dict(d).items() if k not in ("os_product_key",)}
                        d.update({"_kind": kind, "_clientid": cid, "_client": cname, "_siteid": sid, "_site": sname})
                        devices.append(d)
        except Exception as e:  # noqa: BLE001
            log.warning("  %s failed: %s", cname, e)
        log.info("  %s/%s %s (%s devices so far)", i, len(ncl), cname, len(devices))
    (OUT / "nsight_devices.json").write_text(json.dumps(devices, indent=1, default=str), encoding="utf-8")
    (OUT / "done.txt").write_text(time.strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8")
    log.info("Done. Files in %s", OUT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
