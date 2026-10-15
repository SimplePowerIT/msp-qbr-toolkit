"""
N-able N-sight RMM MCP server (read-only) for the Claude desktop app.

Uses the N-sight Data Extraction API to list clients, sites and devices and to pull
per-device hardware details (model, processor, RAM, OS, last user) for QBR hardware
tables. Nothing here changes anything in N-sight.

Configuration:
  NSIGHT_SERVER   API server for your region, e.g. www.am.remote.management  (required,
                  set in claude_desktop_config.json)
  NSIGHT_API_KEY  normally read from Windows Credential Manager
                  (service "nsight-mcp", user "nsight_api_key"); set only to override.

Run `python server.py --test` to check the connection from a terminal.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
import threading
import time
import xml.etree.ElementTree as ET
from typing import Any

import httpx
from mcp.server.fastmcp import FastMCP

logging.basicConfig(
    stream=sys.stderr,
    level=os.getenv("NSIGHT_LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s nsight-mcp: %(message)s",
)
log = logging.getLogger("nsight-mcp")
# httpx logs full request URLs at INFO, and N-sight puts the API key in the URL: keep it out of logs.
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

KEYRING_SERVICE = "nsight-mcp"
MIN_INTERVAL = 1.05  # N-sight asks for at most 60 calls per minute


# --------------------------------------------------------------------------- #
# Config / client
# --------------------------------------------------------------------------- #
def _cfg(name: str, secret: bool = False) -> str:
    val = os.getenv(name)
    if not val and secret:
        try:
            import keyring

            val = keyring.get_password(KEYRING_SERVICE, name.lower())
        except Exception as e:
            log.warning("Could not read %s from Credential Manager: %s", name, e)
    if not val:
        hint = (
            " Run setup.ps1, or store it with: python -m keyring set nsight-mcp nsight_api_key"
            if secret
            else " Add it to the env block in claude_desktop_config.json."
        )
        raise RuntimeError(f"{name} is not set.{hint}")
    return val


def _xml_to_obj(el: ET.Element) -> Any:
    """Element -> dict / text. Repeated child tags become lists."""
    children = list(el)
    if not children:
        return (el.text or "").strip()
    out: dict[str, Any] = {}
    for c in children:
        v = _xml_to_obj(c)
        if c.tag in out:
            if not isinstance(out[c.tag], list):
                out[c.tag] = [out[c.tag]]
            out[c.tag].append(v)
        else:
            out[c.tag] = v
    return out


def _as_list(v: Any) -> list:
    if v is None or v == "":
        return []
    return v if isinstance(v, list) else [v]


def _find(obj: Any, *names: str) -> Any:
    """First value whose key matches one of names (case-insensitive), searched depth-first."""
    want = [n.lower() for n in names]
    if isinstance(obj, dict):
        for n in want:
            for k, v in obj.items():
                if k.lower() == n and v not in ("", None, {}):
                    return v
        for v in obj.values():
            r = _find(v, *names)
            if r not in (None, ""):
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = _find(v, *names)
            if r not in (None, ""):
                return r
    return None


def _dicts(obj: Any) -> list[dict]:
    """Every dict nested anywhere in obj."""
    out = []
    if isinstance(obj, dict):
        out.append(obj)
        for v in obj.values():
            out.extend(_dicts(v))
    elif isinstance(obj, list):
        for v in obj:
            out.extend(_dicts(v))
    return out


class NsightClient:
    def __init__(self) -> None:
        server = _cfg("NSIGHT_SERVER").strip()
        server = re.sub(r"^https?://", "", server).strip("/").removesuffix("/api")
        self.url = f"https://{server}/api/"
        self.key = _cfg("NSIGHT_API_KEY", secret=True)
        self.http = httpx.Client(timeout=60)
        self._lock = threading.Lock()
        self._last = 0.0
        self.cache: dict[str, Any] = {}

    def call(self, service: str, **params: Any) -> dict:
        key = service + repr(sorted(params.items()))
        if key in self.cache:
            return self.cache[key]
        with self._lock:
            wait = MIN_INTERVAL - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.time()
            r = self.http.get(self.url, params={"apikey": self.key, "service": service, **params})
        if r.status_code >= 400:
            raise RuntimeError(f"N-sight {service} failed ({r.status_code}): {r.text[:300]}")
        try:
            root = ET.fromstring(r.content)
        except ET.ParseError:
            raise RuntimeError(f"N-sight {service} returned non-XML: {r.text[:300]}")
        if root.get("status", "OK").upper() != "OK":
            msg = _find(_xml_to_obj(root), "message", "errormessage") or r.text[:300]
            raise RuntimeError(f"N-sight {service} error: {msg}")
        body = _xml_to_obj(root)
        body = body if isinstance(body, dict) else {}
        self.cache[key] = body
        return body


_client: NsightClient | None = None


def ns() -> NsightClient:
    global _client
    if _client is None:
        _client = NsightClient()
    return _client


def _items(body: dict, tag: str) -> list[dict]:
    items = body.get("items", body)
    if isinstance(items, dict):
        return [x for x in _as_list(items.get(tag)) if isinstance(x, dict)]
    return []


# --------------------------------------------------------------------------- #
# Helpers that shape devices for QBR tables
# --------------------------------------------------------------------------- #
def _workstations(clientid: int, site_ids: list[int] | None = None) -> list[dict]:
    c = ns()
    sites = _items(c.call("list_sites", clientid=clientid), "site")
    out = []
    for s in sites:
        sid = int(_find(s, "siteid") or 0)
        if site_ids and sid not in site_ids:
            continue
        for w in _items(c.call("list_workstations", siteid=sid), "workstation"):
            w = dict(w)
            w["_site"] = _find(s, "name")
            w["_siteid"] = sid
            out.append(w)
    return out


def _device_id(w: dict) -> int | None:
    v = _find(w, "workstationid", "deviceid", "id", "wsid")
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _detail(d: dict, key: str) -> str | None:
    m = re.search(rf"(?:^|\n){key}=([^\n]*)", str(d.get("details") or ""))
    return m.group(1).strip() if m else None


# N-sight hardware item types are numeric: 2 BIOS, 4 system board, 9 disk, 13 processor.
def _hw_summary(details: dict) -> dict:
    processors, bios, serial, disks = [], [], None, []
    for d in _dicts(details):
        t = str(d.get("type", "") or d.get("hardwaretype", "")).lower()
        name = d.get("name") or d.get("description")
        if not isinstance(name, str) or not name or str(d.get("deleted", "0")) == "1":
            continue
        if t == "13" or "processor" in t or "cpu" in t:
            processors.append(name)
        elif t == "2" or "bios" in t:
            bios.append(name)
        elif t == "4" and not serial:
            serial = _detail(d, "SerialNumber")
        elif t == "9" and "USB" not in name:
            size = _detail(d, "Size")
            if size and size.isdigit():
                disks.append(f"{name} ({round(int(size) / 1e9)} GB)")
    custom = {
        k: v
        for d in _dicts(details)
        for k, v in d.items()
        if "custom" in k.lower() and isinstance(v, str) and v
    }
    return {
        "manufacturer": _find(details, "manufacturer"),
        "model": _find(details, "model"),
        "serial": serial,
        "windows_product_id": _find(details, "serialnumber"),
        "disks": disks or None,
        "chassis": _find(details, "chassistype", "chassis_type", "chassis"),
        "os": _find(details, "os", "operatingsystem", "os_name"),
        "os_install_date": _find(details, "osinstalldate", "os_install_date", "installdate"),
        "ram": _find(details, "ram", "memory", "totalmemory"),
        "processor": re.sub(r"\s+", " ", sorted(set(processors))[0]) if processors else None,
        "bios": bios[0] if bios else None,
        "custom_fields": custom or None,
        "scan_time": _find(details, "scantime", "scan_time"),
    }


# --------------------------------------------------------------------------- #
# MCP server + tools
# --------------------------------------------------------------------------- #
mcp = FastMCP(
    "nsight",
    instructions=(
        "Read-only access to N-able N-sight RMM. Use list_clients to find a clientid, "
        "then hardware_inventory for a QBR hardware table (model, PC name, last user, "
        "processor, RAM, OS). The API allows 60 calls a minute, so hardware_inventory "
        "pages through devices: call again with next_offset until it is null. "
        "Use device_asset_details for every raw field on one device."
    ),
)


@mcp.tool()
def test_connection() -> dict:
    """Check that the N-sight API key and server work. Returns the server and client count."""
    c = ns()
    return {"ok": True, "server": c.url, "clients_found": len(_items(c.call("list_clients"), "client"))}


@mcp.tool()
def list_clients(search: str | None = None) -> list[dict]:
    """List N-sight clients (clientid and name). `search` filters by name."""
    out = []
    for cl in _items(ns().call("list_clients"), "client"):
        name = _find(cl, "name") or ""
        if search and search.lower() not in str(name).lower():
            continue
        out.append({"clientid": _find(cl, "clientid"), "name": name})
    return out


@mcp.tool()
def list_sites(clientid: int) -> list[dict]:
    """Sites for one N-sight client."""
    return [
        {"siteid": _find(s, "siteid"), "name": _find(s, "name")}
        for s in _items(ns().call("list_sites", clientid=clientid), "site")
    ]


@mcp.tool()
def list_workstations(clientid: int, site_ids: list[int] | None = None) -> list[dict]:
    """Workstations for a client across its sites (or only `site_ids`), with the fields N-sight returns."""
    return _workstations(clientid, site_ids)


@mcp.tool()
def list_servers(clientid: int) -> list[dict]:
    """Servers for a client across all its sites."""
    c = ns()
    out = []
    for s in _items(c.call("list_sites", clientid=clientid), "site"):
        sid = int(_find(s, "siteid") or 0)
        for sv in _items(c.call("list_servers", siteid=sid), "server"):
            sv = dict(sv)
            sv["_site"] = _find(s, "name")
            out.append(sv)
    return out


@mcp.tool()
def device_asset_details(deviceid: int) -> dict:
    """Every asset field N-sight holds for one device: hardware list, software, custom fields."""
    return ns().call("list_device_asset_details", deviceid=deviceid)


@mcp.tool()
def hardware_inventory(clientid: int, offset: int = 0, limit: int = 25, include_software: bool = False) -> dict:
    """
    QBR hardware rows for a client's workstations: PC name, site, last user, manufacturer,
    model, serial, processor, RAM, OS, OS install date, BIOS and custom fields (e.g. Warranty).
    Pages `limit` devices at a time to stay inside the 60-calls-a-minute API limit; call
    again with `offset=next_offset` until next_offset is null.
    """
    ws = sorted(_workstations(clientid), key=lambda w: str(_find(w, "name") or ""))
    rows = []
    for w in ws[offset : offset + limit]:
        did = _device_id(w)
        row = {
            "deviceid": did,
            "pc_name": _find(w, "name"),
            "site": w.get("_site"),
            "last_user": _find(w, "user", "username", "last_user", "lastloggedinuser"),
            "online": _find(w, "online"),
            "last_boot": _find(w, "last_boot_time"),
            "agent_install_date": _find(w, "install_date"),
        }
        if did is not None:
            try:
                det = ns().call("list_device_asset_details", deviceid=did)
                row.update(_hw_summary(det))
                if include_software:
                    row["software"] = [
                        _find(s, "name") for s in _dicts(det) if "version" in s and _find(s, "name")
                    ][:200]
            except Exception as e:
                row["error"] = str(e)
        for k in ("model", "manufacturer", "os"):
            if not row.get(k):
                row[k] = _find(w, k)
        rows.append(row)
    nxt = offset + limit
    return {"total_workstations": len(ws), "offset": offset, "rows": rows, "next_offset": nxt if nxt < len(ws) else None}


def _self_test() -> int:
    try:
        r = test_connection()
        print(f"OK: {r['server']} - {r['clients_found']} clients")
        return 0
    except Exception as e:
        print(f"FAILED: {e}")
        return 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--test", action="store_true", help="check the API key and exit")
    if ap.parse_args().test:
        sys.exit(_self_test())
    mcp.run()
