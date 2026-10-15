"""
Microsoft Graph helpers for the STR generator (app-only, client credentials).

Lists it manages on the configured SharePoint site:
  STR Clients   – one item per active Halo client (Title, HaloId). Feeds the dropdown.
  STR Requests  – Client (lookup to STR Clients), Status, Message, Link. Created by the
                  "Generate STR" button; the generator picks up Status = Requested.
"""

from __future__ import annotations

import logging
import os
import time
from typing import Any
from urllib.parse import quote

import httpx

log = logging.getLogger("str.sp")
GRAPH = "https://graph.microsoft.com/v1.0"
KEYRING_SERVICE = "str-generator"


def _secret() -> str:
    v = os.getenv("STR_GRAPH_CLIENT_SECRET")
    if not v:
        import keyring
        v = keyring.get_password(KEYRING_SERVICE, "graph_client_secret")
    if not v:
        raise RuntimeError("Graph client secret missing. Run setup.ps1, or: python -m keyring set str-generator graph_client_secret")
    return v


class SharePoint:
    def __init__(self, cfg: dict):
        g = cfg["graph"]
        self.tenant, self.client_id = g["tenant_id"], g["client_id"]
        self.hostname = g.get("hostname", "yourco.sharepoint.com")
        self.site_path = g.get("site_path", "")          # "" = root site, or "sites/Operations"
        self.library = g.get("library", "SHAREPOINT")    # document library holding Clients/<Client>
        self.subfolder = g.get("client_subfolder", "Strategic Tech Review")
        self.http = httpx.Client(timeout=120)
        self._tok, self._exp = None, 0.0
        self._site_id = self._drive_id = None
        self._lists: dict[str, str] = {}

    # -- auth / requests ------------------------------------------------------
    def token(self) -> str:
        if self._tok and time.time() < self._exp - 60:
            return self._tok
        r = self.http.post(f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/token", data={
            "grant_type": "client_credentials", "client_id": self.client_id, "client_secret": _secret(),
            "scope": "https://graph.microsoft.com/.default"})
        if r.status_code != 200:
            raise RuntimeError(f"Graph sign-in failed ({r.status_code}): {r.text[:300]}")
        b = r.json()
        self._tok, self._exp = b["access_token"], time.time() + int(b.get("expires_in", 3600))
        return self._tok

    def req(self, method: str, url: str, **kw) -> Any:
        url = url if url.startswith("http") else GRAPH + url
        headers = {"Authorization": f"Bearer {self.token()}", **kw.pop("headers", {})}
        for attempt in range(4):
            r = self.http.request(method, url, headers=headers, **kw)
            if r.status_code in (429, 503):
                time.sleep(min(int(r.headers.get("Retry-After", "5") or 5), 30))
                continue
            if r.status_code >= 400:
                raise RuntimeError(f"Graph {method} {url} failed ({r.status_code}): {r.text[:400]}")
            return r.json() if r.content else None
        raise RuntimeError(f"Graph {method} {url} kept being throttled")

    def paged(self, url: str) -> list:
        out = []
        while url:
            b = self.req("GET", url)
            out += b.get("value", [])
            url = b.get("@odata.nextLink")
        return out

    # -- site, drive, lists ---------------------------------------------------
    @property
    def site_id(self) -> str:
        if not self._site_id:
            path = f"/sites/{self.hostname}" + (f":/{self.site_path.strip('/')}" if self.site_path else "")
            self._site_id = self.req("GET", path)["id"]
        return self._site_id

    @property
    def drive_id(self) -> str:
        if not self._drive_id:
            drives = self.req("GET", f"/sites/{self.site_id}/drives")["value"]
            want = self.library.strip("/").lower()
            match = next((d for d in drives if d.get("webUrl", "").rstrip("/").split("/")[-1].lower() == want
                          or d.get("name", "").lower() == want), None)
            if not match:
                raise RuntimeError(f"Library '{self.library}' not found. Libraries: " + ", ".join(d.get("name", "") for d in drives))
            self._drive_id = match["id"]
        return self._drive_id

    def list_id(self, name: str) -> str | None:
        if name not in self._lists:
            found = self.req("GET", f"/sites/{self.site_id}/lists?$filter=displayName eq '{name}'")["value"]
            if not found:
                return None
            self._lists[name] = found[0]["id"]
        return self._lists[name]

    def ensure_lists(self) -> dict:
        """Create STR Clients and STR Requests if missing. Returns their web URLs."""
        if not self.list_id("STR Clients"):
            b = self.req("POST", f"/sites/{self.site_id}/lists", json={
                "displayName": "STR Clients", "list": {"template": "genericList"},
                "columns": [{"name": "HaloId", "number": {}}]})
            self._lists["STR Clients"] = b["id"]
        if not self.list_id("STR Requests"):
            b = self.req("POST", f"/sites/{self.site_id}/lists", json={
                "displayName": "STR Requests", "list": {"template": "genericList"},
                "columns": [
                    {"name": "Client", "lookup": {"listId": self._lists["STR Clients"], "columnName": "Title"}},
                    {"name": "Quarter", "text": {}, "description": "Optional, e.g. 2026-Q3. Blank = last completed quarter."},
                    {"name": "Status", "choice": {"choices": ["Requested", "Working", "Done", "Failed"]}, "defaultValue": {"value": "Requested"}},
                    {"name": "Message", "text": {"allowMultipleLines": True}},
                    {"name": "Link", "text": {}},
                ]})
            self._lists["STR Requests"] = b["id"]
        self.ensure_columns()
        lists = {}
        for n in ("STR Clients", "STR Requests"):
            lists[n] = self.req("GET", f"/sites/{self.site_id}/lists/{self.list_id(n)}")["webUrl"]
        return lists

    def ensure_columns(self) -> list[str]:
        """Add optional columns to STR Requests that older installs don't have. Returns the names added."""
        lid = self.list_id("STR Requests")
        if not lid:
            return []
        have = {c["name"] for c in self.paged(f"/sites/{self.site_id}/lists/{lid}/columns")}
        added = []
        for col in (
            {"name": "MainContact", "displayName": "Main contact", "text": {},
             "description": "Optional. Name or email of the contact to list first."},
            {"name": "ReviewDate", "displayName": "Review date", "dateTime": {"format": "dateOnly"},
             "description": "Optional. Date of the review meeting, shown on the cover. Blank = [day] placeholder."},
        ):
            if col["name"] not in have:
                self.req("POST", f"/sites/{self.site_id}/lists/{lid}/columns", json=col)
                added.append(col["name"])
        return added

    # -- clients dropdown ----------------------------------------------------
    def sync_clients(self, clients: list[dict]) -> tuple[int, int]:
        lid = self.list_id("STR Clients")
        existing = {int(i["fields"].get("HaloId") or 0): i for i in self.paged(f"/sites/{self.site_id}/lists/{lid}/items?expand=fields&$top=999")}
        added = updated = 0
        for c in clients:
            title = c["name"].rstrip("*").strip() or c["name"]
            cur = existing.get(int(c["id"]))
            if not cur:
                self.req("POST", f"/sites/{self.site_id}/lists/{lid}/items", json={"fields": {"Title": title, "HaloId": int(c["id"])}})
                added += 1
            elif cur["fields"].get("Title") != title:
                self.req("PATCH", f"/sites/{self.site_id}/lists/{lid}/items/{cur['id']}/fields", json={"Title": title})
                updated += 1
        return added, updated

    def halo_id_for_client_item(self, item_id: Any) -> int | None:
        lid = self.list_id("STR Clients")
        f = self.req("GET", f"/sites/{self.site_id}/lists/{lid}/items/{item_id}?expand=fields")["fields"]
        return int(f["HaloId"]) if f.get("HaloId") is not None else None

    # -- requests --------------------------------------------------------------
    def pending_requests(self) -> list[dict]:
        lid = self.list_id("STR Requests")
        items = self.paged(f"/sites/{self.site_id}/lists/{lid}/items?expand=fields&$top=200")
        return [i for i in items if (i["fields"].get("Status") or "Requested") == "Requested"]

    def update_request(self, item_id: str, **fields) -> None:
        lid = self.list_id("STR Requests")
        self.req("PATCH", f"/sites/{self.site_id}/lists/{lid}/items/{item_id}/fields", json=fields)

    # -- upload ----------------------------------------------------------------
    def upload(self, local_path: str, folder: str, filename: str) -> str:
        """Upload into <library>/<folder>/<subfolder>/<filename>; returns the file's web URL."""
        dest = "/".join(p.strip("/") for p in (folder, self.subfolder, filename) if p)
        url = f"/drives/{self.drive_id}/root:/{quote(dest)}:/content"
        with open(local_path, "rb") as f:
            data = f.read()
        if len(data) < 4 * 1024 * 1024:
            item = self.req("PUT", url, content=data, headers={"Content-Type": "application/octet-stream"})
        else:
            sess = self.req("POST", f"/drives/{self.drive_id}/root:/{quote(dest)}:/createUploadSession",
                            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}})
            up, size, chunk = sess["uploadUrl"], len(data), 5 * 1024 * 1024
            item = None
            for start in range(0, size, chunk):
                end = min(start + chunk, size) - 1
                r = self.http.put(up, content=data[start:end + 1], headers={"Content-Range": f"bytes {start}-{end}/{size}"})
                r.raise_for_status()
                item = r.json() if r.content else item
        return (item or {}).get("webUrl", "")

    def folder_url(self, folder: str) -> str:
        dest = "/".join(p.strip("/") for p in (folder, self.subfolder) if p)
        try:
            return self.req("GET", f"/drives/{self.drive_id}/root:/{quote(dest)}")["webUrl"]
        except Exception:  # noqa: BLE001
            return ""
