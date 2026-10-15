"""
Halo PSA MCP server (read-only) for the Claude desktop app.

Tools cover QBR reporting and ticket triage *viewing*. Nothing here writes to Halo.

Configuration (environment variables, set in claude_desktop_config.json):
  HALO_BASE_URL       e.g. https://yourco.halopsa.com           (required)
  HALO_CLIENT_ID      API application Client ID                 (required)
  HALO_TENANT         tenant name (Halo-hosted only)            (optional)
  HALO_SCOPE          token scope, default "all"                (optional)
  HALO_AUTH_URL       override, default {base}/auth/token       (optional)
  HALO_API_URL        override, default {base}/api              (optional)
  HALO_CLIENT_SECRET  normally read from Windows Credential Manager
                      (service "halo-mcp", user "halo_client_secret");
                      set this env var only to override.

Run `python server.py --test` to check the connection from a terminal.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import statistics
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from typing import Any

import httpx
from mcp.server.fastmcp import FastMCP

# stdio transport: stdout is reserved for MCP protocol messages, so log to stderr only.
logging.basicConfig(
    stream=sys.stderr,
    level=os.getenv("HALO_LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s halo-mcp: %(message)s",
)
log = logging.getLogger("halo-mcp")

KEYRING_SERVICE = "halo-mcp"
MAX_RECORDS_CAP = 5000


# --------------------------------------------------------------------------- #
# Config / client
# --------------------------------------------------------------------------- #
def _cfg(name: str, required: bool = True, secret: bool = False) -> str | None:
    val = os.getenv(name)
    if not val and secret:
        try:
            import keyring

            val = keyring.get_password(KEYRING_SERVICE, name.lower())
        except Exception as e:  # keyring missing or backend error
            log.warning("Could not read %s from Credential Manager: %s", name, e)
    if required and not val:
        hint = (
            " Run setup.ps1, or store it with: python -m keyring set halo-mcp halo_client_secret"
            if secret
            else " Add it to the env block in claude_desktop_config.json."
        )
        raise RuntimeError(f"{name} is not set.{hint}")
    return val


def _clean(params: dict | None) -> dict | None:
    if not params:
        return None
    out = {}
    for k, v in params.items():
        if v is None or v == "":
            continue
        out[k] = str(v).lower() if isinstance(v, bool) else v
    return out


def _items(payload: Any, key: str | None = None) -> list:
    """Pull the record list out of a Halo response, tolerating shape differences."""
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        if key and isinstance(payload.get(key), list):
            return payload[key]
        for v in payload.values():
            if isinstance(v, list):
                return v
    return []


class HaloClient:
    def __init__(self) -> None:
        self.base = _cfg("HALO_BASE_URL").rstrip("/")
        self.auth_url = os.getenv("HALO_AUTH_URL") or f"{self.base}/auth/token"
        self.api_url = (os.getenv("HALO_API_URL") or f"{self.base}/api").rstrip("/")
        self.tenant = os.getenv("HALO_TENANT") or None
        self.client_id = _cfg("HALO_CLIENT_ID")
        self.client_secret = _cfg("HALO_CLIENT_SECRET", secret=True)
        self.scope = os.getenv("HALO_SCOPE", "all")
        self._token: str | None = None
        self._expires = 0.0
        self.http = httpx.Client(timeout=60)
        self._lookups: dict[str, dict] = {}

    # -- auth ----------------------------------------------------------------
    def token(self) -> str:
        if self._token and time.time() < self._expires - 60:
            return self._token
        r = self.http.post(
            self.auth_url,
            params={"tenant": self.tenant} if self.tenant else None,
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": self.scope,
            },
        )
        if r.status_code != 200:
            raise RuntimeError(f"Halo token request failed ({r.status_code}): {r.text[:300]}")
        body = r.json()
        self._token = body["access_token"]
        self._expires = time.time() + int(body.get("expires_in", 3600))
        return self._token

    # -- requests ------------------------------------------------------------
    def get(self, path: str, params: dict | None = None) -> Any:
        url = f"{self.api_url}/{path.lstrip('/')}"
        for attempt in range(4):
            r = self.http.get(url, params=_clean(params), headers={"Authorization": f"Bearer {self.token()}"})
            if r.status_code == 401 and attempt == 0:
                self._token = None
                continue
            if r.status_code == 429:
                time.sleep(min(int(r.headers.get("Retry-After", "5") or 5), 30))
                continue
            if r.status_code >= 400:
                raise RuntimeError(f"Halo GET {path} failed ({r.status_code}): {r.text[:300]}")
            return r.json() if r.content else None
        raise RuntimeError(f"Halo GET {path} failed after retries (rate limited or unauthorized).")

    def get_paged(self, path: str, list_key: str, params: dict | None, max_records: int) -> list:
        max_records = max(1, min(max_records, MAX_RECORDS_CAP))
        size = min(100, max_records)
        out: list = []
        page = 1
        while len(out) < max_records:
            payload = self.get(path, {**(params or {}), "pageinate": True, "page_size": size, "page_no": page})
            batch = _items(payload, list_key)
            out.extend(batch)
            total = payload.get("record_count") if isinstance(payload, dict) else None
            if not batch or len(batch) < size or (total is not None and len(out) >= total):
                break
            page += 1
        return out[:max_records]

    def lookup(self, path: str) -> dict:
        """id -> name map for small reference lists (Status, TicketType, Agent). Cached."""
        if path not in self._lookups:
            try:
                self._lookups[path] = {
                    i.get("id"): i.get("name") for i in _items(self.get(path)) if isinstance(i, dict)
                }
            except Exception as e:
                log.warning("Lookup %s failed: %s", path, e)
                self._lookups[path] = {}
        return self._lookups[path]

    def name_of(self, path: str, item_id: Any) -> Any:
        """
        Name for one id from a reference list. The list endpoints leave out inactive records
        (e.g. agents who have left, retired ticket types), so on a miss fetch {path}/{id}
        directly, which returns inactive records too. Falls back to the id. Cached.
        """
        if item_id in (None, "", 0, -1):
            return item_id
        names = self.lookup(path)
        if item_id in names and names[item_id]:
            return names[item_id]
        name = None
        try:
            rec = self.get(f"{path}/{item_id}")
            if isinstance(rec, dict):
                name = rec.get("name") or rec.get("agent_name")
                if name and (rec.get("inactive") or rec.get("is_disabled")):
                    name = f"{name} (inactive)"
        except Exception as e:
            log.warning("Lookup %s/%s failed: %s", path, item_id, e)
        names[item_id] = name or item_id
        return names[item_id]


_client: HaloClient | None = None


def halo() -> HaloClient:
    global _client
    if _client is None:
        _client = HaloClient()
    return _client


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _dt(value: Any) -> datetime | None:
    """Parse Halo date strings; Halo uses 1900-01-01 for 'empty'."""
    if not value or not isinstance(value, str):
        return None
    try:
        d = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if d.year < 1950:
        return None
    return d if d.tzinfo else d.replace(tzinfo=timezone.utc)


def _first(d: dict, *keys: str) -> Any:
    for k in keys:
        v = d.get(k)
        if v not in (None, "", 0) or (v == 0 and k.endswith("_id")):
            return v
    return None


def _scalars(d: dict, extra_drop: tuple[str, ...] = ()) -> dict:
    """Keep only scalar fields with meaningful values (drops nested objects and Halo's empty dates)."""
    out = {}
    for k, v in d.items():
        if k in extra_drop or isinstance(v, (dict, list)) or v in (None, ""):
            continue
        if isinstance(v, str) and v.startswith("1900-01-01"):
            continue
        out[k] = v
    return out


def _period(quarter: str | None, start_date: str | None, end_date: str | None) -> tuple[str, str]:
    """Resolve a quarter like '2026-Q3' or explicit ISO dates. Defaults to the last 90 days."""
    if quarter:
        try:
            y, q = quarter.upper().replace(" ", "").split("-Q")
            y, q = int(y), int(q)
            assert 1 <= q <= 4
        except Exception:
            raise ValueError("quarter must look like '2026-Q3'")
        start = date(y, 3 * (q - 1) + 1, 1)
        end = date(y + 1, 1, 1) if q == 4 else date(y, 3 * q + 1, 1)
        return start.isoformat(), (end - timedelta(days=1)).isoformat()
    end = end_date or date.today().isoformat()
    start = start_date or (date.fromisoformat(end[:10]) - timedelta(days=90)).isoformat()
    return start[:10], end[:10]


def _slim_ticket(t: dict, with_details: bool = False) -> dict:
    c = halo()
    out = {
        "id": t.get("id"),
        "summary": t.get("summary"),
        "client": _first(t, "client_name"),
        "site": _first(t, "site_name"),
        "user": _first(t, "user_name"),
        "status": c.name_of("Status", t.get("status_id")),
        "type": c.name_of("TicketType", t.get("tickettype_id")),
        "priority_id": t.get("priority_id"),
        "category": _first(t, "category_1"),
        "team": _first(t, "team"),
        "agent": c.name_of("Agent", t.get("agent_id")) if t.get("agent_id") else "Unassigned",
        "opened": t.get("dateoccurred"),
        "closed": _first(t, "dateclosed", "datecleared"),
        "fix_by": _first(t, "fixbydate"),
    }
    if with_details:
        details = t.get("details") or t.get("details_html") or ""
        out["details"] = details[:2000]
    return {k: v for k, v in out.items() if v not in (None, "") and not str(v).startswith("1900-01-01")}


def _fetch_tickets(
    client_id: int | None,
    start: str | None,
    end: str | None,
    date_field: str,
    status: str,
    search: str | None,
    max_results: int,
) -> list[dict]:
    params: dict[str, Any] = {"client_id": client_id, "search": search}
    if start or end:
        params.update({"startdate": start, "enddate": end, "datesearch": date_field})
    if status == "open":
        params["open_only"] = True
    elif status == "closed":
        params["closed_only"] = True
    return halo().get_paged("Tickets", "tickets", params, max_results)


# --------------------------------------------------------------------------- #
# MCP server + tools
# --------------------------------------------------------------------------- #
mcp = FastMCP(
    "halo",
    instructions=(
        "Read-only access to Halo PSA. Use list_clients to find a client_id, then "
        "ticket_summary for QBR metrics, search_tickets/get_ticket for detail, "
        "list_assets and list_contracts for lifecycle and agreement context, and "
        "new_tickets to review recent tickets for triage suggestions."
    ),
)


@mcp.tool()
def test_connection() -> dict:
    """Check that the Halo API credentials work. Returns the instance URL and a sample count."""
    c = halo()
    c.token()
    statuses = c.lookup("Status")
    return {"ok": True, "base_url": c.base, "tenant": c.tenant, "ticket_statuses_found": len(statuses)}


@mcp.tool()
def list_clients(search: str | None = None, include_inactive: bool = False, max_results: int = 200) -> list[dict]:
    """List Halo clients (customers). Use `search` to filter by name. Returns id and name for each."""
    rows = halo().get_paged(
        "Client", "clients", {"search": search, "includeinactive": include_inactive}, max_results
    )
    return [
        {k: v for k, v in {"id": r.get("id"), "name": r.get("name"), "inactive": r.get("inactive")}.items()
         if v is not None}
        for r in rows
    ]


@mcp.tool()
def get_client(client_id: int) -> dict:
    """Full details for one client (scalar fields only)."""
    return _scalars(halo().get(f"Client/{client_id}", {"includedetails": True}) or {})


@mcp.tool()
def search_tickets(
    client_id: int | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    date_field: str = "dateoccurred",
    status: str = "all",
    search: str | None = None,
    max_results: int = 200,
) -> list[dict]:
    """
    Search tickets. Dates are ISO (YYYY-MM-DD). `date_field` is what the date range applies to:
    'dateoccurred' (opened) or 'dateclosed'. `status`: 'all', 'open' or 'closed'.
    """
    rows = _fetch_tickets(client_id, start_date, end_date, date_field, status, search, max_results)
    return [_slim_ticket(t) for t in rows]


@mcp.tool()
def get_ticket(ticket_id: int, include_actions: bool = True, max_actions: int = 30) -> dict:
    """One ticket with all its fields, plus its most recent actions (notes, emails, updates)."""
    c = halo()
    t = c.get(f"Tickets/{ticket_id}", {"includedetails": True}) or {}
    out = {"ticket": _scalars(t), "summary_view": _slim_ticket(t, with_details=True)}
    if include_actions:
        acts = _items(c.get("Actions", {"ticket_id": ticket_id, "count": max_actions}), "actions")
        out["actions"] = [
            {
                k: v
                for k, v in {
                    "id": a.get("id"),
                    "when": a.get("datetime") or a.get("actiondatecreated"),
                    "who": a.get("who") or a.get("agent_name"),
                    "outcome": a.get("outcome"),
                    "note": (a.get("note") or a.get("note_html") or "")[:1500],
                    "hidden_from_user": a.get("hiddenfromuser"),
                }.items()
                if v not in (None, "")
            }
            for a in acts[:max_actions]
        ]
    return out


@mcp.tool()
def new_tickets(hours: int = 24, client_id: int | None = None, open_only: bool = True, max_results: int = 50) -> list[dict]:
    """
    Tickets opened in the last `hours`, with their description text, for triage review.
    Read-only: use the output to suggest category, priority, team and next steps.
    """
    start = (datetime.now(timezone.utc) - timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%S")
    end = (datetime.now(timezone.utc) + timedelta(days=1)).date().isoformat()
    rows = _fetch_tickets(
        client_id, start, end, "dateoccurred", "open" if open_only else "all", None, max_results
    )
    c = halo()
    out = []
    for t in rows:
        # List results may omit the description; fetch it per ticket (bounded by max_results).
        if not (t.get("details") or t.get("details_html")):
            try:
                t = c.get(f"Tickets/{t['id']}", {"includedetails": True}) or t
            except Exception as e:
                log.warning("Could not load ticket %s: %s", t.get("id"), e)
        out.append(_slim_ticket(t, with_details=True))
    return out


@mcp.tool()
def ticket_summary(
    client_id: int | None = None,
    quarter: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    max_tickets: int = 3000,
) -> dict:
    """
    QBR metrics for a client over a period: volume, monthly trend, breakdowns by type,
    category, status, team, agent, site and top requesters, resolution times and SLA results.
    Pass `quarter` like '2026-Q3', or start_date/end_date (ISO). Defaults to the last 90 days.
    """
    start, end = _period(quarter, start_date, end_date)
    raw = _fetch_tickets(client_id, start, end, "dateoccurred", "all", None, max_tickets)
    # Merged duplicates (e.g. a reply that became its own ticket) are not real work items.
    merged = [t for t in raw if (t.get("merged_into_id") or 0) > 0]
    raw = [t for t in raw if not (t.get("merged_into_id") or 0) > 0]
    tickets = [_slim_ticket(t) for t in raw]

    def top(field: str, n: int = 10) -> list:
        return Counter(t.get(field) or "(none)" for t in tickets).most_common(n)

    monthly = Counter((_dt(t.get("opened")) or datetime.min).strftime("%Y-%m") for t in tickets)

    # SLA uses Halo's own verdict on each ticket (I = met, O = missed), which already allows
    # for on-hold time and working hours. Tickets marked "exclude from SLA" are skipped, and so
    # are verdicts other than I/O (e.g. X = not measured). Resolution counts closed tickets only.
    res_hours, fix_met, fix_total, resp_met, resp_total, excluded = [], 0, 0, 0, 0, 0
    sla_states: Counter = Counter()
    for t in raw:
        opened, closed = _dt(t.get("dateoccurred")), _dt(_first(t, "dateclosed", "datecleared"))
        if opened and closed and closed >= opened:
            res_hours.append((closed - opened).total_seconds() / 3600)
        for k in ("slaresponsestate", "slastate"):
            if t.get(k) not in (None, ""):
                sla_states[f"{k}={t[k]}"] += 1
        if t.get("excludefromsla"):
            excluded += 1
            continue
        resp = str(t.get("sla_first_response_state") or t.get("slaresponsestate") or "").upper()
        if resp in ("I", "O"):
            resp_total += 1
            resp_met += resp == "I"
        fix = str(t.get("slastate") or "").upper()
        if closed and fix in ("I", "O"):
            fix_total += 1
            fix_met += fix == "I"

    def pct(a: int, b: int) -> float | None:
        return round(100 * a / b, 1) if b else None

    return {
        "client_id": client_id,
        "period": {"start": start, "end": end},
        "tickets_opened": len(tickets),
        "tickets_closed": sum(1 for t in tickets if t.get("closed")),
        "truncated": len(raw) >= max_tickets,
        "monthly_opened": dict(sorted(monthly.items())),
        "by_type": top("type"),
        "by_category": top("category", 15),
        "by_status": top("status"),
        "by_team": top("team"),
        "by_agent": top("agent"),
        "by_site": top("site"),
        "top_requesters": top("user"),
        "resolution_hours": {
            "count": len(res_hours),
            "median": round(statistics.median(res_hours), 1) if res_hours else None,
            "mean": round(statistics.mean(res_hours), 1) if res_hours else None,
        },
        "sla": {
            "resolution_met_pct": pct(fix_met, fix_total),
            "resolution_measured": fix_total,
            "response_met_pct": pct(resp_met, resp_total),
            "response_measured": resp_total,
            "excluded_from_sla": excluded,
            "merged_tickets_skipped": len(merged),
            "method": "Halo SLA state per ticket (I=met, O=missed); excluded-from-SLA tickets skipped",
            "raw_sla_state_fields": dict(sla_states),
        },
    }


@mcp.tool()
def list_assets(client_id: int, search: str | None = None, include_inactive: bool = False, max_results: int = 500) -> list[dict]:
    """
    Assets (devices) for a client, with type, site, assigned user and any warranty/purchase
    fields found. Useful for lifecycle and hardware-refresh recommendations in QBRs.
    """
    rows = halo().get_paged(
        "Asset", "assets",
        {"client_id": client_id, "search": search, "includeinactive": include_inactive},
        max_results,
    )
    keep = ("id", "inventory_number", "key_field", "key_field2", "key_field3",
            "assettype_name", "site_name", "username", "inactive", "status_name")
    out = []
    for a in rows:
        row = {k: a.get(k) for k in keep if a.get(k) not in (None, "")}
        for k, v in a.items():
            if any(w in k.lower() for w in ("warranty", "purchase", "lastupdate", "age")) and v not in (None, ""):
                if not (isinstance(v, str) and v.startswith("1900-01-01")) and not isinstance(v, (dict, list)):
                    row[k] = v
        out.append(row)
    return out


@mcp.tool()
def list_contracts(client_id: int) -> list[dict]:
    """Contracts / agreements for a client (scalar fields only)."""
    return [_scalars(c) for c in _items(halo().get("ClientContract", {"client_id": client_id}), "contracts")]


@mcp.tool()
def list_licences(client_id: int, include_inactive: bool = False) -> list[dict]:
    """
    Software licences / subscriptions recorded against a client (Halo's Licences tab), e.g.
    Microsoft 365 plans synced from a CSP or Pax8 integration: name, quantity, dates, vendor.
    """
    params = {"client_id": client_id}
    if include_inactive:
        params["includeinactive"] = True
    rows = _items(halo().get("SoftwareLicence", params), "licences")
    return [_scalars(r) for r in rows if isinstance(r, dict)]


@mcp.tool()
def list_recurring_invoices(client_id: int, include_lines: bool = True) -> list[dict]:
    """
    Recurring invoices for a client with their line items (description, quantity, price).
    Billed quantities of Microsoft 365 and other subscriptions usually show up here.
    """
    c = halo()
    heads = _items(c.get("RecurringInvoice", {"client_id": client_id}), "invoices")
    out = []
    for h in heads:
        if not isinstance(h, dict):
            continue
        row = _scalars(h)
        if include_lines and h.get("id") is not None:
            try:
                full = c.get(f"RecurringInvoice/{h['id']}", {"includedetails": True}) or {}
                lines = full.get("lines") or full.get("invoice_lines") or []
                row["lines"] = [
                    {k: l.get(k) for k in ("item_code", "item_shortdescription", "item_longdescription",
                                            "description", "qty_order", "quantity", "unit_price",
                                            "net_amount", "item_id") if l.get(k) not in (None, "")}
                    for l in lines if isinstance(l, dict)
                ]
            except Exception as e:
                row["lines_error"] = str(e)
        out.append(row)
    return out


@mcp.tool()
def list_client_services(client_id: int) -> dict:
    """
    The client's Subscribed Services tab (service catalogue subscriptions, e.g. backup,
    security and support services). Tries the Halo endpoints that serve it and returns
    the first that answers, with the endpoint used.
    """
    c = halo()
    attempts = [
        ("Service", {"client_id": client_id, "subscribed_only": True}),
        ("Service", {"client_id": client_id, "showsubscribed": True}),
        ("ServiceUsers", {"client_id": client_id}),
        ("ServiceSubscription", {"client_id": client_id}),
        ("Service", {"client_id": client_id}),
    ]
    errors = []
    for path, params in attempts:
        try:
            rows = _items(c.get(path, params))
        except Exception as e:
            errors.append(f"{path} {params}: {str(e)[:120]}")
            continue
        rows = [r for r in rows if isinstance(r, dict)]
        if rows:
            return {"endpoint": path, "params": params, "count": len(rows), "services": [_scalars(r) for r in rows[:200]]}
        errors.append(f"{path} {params}: no rows")
    return {"services": [], "tried": errors}


@mcp.tool()
def api_get(path: str, params: dict | None = None, max_items: int = 50, raw: bool = False,
            keys: list[str] | None = None, max_chars: int = 20000) -> Any:
    """
    Read-only escape hatch: GET any Halo API path (e.g. 'Service', 'Service/12', 'Client/25')
    with optional query params. By default lists are trimmed to `max_items` and to scalar fields.
    raw=True returns the full JSON including nested lists/objects (truncated to `max_chars`);
    `keys` limits a dict response to those top-level keys (or, with raw=False, lists the
    top-level keys holding nested data). Use it to find where data lives. Never writes.
    """
    path = path.strip().lstrip("/")
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_/]*", path):
        raise ValueError("path must look like 'Service' or 'Service/12'")
    payload = halo().get(path, params or None)
    if isinstance(payload, dict) and keys:
        payload = {k: payload.get(k) for k in keys}
    if raw:
        text = json.dumps(payload, default=str)
        return text if len(text) <= max_chars else text[:max_chars] + f"... [truncated, {len(text)} chars]"
    if isinstance(payload, dict) and ("id" in payload or not _items(payload)):
        out = _scalars(payload)
        nested = {k: (f"list[{len(v)}]" if isinstance(v, list) else "object")
                  for k, v in payload.items() if isinstance(v, (list, dict)) and v}
        if nested:
            out["_nested_keys"] = nested
        return out
    if isinstance(payload, list) or (isinstance(payload, dict) and _items(payload)):
        rows = _items(payload) if isinstance(payload, dict) else payload
        meta = {k: v for k, v in payload.items() if not isinstance(v, (list, dict))} if isinstance(payload, dict) else {}
        return {"meta": meta, "count": len(rows), "items": [_scalars(r) if isinstance(r, dict) else r for r in rows[:max_items]]}
    return _scalars(payload) if isinstance(payload, dict) else payload


@mcp.tool()
def run_report(report_id: int, max_rows: int = 500) -> dict:
    """Run a saved Halo report by its ID and return its rows (needs the read:reporting scope)."""
    payload = halo().get(f"Report/{report_id}", {"loadreport": True}) or {}
    report = payload.get("report", payload) if isinstance(payload, dict) else {}
    rows = _items(report, "rows")
    return {"name": payload.get("name"), "row_count": len(rows), "rows": rows[:max_rows]}


# --------------------------------------------------------------------------- #
def _self_test() -> int:
    try:
        c = halo()
        c.token()
        print(f"Token OK for {c.base}" + (f" (tenant {c.tenant})" if c.tenant else ""))
        clients = c.get_paged("Client", "clients", {}, 5)
        print(f"Clients API OK - first {len(clients)}: " + ", ".join(str(x.get("name")) for x in clients))
        tickets = c.get_paged("Tickets", "tickets", {}, 1)
        print(f"Tickets API OK - sample fields: {sorted(tickets[0].keys())[:25] if tickets else 'no tickets returned'}")
        return 0
    except Exception as e:
        print(f"FAILED: {e}")
        return 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--test", action="store_true", help="check credentials and exit")
    if ap.parse_args().test:
        sys.exit(_self_test())
    mcp.run()
