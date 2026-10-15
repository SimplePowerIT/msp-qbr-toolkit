"""
Strategic Technology Review (STR) generator for MSPs using HaloPSA and N-able N-sight.

Builds, for one Halo client and one quarter:
  <Client>_<Month><Year>_Strategic_Technology_Review.docx  (the Word template, filled)
  <Client>_<Quarter>_Strategic_Technology_Review.pdf        (portrait summary)

Data comes from the existing read-only Halo and N-sight connectors in the parent
folder (their code and Windows Credential Manager secrets are reused). Anything the
systems don't record becomes a gray [placeholder]; nothing is guessed.
"""

from __future__ import annotations

import html
import importlib.util
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse
from xml.sax.saxutils import escape as xesc

log = logging.getLogger("str")
HERE = Path(__file__).resolve().parent
ROOT = HERE.parent  # ...\halo-mcp


# --------------------------------------------------------------------------- #
# Config and connector loading
# --------------------------------------------------------------------------- #
def load_config() -> dict:
    cfg_path = HERE / "config.json"
    cfg = json.loads(cfg_path.read_text(encoding="utf-8-sig")) if cfg_path.exists() else {}
    # Halo env comes from the Claude config snippet the Halo setup wrote, unless overridden.
    snip = ROOT / "claude_config_snippet.json"
    if snip.exists():
        try:
            env = json.loads(snip.read_text(encoding="utf-8-sig"))["mcpServers"]["halo"]["env"]
            for k, v in env.items():
                os.environ.setdefault(k, v)
        except Exception as e:  # noqa: BLE001
            log.warning("Could not read Halo settings from %s: %s", snip, e)
    nsnip = ROOT / "nsight-mcp" / "claude_config_snippet.json"
    if nsnip.exists():
        try:
            env = json.loads(nsnip.read_text(encoding="utf-8-sig"))["mcpServers"]["nsight"]["env"]
            for k, v in env.items():
                os.environ.setdefault(k, v)
        except Exception as e:  # noqa: BLE001
            log.warning("Could not read N-sight settings from %s: %s", nsnip, e)
    for k, v in (cfg.get("env") or {}).items():
        os.environ[k] = str(v)
    _apply_branding(cfg)
    cfg.setdefault("template", str(ROOT / "templates" / "Client_MonthDATE_Strategic_Technology_Review_Template.docx"))
    cfg.setdefault("output_dir", str(HERE / "output"))
    return cfg


# Branding: set in config.json under "branding" (see config.example.json).
BRAND: dict[str, str] = {"company": "Your MSP", "helpdesk_url": "", "review_url": "", "phone": "", "email": "", "logo": ""}


def _apply_branding(cfg: dict) -> None:
    b = cfg.get("branding") or {}
    BRAND.update({k: str(v) for k, v in b.items() if k in BRAND and v is not None})
    LOCAL_ADMIN_USERS.update(u.lower() for u in (b.get("local_admin_users") or []))
    svc = cfg.get("services") or {}
    for key, target in (("plan", PLAN_SERVICES), ("backup", BACKUP_SERVICES), ("other", OTHER_SERVICES)):
        if key in svc:
            target.clear()
            target.update({int(k): str(v) for k, v in (svc[key] or {}).items()})


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)  # type: ignore[union-attr]
    return mod


_halo = None
_nsight = None


def halo_mod():
    global _halo
    if _halo is None:
        _halo = _load_module("halo_srv", ROOT / "server.py")
    return _halo


def nsight_mod():
    global _nsight
    if _nsight is None:
        _nsight = _load_module("nsight_srv", ROOT / "nsight-mcp" / "server.py")
    return _nsight


# --------------------------------------------------------------------------- #
# Reference data (keep in step with the msp-qbr skill)
# --------------------------------------------------------------------------- #
MODELS = {
    "20RY": "Lenovo T490", "20N4": "Lenovo T590", "20S6": "Lenovo T15 Gen 1", "20W4": "Lenovo T15 Gen 2", "20W0": "Lenovo T14 Gen 2",
    "21AH": "Lenovo T14 Gen 3", "21ML": "Lenovo T14 Gen 5", "20XW": "Lenovo X1 Carbon Gen 9",
    "21BV": "Lenovo T16 Gen 1", "21HH": "Lenovo T16 Gen 2", "21HM": "Lenovo X1 Carbon Gen 11",
    "21MN": "Lenovo T16 Gen 3", "21KC": "Lenovo X1 Carbon Gen 12", "21KS": "Lenovo P16s Gen 3",
}
# Exact processor release months, then generation fallbacks.
CPU_RELEASE = {
    "i7-8565U": "2018-08", "i7-10510U": "2019-08", "i7-1165G7": "2020-09", "i5-1145G7": "2020-09",
    "i5-11500": "2021-03", "i7-1260P": "2022-02", "i5-1235U": "2022-02", "i5-12500": "2022-01",
    "i7-12700": "2022-01", "i7-1355U": "2023-01", "i5-13500": "2023-01",
    "Ultra 5 125U": "2023-12", "Ultra 7 155U": "2023-12", "Ultra 7 155H": "2023-12", "Ultra 7 258V": "2024-09",
}
GEN_RELEASE = {"6": "2015-09", "7": "2016-08", "8": "2017-10", "9": "2018-10", "10": "2019-08", "11": "2020-09",
               "12": "2022-01", "13": "2023-01", "14": "2023-10"}
ULTRA_SERIES_RELEASE = {"1": "2023-12", "2": "2024-09", "3": "2026-01"}

UNSUPPORTED_SOFTWARE = {"acrobat pro 2020": "out of support since Nov 30, 2025",
                        "acrobat standard 2020": "out of support since Nov 30, 2025",
                        "acrobat 2020": "out of support since Nov 30, 2025"}

M365_SKUS = {"SPB": "Business Premium", "O365_BUSINESS_PREMIUM": "Business Standard",
             "O365_BUSINESS_ESSENTIALS": "Business Basic", "O365_BUSINESS": "365 Apps for Business",
             "SMB_BUSINESS": "365 Apps for Business"}
# Licensing page: four plan cards (base order) and how to find each plan's own distributor (renewal) record
PLAN_CARDS = [("Business Premium", ("business premium",)), ("Business Standard", ("business standard",)),
              ("Business Basic", ("business basic",)), ("Microsoft 365 Copilot", ("copilot",))]
# Subscribed Halo services that are third-party products without seat counts -> name on "Other product licensing"
# Halo subscribed-service name -> product shown under "Other product licensing" (example entries; edit for your services)
OTHER_PRODUCT_SERVICES = {"8x8 Phone Service": "8x8", "8x8": "8x8"}
FREE_SKU_HINTS = ("FREE", "TRIAL", "VIRAL", "_DEV", "WINDOWS_STORE", "POWER_BI_STANDARD", "PREVIEW", "FLOW_FREE")

# Halo service IDs -> names. These are specific to each Halo instance, so they are set in
# config.json under "services" (plan / backup / other). Empty here on purpose.
PLAN_SERVICES: dict[int, str] = {}
BACKUP_SERVICES: dict[int, str] = {}
OTHER_SERVICES: dict[int, str] = {}
# Halo service categories that are not real client subscriptions
SKIP_SERVICE_CATEGORIES = {"Archive", "Users and Services"}


def firewall_text(f, multi: bool) -> str:
    return f["model"] + (f" – {f['site']}" if multi and f.get("site") else "") + (f" (warranty to {f['warranty']})" if f.get("warranty") else "")


def training_text(d) -> str:
    t = d.get("training_date")
    return f"{MONTHS[t.month - 1]} {t.day}, {t.year}" if t else "??"


def clean_lic_name(n: str) -> str:
    """Distributor names -> readable: 'NCE Power BI Pro (Non-Profit Pricing)- Annual Paid Monthly' -> 'Power BI Pro (nonprofit)'."""
    n = re.sub(r"^NCE\s+", "", n or "")
    n = re.sub(r"\s*-\s*(Annual|Monthly|Triennial|Month|Year)\b.*$", "", n, flags=re.I)
    n = re.sub(r"\((?:[^)]*\b)?non-?profit[^)]*\)", "(nonprofit)", n, flags=re.I)
    return re.sub(r"\s{2,}", " ", n).strip()


def client_since_text(d) -> str | None:
    cs = d.get("client_since")
    return f"Client since {MONTHS[cs.month - 1]} {cs.day}, {cs.year}" if cs else None


def IDENTITY_TEXT(bp) -> str:
    return ("Entra ID P1 · " if bp else "") + "Microsoft Security Defaults / Conditional Access"


def ENDPOINT_TEXT(services) -> str:
    return "Huntress MDR + Defender"


def IDENTITY_LINES(bp, services) -> list:
    """Identity protection items, one per line: Entra ID / Security Defaults, then Huntress ITDR if subscribed."""
    return [IDENTITY_TEXT(bp)] + (["Huntress ITDR"] if "Huntress ITDR" in services else [])

CATEGORY_LABELS = {"Client Request>All Other": "General requests", "All Other": "Other",
                   "Client Request>Employee add/remove/change": "Employee add / remove",
                   "EDR Threat": "Security alerts", "(none)": "Uncategorized",
                   "Internal>Scheduled Ticket": "Scheduled maintenance"}

_NUM = ["No", "One", "Two", "Three", "Four", "Five", "Six"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
          "October", "November", "December"]


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _dt(v: Any) -> datetime | None:
    if not v or not isinstance(v, str):
        return None
    try:
        d = datetime.fromisoformat(v.replace("Z", "+00:00"))
    except ValueError:
        return None
    return None if d.year < 1950 else d.replace(tzinfo=None)


def _mdy(d: datetime | None) -> str:
    return f"{d:%b} {d.day}, {d.year}" if d else "—"


def last_quarter(today: date | None = None) -> str:
    today = today or date.today()
    q = (today.month - 1) // 3  # 0..3, current quarter index
    return f"{today.year - 1}-Q4" if q == 0 else f"{today.year}-Q{q}"


def normalize_quarter(text: str | None, today: date | None = None) -> str | None:
    """Accept '2026-Q3', 'Q3 2026', '2026 Q3', 'q3', '3' -> '2026-Q3' (bare quarter = this year). Blank -> None."""
    if not text or not str(text).strip():
        return None
    t = str(text).upper().replace(" ", "").replace("-", "").replace("/", "")
    m = re.fullmatch(r"(?:(\d{4})Q([1-4])|Q([1-4])(\d{4})?|([1-4]))", t)
    if not m:
        raise ValueError(f"Quarter '{text}' not understood - use something like Q3 2026, or leave it blank")
    year = m.group(1) or m.group(4)
    qn = m.group(2) or m.group(3) or m.group(5)
    if not year:
        year = (today or date.today()).year
    return f"{year}-Q{qn}"


def parse_review_date(text: str | None) -> date | None:
    """Review meeting date from the request form. SharePoint date-only columns arrive as UTC, e.g.
    2026-10-20T06:00:00Z for Oct 20 in Mountain time; a time after noon UTC means the next local day."""
    t = (text or "").strip()
    if not t:
        return None
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}))?", t)
    try:
        if m:
            dt = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if m.group(4) and int(m.group(4)) >= 12:
                dt = dt + timedelta(days=1)
            return dt
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%b %d %Y", "%b %d, %Y", "%B %d %Y", "%B %d, %Y", "%d %b %Y", "%d %B %Y"):
            try:
                return datetime.strptime(t, fmt).date()
            except ValueError:
                pass
    except ValueError:
        return None
    return None


def current_quarter(today: date | None = None) -> str:
    today = today or date.today()
    return f"{today.year}-Q{(today.month - 1) // 3 + 1}"


def prev_quarter(qs: str) -> str:
    y, q = qs.split("-Q")
    y, q = int(y), int(q)
    return f"{y - 1}-Q4" if q == 1 else f"{y}-Q{q - 1}"


def quarter_months(qs: str) -> tuple[str, str]:
    y, q = qs.split("-Q")
    first = (int(q) - 1) * 3
    return MONTHS[first], MONTHS[first + 2]


def safe_name(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9 _.-]+", "", s).strip().replace(" ", "_") or "Client"


def cpu_short(cpu: str | None) -> tuple[str, str | None]:
    """('11th Gen i5', '2020-09') from an N-sight processor string."""
    if not cpu:
        return "[processor]", None
    c = re.sub(r"\(R\)|\(TM\)|CPU|@.*$", "", cpu).replace("  ", " ").strip()
    m = re.search(r"Ultra\s+(\d)\s+(\d)(\d\d)([A-Z]?)", c)
    if m:
        key = f"Ultra {m.group(1)} {m.group(2)}{m.group(3)}{m.group(4)}"
        return f"Core {key}", CPU_RELEASE.get(key) or ULTRA_SERIES_RELEASE.get(m.group(2))
    m = re.search(r"(i[3579])-(\d{4,5})([A-Z]*\d?)", c)
    if m:
        model = f"{m.group(1)}-{m.group(2)}{m.group(3)}"
        num = m.group(2)
        gen = num[:2] if len(num) == 5 or num[:2] in ("10", "11", "12", "13", "14") else num[0]
        label = {"1": "st", "2": "nd", "3": "rd"}.get(gen[-1] if gen not in ("11", "12", "13") else "", "th")
        return f"{gen}{label} Gen {m.group(1)}", CPU_RELEASE.get(model) or GEN_RELEASE.get(gen)
    return c, None


def age_years(release: str | None, ref: date) -> float | None:
    if not release:
        return None
    y, m = map(int, release.split("-"))
    return round((ref.year + (ref.month - 1) / 12) - (y + (m - 1) / 12), 1)


def in_service_from_name(pc_name: str | None, ref: date) -> str | None:
    """In-service month/year from the PC name. Two formats are recognised:
      - ends in YYYY-MM (e.g. PC2023-07 = July 2023; '-', '_' or '.' between)
      - ends in -MMYY  (e.g. ACME-0624 = June 2024)
    9999, impossible months and future dates = unknown (falls back to processor age)."""
    name = (pc_name or "").strip()
    cands = []
    m = re.search(r"(?<!\d)((?:19|20)\d{2})[-_.](\d{2})$", name)
    if m:
        cands.append((int(m.group(1)), int(m.group(2))))
    m = re.search(r"-(\d{2})(\d{2})$", name)
    if m:
        cands.append((2000 + int(m.group(2)), int(m.group(1))))
    for y, mm in cands:
        if 1 <= mm <= 12 and (y, mm) <= (ref.year, ref.month):
            return f"{y}-{mm:02d}"
    return None


def friendly_model(manufacturer: str | None, model: str | None) -> str:
    model = (model or "").strip()
    if (manufacturer or "").lower().startswith("dell"):
        return (model if model.lower().startswith("dell ") else f"Dell {model}") if model else "Dell"
    if model[:4] in MODELS:
        return MODELS[model[:4]]
    return f"Lenovo {model}" if (manufacturer or "").upper().startswith("LENOVO") else (model or "[model]")


def incident_title(t: dict) -> str:
    s = (t.get("summary") or "").lower()
    if "itdr" in s or "identity" in s or "@" in s:
        return "Account takeover blocked"
    if "ransom" in s:
        return "Ransomware blocked"
    return "Threat detected and contained"


def incident_text(t: dict) -> str:
    """Plain-English one-liner from a Huntress / EDR ticket summary."""
    s = t.get("summary") or ""
    who = re.search(r"[Ii]ncident on ([^\s(]+@[^\s(]+)", s)
    dev = re.search(r"[Ii]ncident on ([A-Za-z0-9_-]+)", s)
    tool = "Huntress" if "huntress" in s.lower() else "Our security monitoring"
    if who:
        mbox = who.group(1).split("@")[0] + "@"
        return f"{tool} caught suspicious sign-in activity on the {mbox} mailbox and locked the account out automatically."
    if dev:
        return f"{tool} caught a threat on {dev.group(1)} and isolated the device automatically."
    return f"{tool} flagged: {s}"


WORD_LABELS = {
    "Primary Contact:": "Primary contact:", "Review Date:": "Review date:", "Support Plan:": "Support plan:",
    "Contract Renewal Date": "Contract renewal date",
    "Current Firewall/Wi-Fi:": "Current firewall / Wi-Fi:",
    "Microsoft 365 Identity Protection :": "Microsoft 365 identity protection:",
    "Last Cybersecurity Awareness Training:": "Most recent organization cybersecurity awareness training:",
    "Endpoint Security Protection:": "Endpoint security protection:",
    "Server": "Servers and Network",
    "Server/OS:": "Server / OS:", "VM/Cloud:": "VM / cloud:", "Age/Warranty:": "Age / warranty:",
    "Server/NAS:": "Server / NAS:",
}

# Account names shown as "Unknown" in the PC user column. Add your own in config.json: branding.local_admin_users
LOCAL_ADMIN_USERS = {"administrator", "admin", "localadmin", "defaultuser0"}


def clean_user(u: str | None) -> str:
    u = (u or "").split("\\")[-1]
    if u.strip().lower() in LOCAL_ADMIN_USERS:
        return "Unknown"
    return re.sub(r"(?<=[a-z])(?=[A-Z])", " ", u) or "—"


def _find_client_in(obj: Any, client_id: int, client_name: str) -> bool:
    """Search a nested Halo record for this client (by id or name)."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            kl = k.lower()
            if kl in ("client_id", "clientid", "area_id") and str(v) == str(client_id):
                return True
            if kl in ("client_name", "clientname", "area_name") and isinstance(v, str) and v.strip() == client_name.strip():
                return True
            if isinstance(v, (dict, list)) and _find_client_in(v, client_id, client_name):
                return True
        # Service subscribers: {"type": "A", "data_id": <client id>, "data_name": "<client>"}
        if str(obj.get("type", "")).upper() == "A" and (str(obj.get("data_id")) == str(client_id)
                                                        or (obj.get("data_name") or "").strip() == client_name.strip()):
            return True
        # Stakeholder groups describe the client in the text: Client "ACME*" - Stakeholder Subscriber
        if isinstance(obj.get("desc"), str) and f'Client "{client_name.strip()}"' in obj["desc"]:
            return True
    elif isinstance(obj, list):
        return any(_find_client_in(x, client_id, client_name) for x in obj)
    return False


# --------------------------------------------------------------------------- #
# Data collection
# --------------------------------------------------------------------------- #
def collect(client_id: int, quarter: str | None = None, overrides: dict | None = None,
            main_contact: str | None = None) -> dict:
    H = halo_mod()
    c = H.halo()
    quarter = quarter or last_quarter()
    ref = date.today()
    client = H.get_client(client_id)
    name = client.get("name") or f"Client {client_id}"
    legal = client.get("accountsid") or client.get("trading_name") or name
    d: dict[str, Any] = {"client_id": client_id, "name": name, "legal": legal, "quarter": quarter,
                         "phone": client.get("main_phonenumber"), "website": client.get("website"),
                         "nsight_id": client.get("gficlientid"), "ref": ref, "warnings": []}

    # Primary contact: the client's "Important contact" in Halo
    # Primary contact + secondary contacts: the client's "Important contact" users in Halo.
    # The request's Main contact box (name or email) picks who is listed first.
    d["contact"], d["co_contacts"] = None, []
    try:
        users = [u for u in _items_raw(c, "Users", {"client_id": client_id}) if isinstance(u, dict) and not u.get("inactive")]
        def _c(u):
            return {"name": u.get("name"), "email": u.get("emailaddress"),
                    "title": u.get("other2") or u.get("jobtitle") or u.get("title")}
        vip = [u for u in users if u.get("isimportantcontact")] + [u for u in users if u.get("isimportantcontact2") and not u.get("isimportantcontact")]
        people = [_c(u) for u in vip]
        want = (main_contact or "").strip().lower()
        if want:
            def _hit(p):
                return want in (p["name"] or "").lower() or want == (p["email"] or "").lower()
            pick = next((p for p in people if _hit(p)), None)
            if not pick:
                pick = next((_c(u) for u in users if want in (u.get("name") or "").lower()
                             or want == (u.get("emailaddress") or "").lower()), None)
            if pick:
                people = [pick] + [p for p in people if p["email"] != pick["email"] or p["name"] != pick["name"]]
            else:
                d["warnings"].append(f"Main contact '{main_contact}' not found in Halo; used the Important contact")
        if people:
            d["contact"], d["co_contacts"] = people[0], people[1:3]
    except Exception as ex:  # noqa: BLE001
        log.info("Contact lookup failed: %s", ex)

    # Tickets
    d["tq"] = H.ticket_summary(client_id=client_id, quarter=quarter)
    d["tp"] = H.ticket_summary(client_id=client_id, quarter=prev_quarter(quarter))
    s, e = d["tq"]["period"]["start"], d["tq"]["period"]["end"]
    d["tickets"] = H.search_tickets(client_id=client_id, start_date=s, end_date=e, max_results=500)
    d["open"] = H.search_tickets(client_id=client_id, status="open", max_results=200)

    # Assets, licences
    d["assets"] = H.list_assets(client_id)
    d["licences"] = [r for r in _items_raw(c, "SoftwareLicence", {"client_id": client_id}) if isinstance(r, dict)]
    try:
        d["contracts"] = H.list_contracts(client_id)
    except Exception:  # noqa: BLE001
        d["contracts"] = []

    # Subscribed services: read them from the client record (Halo "Subscribed Services" tab).
    d["services"] = set()
    d["client_since"] = None
    d["wifi"] = None
    d["training_date"] = None
    try:
        crec = c.get(f"Client/{client_id}", {"includedetails": True}) or {}
        # "Client Since" custom field (CFClientSince) on the client overview
        for cf in crec.get("customfields") or []:
            # "Wi-Fi" custom field (any field labelled Wi-Fi / WiFi / Wireless), e.g. "UniFi"
            _lbl = re.sub(r"[^a-z]", "", (cf.get("label") or cf.get("name") or "").lower())
            if _lbl in ("wifi", "cfwifi", "wireless", "cfwireless") and (cf.get("display") or cf.get("value")):
                d["wifi"] = str(cf.get("display") or cf.get("value")).strip()
            # Last cybersecurity awareness training date (any client field labelled ...Training / ...Awareness)
            # Cybersecurity Awareness Training tab: CFLastCyberTraining ("Last Training Date"). Never the Next date.
            _nm = (cf.get("name") or "").lower()
            _is_last_training = _nm == "cflastcybertraining" or (("training" in _lbl or "awareness" in _lbl) and "next" not in _lbl and "next" not in _nm)
            if _is_last_training and (cf.get("display") or cf.get("value")):
                _disp = (cf.get("display") or "").strip()
                try:
                    d["training_date"] = datetime.strptime(_disp, "%m/%d/%Y").date()
                except ValueError:
                    _v = _dt(cf.get("value"))
                    d["training_date"] = _v.date() if _v else None
            if (cf.get("name") or "").lower() == "cfclientsince" or (cf.get("label") or "").strip().lower() == "client since":
                disp = (cf.get("display") or "").strip()
                try:
                    d["client_since"] = datetime.strptime(disp, "%m/%d/%Y").date() if disp else None
                except ValueError:
                    d["client_since"] = (_dt(cf.get("value")) or None) and _dt(cf.get("value")).date()
        for g in crec.get("group_service_subscriptions") or []:
            nm = (g.get("entity_name") or "").strip()
            if nm:
                d["services"].add(nm)
    except Exception as ex:  # noqa: BLE001
        log.info("Client subscription lookup failed: %s", ex)
    # Fallback / cross-check: look for this client inside each known service record.
    for sid, sname in ({} if d["services"] else {**PLAN_SERVICES, **BACKUP_SERVICES, **OTHER_SERVICES}).items():
        try:
            rec = c.get(f"Service/{sid}", {"includedetails": True})
            if _find_client_in({k: v for k, v in (rec or {}).items() if isinstance(v, (list, dict))}, client_id, name):
                d["services"].add(sname)
        except Exception as ex:  # noqa: BLE001
            log.info("Service %s lookup failed: %s", sid, ex)
    for sname in (overrides or {}).get("services", []):
        d["services"].add(sname)
    _svc_cat = {}
    try:
        for sv in _items_raw(c, "Service", {}):
            if isinstance(sv, dict) and sv.get("name"):
                _svc_cat[sv["name"].strip()] = sv.get("service_category_name") or ""
    except Exception:  # noqa: BLE001
        pass
    d["services"] = {x for x in d["services"] if _svc_cat.get(x, "") not in SKIP_SERVICE_CATEGORIES}
    # Backup check for the Unifi Wi-Fi marker service: look for this client in the service's own subscribers
    if not any(x.strip().lower() == "unifi" for x in d["services"]):
        try:
            for sv in _items_raw(c, "Service", {}):
                if isinstance(sv, dict) and (sv.get("name") or "").strip().lower() == "unifi":
                    rec = c.get(f"Service/{sv['id']}", {"includedetails": True}) or {}
                    if _find_client_in({k: v for k, v in rec.items() if isinstance(v, (list, dict))}, client_id, name):
                        d["services"].add("Unifi")
        except Exception as ex:  # noqa: BLE001
            log.info("Unifi service check failed: %s", ex)
    log.info("Subscribed services for %s: %s", name, ", ".join(sorted(d["services"])) or "none")

    # N-sight workstations and servers
    d["pcs"], d["nsight_servers"] = [], []
    if d["nsight_id"]:
        try:
            N = nsight_mod()
            off = 0
            while off is not None:
                page = N.hardware_inventory(int(d["nsight_id"]), offset=off, limit=25)
                d["pcs"].extend(page["rows"])
                off = page["next_offset"]
            d["nsight_servers"] = N.list_servers(int(d["nsight_id"]))
        except Exception as ex:  # noqa: BLE001
            d["warnings"].append(f"N-sight lookup failed: {ex}")
    else:
        d["warnings"].append("No N-sight client id on the Halo client record (gficlientid).")
    return d


def _items_raw(c, path: str, params: dict) -> list:
    payload = c.get(path, params)
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for v in payload.values():
            if isinstance(v, list):
                return v
    return []


# --------------------------------------------------------------------------- #
# Analysis -> plain values used by both the Word document and the PDF
# --------------------------------------------------------------------------- #
def analyse(d: dict) -> dict:
    now = datetime.now()
    a: dict[str, Any] = {}
    month_from, month_to = quarter_months(d["quarter"])
    a["period_label"] = f"{month_from}–{month_to} {d['quarter'][:4]}"
    _last_day = {"March": 31, "June": 30, "September": 30, "December": 31}[month_to]
    a["period_full"] = f"{month_from} 1 – {month_to} {_last_day}, {d['quarter'][:4]}"
    rd = d.get("review_date")
    a["review_month"] = f"{MONTHS[(rd or d['ref']).month - 1]}"
    a["year"] = (rd or d["ref"]).year
    a["review_day"] = rd.day if rd else None
    a["plan"] = next((p for p in PLAN_SERVICES.values() if p in d["services"]), None)

    # Licences
    lic = d["licences"]
    tenant = [l for l in lic if l.get("type") == 0 and not l.get("deleted")]
    dist = [l for l in lic if l.get("type") == 1 and not l.get("deleted")]
    m365 = {v: None for v in dict.fromkeys(M365_SKUS.values())}
    for l in tenant:
        label = M365_SKUS.get(str(l.get("product_sku")))
        if label and (l.get("count") or 0) > 0:
            m365[label] = (int(l.get("count") or 0), int(l.get("consumedcount") or 0))
    a["m365"] = m365
    nonprofit = any("nonprofit" in (l.get("name") or "").lower() and "premium" in (l.get("name") or "").lower() for l in dist)
    a["bp_nonprofit"] = nonprofit
    future_ms = [_dt(l.get("end_date")) for l in dist
                 if "microsoft 365" in (l.get("name") or "").lower() and l.get("billing_cycle") in ("year", "month")]
    future_ms = [x for x in future_ms if x and now < x < datetime(2090, 1, 1)]
    a["renewal"] = min(future_ms) if future_ms else None
    a["stale_dist"] = [l.get("name") for l in dist if (_dt(l.get("end_date")) or datetime.max) < now]
    has_bp = bool(m365.get("Business Premium"))
    defender_lic = next((l for l in tenant if l.get("product_sku") == "ATP_ENTERPRISE" and (l.get("count") or 0) > 0), None)
    a["defender"] = "Plan 1 – included in Business Premium" if has_bp else (
        f"Plan 1 – {defender_lic['count']} licences" if defender_lic else None)
    others = []
    for l in tenant:
        sku = str(l.get("product_sku") or "")
        cnt = int(l.get("count") or 0)
        if sku in M365_SKUS or sku == "ATP_ENTERPRISE" or cnt == 0 or cnt >= 10000 or any(h in sku for h in FREE_SKU_HINTS):
            continue
        ren = next((_dt(x.get("end_date")) for x in dist if _name_core(x.get("name")) == _name_core(l.get("name"))
                    and (_dt(x.get("end_date")) or datetime.min) > now), None)
        others.append({"name": l.get("name"), "seats": cnt, "used": l.get("consumedcount"),
                       "renews": _mdy(ren) if ren and ren.year < 2090 else "—"})
    for l in dist:
        nm = (l.get("name") or "")
        end = _dt(l.get("end_date"))
        if "microsoft" in nm.lower() or not end or end < now:
            continue
        _c = l.get("count")
        try:
            _c = int(float(_c)) if _c not in (None, "") else _c
        except (TypeError, ValueError):
            pass
        row = {"name": clean_lic_name(nm.split("--")[0].strip()), "seats": _c, "used": None,
               "renews": "Perpetual" if end.year >= 2090 else _mdy(end)}
        # Same product already listed from the Microsoft tenant (e.g. "Power BI Pro" vs
        # "NCE Power BI Pro (Non-Profit Pricing)"): keep this active distributor row, take "in use" from the tenant
        _key = lambda t: re.sub(r"[^a-z0-9]", "", (t or "").lower())
        dup = next((o for o in others if o.get("renews") == "—" and _key(o["name"]) and _key(o["name"]) in _key(row["name"])), None)
        if dup:
            row["used"] = dup.get("used")
            others.remove(dup)
        others.append(row)
    # Unsupported software from Halo software assets
    sw = Counter((x.get("key_field") or x.get("inventory_number") or "").strip()
                 for x in d["assets"] if (x.get("assettype_name") or "").lower() == "software")
    a["unsupported"] = []
    for nm, n in sw.items():
        why = UNSUPPORTED_SOFTWARE.get(nm.lower())
        if why:
            a["unsupported"].append({"name": nm, "count": n, "why": why})
            others.append({"name": f"Adobe {nm}" if "acrobat" in nm.lower() else nm, "seats": n, "used": n,
                           "renews": "Unsupported", "why": why})
    a["other_lic"] = others

    # --- Licensing page (Oct 2026 layout) ---
    def _plan_renewal(keys):
        ends = [_dt(x.get("end_date")) for x in dist
                if any(k in (x.get("name") or "").lower() for k in keys)]
        ends = [e for e in ends if e and now < e < datetime(2090, 1, 1)]
        return min(ends) if ends else None
    copilot_t = [l for l in tenant if "copilot" in (l.get("name") or "").lower()
                 and 0 < int(l.get("count") or 0) < 10000 and "chat" not in (l.get("name") or "").lower()]
    plan_seats = dict(m365)
    plan_seats["Microsoft 365 Copilot"] = ((sum(int(l.get("count") or 0) for l in copilot_t),
                                            sum(int(l.get("consumedcount") or 0) for l in copilot_t)) if copilot_t else None)
    cards = []
    for label, keys in PLAN_CARDS:
        v = plan_seats.get(label)
        ren = _plan_renewal(keys) if v else None
        cards.append({"label": label, "seats": v, "renews": ren,
                      "soon": bool(ren and (ren - now).days <= 90),
                      "nonprofit": bool(v and label == "Business Premium" and a["bp_nonprofit"])})
    a["plan_cards"] = [c for c in cards if c["seats"]] + [c for c in cards if not c["seats"]]
    # "Also licensed" display rows: Defender, 365 Apps (if seats), then other seat-based subscriptions (Copilot has its own card)
    also = []
    bp_v = m365.get("Business Premium")
    if bp_v:
        also.append({"name": "Defender for Office 365 Plan 1", "seats": bp_v[0], "used": bp_v[1], "renews": "With Business Premium"})
    elif defender_lic:
        dren = _plan_renewal(("defender",))
        also.append({"name": "Defender for Office 365 Plan 1", "seats": int(defender_lic.get("count") or 0),
                     "used": defender_lic.get("consumedcount"), "renews": _mdy(dren) if dren else "—"})
    apps = m365.get("365 Apps for Business")
    if apps:
        aren = _plan_renewal(("apps for business",))
        also.append({"name": "Microsoft 365 Apps for Business", "seats": apps[0], "used": apps[1], "renews": _mdy(aren) if aren else "—"})
    also += [o for o in others if "copilot" not in (o.get("name") or "").lower()]
    a["also_lic"] = also
    # "Other product licensing": third-party products without seat counts (licence records + subscribed services)
    prods = []
    _ms = re.compile(r"(?i)microsoft|office|power ?bi|exchange|teams|copilot|defender|entra|intune|visio|project|windows|azure|dynamics|fabric|onedrive|sharepoint")
    for l in lic:
        if l.get("type") == 0 or l.get("deleted"):
            continue
        nm = (l.get("name") or "").split("--")[0].strip()
        end = _dt(l.get("end_date"))
        if not nm or _ms.search(nm) or (end and end < now):
            continue
        try:
            cnt = float(l.get("count") or 0)
        except (TypeError, ValueError):
            cnt = 0
        if cnt <= 0:
            prods.append(nm)
    for sv in d.get("services") or ():
        nm = OTHER_PRODUCT_SERVICES.get(sv.strip())
        if nm:
            prods.append(nm)
    a["other_products"] = list(dict.fromkeys(prods))

    # Assets
    assets = d["assets"]
    # Every firewall / router in Halo, one per site
    a["firewalls"] = []
    for fw in [x for x in assets if (x.get("assettype_name") or "").lower() in ("router", "firewall") and not x.get("inactive")]:
        raw = " ".join(str(fw.get(k) or "") for k in ("key_field", "inventory_number"))
        m = re.search(r"(?i)\b(TZ\s?\d+\w*|NSa\s?\d+\w*|SOHO\s?\w+)\b", raw)
        model = ("SonicWall " + m.group(1).upper().replace(" ", "")) if m else ((fw.get("key_field") or fw.get("inventory_number") or "").strip() or "[model]")
        we = _dt(fw.get("warranty_end"))
        a["firewalls"].append({"model": model, "site": (fw.get("site_name") or "").strip(),
                               "warranty": f"{we:%b %Y}" if we else None})
    a["firewall"] = a["firewalls"][0] if a["firewalls"] else None
    a["printers"] = [x.get("inventory_number") for x in assets if (x.get("assettype_name") or "").lower() == "printer"]
    servers = []
    for x in assets:
        if (x.get("assettype_name") or "").lower() != "server":
            continue
        kf = (x.get("key_field") or "")
        ws, we = _dt(x.get("warranty_start")), _dt(x.get("warranty_end"))
        is_vm = "virtual" in kf.lower()
        servers.append({"name": x.get("inventory_number"), "vm": "VM" if is_vm else f"Dell {kf}".replace("Dell Dell", "Dell") if kf else "[model]",
                        "age": "Runs on host hardware" if is_vm else (f"{ws:%Y} / to {we:%b %Y}" if ws and we else "[age / warranty]"),
                        "obs": None if is_vm else ("Current" if we and we > now else ("Out of warranty" if we else None))})
    a["servers"] = servers

    # Backup (rules from the skill)
    svc = d["services"]
    a["backup_server"] = next((s for s in ("Cove Server Premium", "Cove Server Standard", "Cove Server Basic") if s in svc), None) \
        or ("N/A – no on-site server" if not servers else None)
    a["backup_m365"] = "Cove M365 Backup" if "Cove M365 Backup" in svc else None
    a["backup_pc"] = "Cove PC All-In" if "Cove PC All-In" in svc else "Cove Documents"

    # Wi-Fi: the client's Wi-Fi custom field if set, otherwise "UniFi" when subscribed to the Unifi service
    if not d.get("wifi") and any(x.strip().lower() == "unifi" for x in svc):
        d["wifi"] = "UniFi"
    a["services_list"] = sorted((x for x in svc if x.strip().lower() != "unifi"),
                                key=lambda x: (x not in PLAN_SERVICES.values(), x.lower()))
    a["endpoint"] = ENDPOINT_TEXT(svc)
    # Security Awareness Training: "Yes" when a SAT service (Huntress SAT) is subscribed in Halo
    a["sat"] = "Yes" if any(x == "Huntress SAT" or "awareness" in x.lower() or x.upper().endswith(" SAT") for x in svc) else "No"
    a["identity"] = IDENTITY_LINES(bool(m365.get("Business Premium")), svc)

    # LastPass
    lp_open = [t for t in d["open"] if "lastpass" in (t.get("summary") or "").lower()]
    a["lastpass"] = ("In progress (" + ", ".join(f"#{t['id']}" for t in lp_open) + ")") if lp_open else (
        "In use" if "LastPass" in svc else None)

    # Tickets
    tq, tp = d["tq"], d["tp"]
    a["t_resolved"], a["t_opened"] = tq["tickets_closed"], tq["tickets_opened"]
    a["t_resolved_prev"] = tp["tickets_closed"]
    a["sla_resp"], a["sla_res"] = tq["sla"]["response_met_pct"], tq["sla"]["resolution_met_pct"]
    a["sla_resp_prev"], a["sla_res_prev"] = tp["sla"]["response_met_pct"], tp["sla"]["resolution_met_pct"]
    a["sla_measured"] = tq["sla"]["resolution_measured"]
    a["monthly"] = tq["monthly_opened"]
    cats = []
    for cat, n in tq["by_category"]:
        label = CATEGORY_LABELS.get(cat) or str(cat).split(">")[-1]
        cats.append((label, n))
    a["categories"] = cats
    def _is_threat(t):
        return (t.get("type") or "").strip().lower() == "threat alert" or "threat" in (t.get("category") or "").lower()
    seen, threats = set(), []
    for t in list(d["open"]) + list(d["tickets"]):
        if _is_threat(t) and t.get("id") not in seen:
            seen.add(t.get("id"))
            threats.append(t)
    threats.sort(key=lambda t: t.get("opened") or "", reverse=True)
    a["incidents"] = threats
    a["incidents_open"] = [t for t in threats if (t.get("status") or "").lower() not in ("closed", "resolved", "cancelled")]

    # PCs
    rows = []
    stale_cut = now.timestamp() - 30 * 86400
    for p in d["pcs"]:
        cpu_label, rel = cpu_short(p.get("processor"))
        # Age: from the in-service date in the PC name when there is one, otherwise from the processor release
        svc_date = in_service_from_name(p.get("pc_name"), d["ref"])
        age = age_years(svc_date or rel, d["ref"])
        age_src = "service" if svc_date else "cpu"
        ram = p.get("ram")
        ram_gb = f"{round(int(ram) / 2**30)} GB" if str(ram or "").isdigit() else "[ram]"
        scan = _dt((p.get("scan_time") or "").replace(" ", "T"))
        rows.append({"model": friendly_model(p.get("manufacturer"), p.get("model")), "name": p.get("pc_name"),
                     "user": clean_user(p.get("last_user")), "cpu": cpu_label, "ram": ram_gb, "age": age,
                     "w10": "windows 10" in (p.get("os") or "").lower(),
                     "stale": bool(scan and scan.timestamp() < stale_cut),
                     "site": (p.get("site") or "").strip() or "Unassigned site", "age_src": age_src})
    rows.sort(key=lambda r: (-(r["age"] if r["age"] is not None else -1), r["name"] or ""))
    site_count = {}
    for r in rows:
        site_count[r["site"]] = site_count.get(r["site"], 0) + 1
    # More than one N-sight site: group by site (largest site first), still oldest to newest within each
    a["pc_sites"] = sorted(site_count, key=lambda x: (-site_count[x], x.lower())) if len(site_count) > 1 else []
    a["pc_site_count"] = site_count
    if a["pc_sites"]:
        order = {x: i for i, x in enumerate(a["pc_sites"])}
        rows.sort(key=lambda r: order[r["site"]])  # stable: keeps oldest-first inside each site
    a["pcs"] = rows
    _rank = {x: i for i, x in enumerate(sorted(site_count, key=lambda x: (-site_count[x], x.lower())))}
    a["firewalls"].sort(key=lambda f: (_rank.get(f["site"], 99), f["site"].lower()))
    a["firewall"] = a["firewalls"][0] if a["firewalls"] else None
    a["halo_ws"] = sum(1 for x in assets if (x.get("assettype_name") or "").lower() == "workstation")
    a["pc_now"] = [r for r in rows if r["age"] is not None and r["age"] >= 5]
    a["pc_soon"] = [r for r in rows if r["age"] is not None and 4 <= r["age"] < 5]

    # Recommendations (rule-based, max 5)
    rec = []
    for t in a["incidents_open"][:1]:
        rec.append(f"Finish the security incident clean-up (#{t['id']}) and review Microsoft Security Defaults / Conditional Access sign-ins")
    if threats and not a["incidents_open"]:
        rec.append("Follow up on this quarter's security incident: review Microsoft Security Defaults / Conditional Access sign-ins and schedule awareness training")
    if lp_open:
        rec.append("Roll out LastPass and use it to rotate shared passwords")
    def _names(rs):
        return ", ".join(r["name"] for r in rs[:3]) + (f" and {len(rs) - 3} more" if len(rs) > 3 else "")
    nxt = d["ref"].year + 1
    if a["pc_now"] and a["pc_soon"]:
        rec.append(f"Replace {_names(a['pc_now'])} now; budget {_names(a['pc_soon'])} for {nxt}")
    elif a["pc_now"]:
        rec.append(f"Replace {_names(a['pc_now'])} now")
    elif a["pc_soon"]:
        rec.append(f"Budget to replace {_names(a['pc_soon'])} in {nxt}")
    for u in a["unsupported"]:
        rec.append(f"Replace the {u['count']} unsupported {u['name']} licence{'s' if u['count'] > 1 else ''}")
    bp = m365.get("Business Premium")
    if bp and bp[1] >= bp[0]:
        rec.append(f"Add a Business Premium seat before the next hire ({bp[1]} of {bp[0]} in use)")
    a["recs"] = rec[:5]

    # Data gaps for the run log
    gaps = []
    if a["halo_ws"] != len(rows):
        gaps.append(f"Halo has {a['halo_ws']} workstations; N-sight has {len(rows)}")
    if not a["plan"]:
        gaps.append("Support plan not found in Halo subscribed services")
    if not a["backup_m365"]:
        gaps.append("No Microsoft 365 backup service found in Halo")
    if a["stale_dist"]:
        gaps.append(f"{len(a['stale_dist'])} expired distributor subscription records in Halo")
    a["gaps"] = gaps + a.get("lic_gaps", []) + d["warnings"]
    return a


def _name_core(n: Any) -> str:
    n = str(n or "").lower()
    n = re.sub(r"\b(nce|annual|monthly|commit|paid|\(.*?\)|-)\b", " ", n)
    return re.sub(r"\s+", " ", n).strip()


# --------------------------------------------------------------------------- #
# Word document
# --------------------------------------------------------------------------- #
def _B(t): return f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">{xesc(str(t))}</w:t></w:r>'
def _PH(t): return f'<w:r><w:rPr><w:b/><w:bCs/><w:highlight w:val="lightGray"/></w:rPr><w:t xml:space="preserve">{xesc(str(t))}</w:t></w:r>'
def _N(t): return f'<w:r><w:t xml:space="preserve">{xesc(str(t))}</w:t></w:r>'
def _V(v, ph): return _B(v) if v not in (None, "") else _PH(ph)


class _Doc:
    def __init__(self, xml: str):
        self.x = xml

    def fill(self, pid: str, runs: str):
        m = re.search(r'<w:p w14:paraId="%s"[^>]*?/>' % pid, self.x)
        if m:
            self.x = self.x[:m.start()] + m.group(0)[:-2] + ">" + runs + "</w:p>" + self.x[m.end():]
            return
        m = re.search(r'(<w:p w14:paraId="%s"[^>]*>(?:<w:pPr>.*?</w:pPr>)?)(.*?)</w:p>' % pid, self.x, re.S)
        if not m:
            raise ValueError(f"Template paragraph {pid} not found")
        self.x = self.x[:m.start()] + m.group(1) + m.group(2) + runs + "</w:p>" + self.x[m.end():]

    def set_text(self, pid: str, runs: str):
        m = re.search(r'(<w:p w14:paraId="%s"[^>]*>(?:<w:pPr>.*?</w:pPr>)?)(.*?)</w:p>' % pid, self.x, re.S)
        if not m:
            raise ValueError(f"Template paragraph {pid} not found")
        self.x = self.x[:m.start()] + m.group(1) + runs + "</w:p>" + self.x[m.end():]

    def after(self, pid: str, xml: str):
        i = self.x.index('w14:paraId="%s"' % pid)
        j = self.x.index("</w:p>", i) + 6
        self.x = self.x[:j] + xml + self.x[j:]

    def after_table_of(self, pid: str, xml: str):
        i = self.x.index('w14:paraId="%s"' % pid)
        j = self.x.index("</w:tbl>", i) + 8
        self.x = self.x[:j] + xml + self.x[j:]

    def rows_into_table_of(self, pid: str, xml: str):
        i = self.x.index('w14:paraId="%s"' % pid)
        j = self.x.index("</w:tbl>", i)
        self.x = self.x[:j] + xml + self.x[j:]

    def relabel(self, old: str, new: str) -> bool:
        """Replace the text of the template paragraph whose whole text is `old` (keeps the first run's formatting)."""
        for m in re.finditer(r'<w:p[ >](?:(?!</w:p>).)*</w:p>', self.x, re.S):
            p = m.group(0)
            if "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)).strip() != old:
                continue
            ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
            rpr = re.search(r"<w:r>\s*(<w:rPr>.*?</w:rPr>)", p, re.S) or re.search(r"<w:r [^>]*>\s*(<w:rPr>.*?</w:rPr>)", p, re.S)
            start = re.match(r"<w:p[^>]*>", p).group(0)
            newp = start + (ppr.group(0) if ppr else "") + "<w:r>" + (rpr.group(1) if rpr else "") + \
                f'<w:t xml:space="preserve">{xesc(new)}</w:t></w:r></w:p>'
            self.x = self.x[:m.start()] + newp + self.x[m.end():]
            return True
        return False

    def remove_row_of(self, pid: str):
        i = self.x.find('w14:paraId="%s"' % pid)
        if i < 0:
            return
        s0 = max(self.x.rfind("<w:tr>", 0, i), self.x.rfind("<w:tr ", 0, i))
        e0 = self.x.find("</w:tr>", i) + 7
        self.x = self.x[:s0] + self.x[e0:]

    def drop_empty(self, pid: str):
        self.x = re.sub(r'<w:p w14:paraId="%s"[^>]*/>' % pid, "", self.x)


def build_docx(d: dict, a: dict, out_path: Path, template: Path) -> Path:
    with zipfile.ZipFile(template) as z:
        files = {n: z.read(n) for n in z.namelist()}
    D = _Doc(files["word/document.xml"].decode("utf-8"))
    P = lambda runs: '<w:p><w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>' + runs + "</w:p>"  # noqa: E731
    seats = lambda v: f"{v[0]} ({v[1]} assigned)" if v else "None"  # noqa: E731

    # Header
    D.fill("04868F2A", _B(f"{d['name'].rstrip('*')} ({d['legal']})" if d["legal"] != d["name"] else d["name"]))
    if client_since_text(d):
        D.after("04868F2A", P(_N(client_since_text(d))))
    ct = d.get("contact") or {}
    D.fill("392EFF85", _V(", ".join(x for x in (ct.get("name"), ct.get("title")) if x) or None, "[Name]"))
    D.fill("5B0D732C", _V(d.get("phone"), "[phone]"))
    D.fill("603432BB", _B(f"{a['review_month']} {a['review_day']}, {a['year']}") if a.get("review_day")
           else _B(f"{a['review_month']} ") + _PH("[day]") + _B(f", {a['year']}"))
    D.fill("4D8B1FF3", _V(ct.get("email"), "[Email]"))
    for cc in reversed(d.get("co_contacts") or []):
        D.after("392EFF85", P(_N("Secondary contact: ") + _B(", ".join(x for x in (cc.get("name"), cc.get("title")) if x))))
        if cc.get("email"):
            D.after("4D8B1FF3", P(_N("Secondary contact: ") + _B(cc["email"])))
    D.fill("6D559C52", _V(a["plan"], "[support plan]"))
    # Microsoft 365 licensing
    bp = a["m365"]["Business Premium"]
    # Four plan boxes, same order as the slides (licensed first): label, seats line, own renewal line
    slots = [("12DA27B4", "5631E714"), ("537AC291", "2787242C"), ("7A8FF0EA", None), ("7F1A2EC6", None)]
    RN = lambda runs: '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' + runs + "</w:p>"  # noqa: E731
    for (lab_pid, val_pid), c in zip(slots, a["plan_cards"]):
        v = c["seats"]
        D.set_text(lab_pid, _N(c["label"]))
        line1 = (f"{v[0]} ({v[1]} assigned)" + (" – nonprofit" if c["nonprofit"] else "")) if v else "None"
        if not v:
            line2 = _N("—")
        elif c["renews"]:
            line2 = (f'<w:r><w:rPr><w:b/><w:bCs/><w:color w:val="B7791F"/></w:rPr><w:t xml:space="preserve">Renews {xesc(_mdy(c["renews"]))}</w:t></w:r>'
                     if c["soon"] else _N(f"Renews {_mdy(c['renews'])}"))
        else:
            line2 = _N("Free nonprofit grant")
        if val_pid:
            D.fill(val_pid, _B(line1))
            D.after(val_pid, RN(line2))
        else:
            D.after(lab_pid, RN(line2))
            D.after(lab_pid, P(_B(line1)))
    _word_also = list(a["also_lic"])
    if a["other_products"]:
        D.after_table_of("12DA27B4", '<w:p><w:pPr><w:spacing w:before="60" w:after="0"/></w:pPr>' +
                         _N("Other product licensing: ") + _N("; ".join(f"{p} (subscribed)" for p in a["other_products"]) + ".") + "</w:p>")
    if _word_also:
        parts = []
        for o in _word_also:
            s = f"{o['name']} ×{o['seats']}"
            if o["used"] is not None:
                s += f" ({o['used']} assigned)"
            if o["renews"] == "Unsupported":
                # Unsupported software: yellow highlight, same as obsolete PCs.
                s += " – " + (o.get("why") or "unsupported")
                parts.append(f'<w:r><w:rPr><w:b/><w:bCs/><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">{xesc(s)}</w:t></w:r>')
                continue
            if o["renews"] == "Perpetual":
                s += " – perpetual"
            elif o["renews"] not in ("—", None):
                s += (" – included with Business Premium" if str(o["renews"]).startswith("With ") else " – renews " + o["renews"])
            parts.append(_N(s))
        D.after_table_of("12DA27B4", '<w:p><w:pPr><w:spacing w:before="120" w:after="0"/></w:pPr>' +
                         _N("Also licensed: ") + _N(" · ").join(parts) + _N(".") + "</w:p>")
    # Old "Contract Renewal Date | Defender for Office 365" row: renewals now live in each plan box, Defender in "Also licensed"
    D.remove_row_of("4139BD1D")
    # Security
    fw = a["firewall"]
    fws = a["firewalls"]
    D.fill("65BB80AD", (_B("; ".join(firewall_text(f, len(fws) > 1) for f in fws)) if fws else _PH("[firewall]"))
           + _B("; Wi-Fi ") + (_B(d["wifi"]) if d.get("wifi") else _PH("[model]")))
    D.fill("0BAE2DB1", _B("Yes – Plan 1 (Business Premium)") if bp else _V(a["defender"], "[Yes/No]"))
    D.fill("720790D1", _B(a["identity"][0]))
    for extra in reversed(a["identity"][1:]):
        D.after("720790D1", P(_B(extra)))
    if a["lastpass"]:
        D.fill("17A6AB0E", _B(a["lastpass"]))
    D.fill("7967A832", _B(training_text(d)))
    D.set_text("29E8A125", _B(a["endpoint"]))
    D.rows_into_table_of("29E8A125", '<w:tr><w:tc><w:tcPr><w:tcW w:w="9350" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>'
                         '<w:p>' + _N("User web-based security awareness training:") + '</w:p>'
                         '<w:p><w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>' + _B(a["sat"]) + '</w:p></w:tc></w:tr>')
    # Server table
    W = [2515, 2159, 2338, 2338]
    rows = a["servers"] or [{"name": "None on-site", "vm": "None", "age": "—", "obs": "—", "_none": True}]
    xml = ""
    for s in rows:
        first = _B(s["name"]) if s.get("_none") else _B(f"{s['name']} – ") + _PH("[OS]")
        cells = [first, _B(s["vm"]), _B(s["age"]), _V(s["obs"], "[OS support end]")]
        xml += "<w:tr>" + "".join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr><w:p>{r}</w:p></w:tc>'
                                  for w, r in zip(W, cells)) + "</w:tr>"
    D.rows_into_table_of("2A88CC48", xml)
    # Backup
    D.fill("6AA39218", _V(a["backup_server"], "[Product / status]"))
    D.fill("202887A6", _V(a["backup_m365"], "[Product / status]"))
    D.fill("4DE8B87B", _B(a["backup_pc"]))
    # Tickets resolved
    D.set_text("5E427042", _N(f"Tickets resolved in the last 3 months: {a['t_resolved']} ({a['period_label']})"))
    # PC hardware table
    CW = [2200, 1500, 1850, 1900, 820, 1090]

    def cell(w, text, hl=None, hdr=False, extra=""):
        rpr = "<w:rPr>" + (f'<w:highlight w:val="{hl}"/>' if hl else "") + "</w:rPr>"
        sh = '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if hdr else ""
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{sh}</w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                f'<w:r>{rpr}<w:t xml:space="preserve">{xesc(str(text))}</w:t></w:r>{extra}</w:p></w:tc>')
    trs = ["<w:tr><w:trPr><w:tblHeader/></w:trPr>" + "".join(cell(w, t, hdr=True) for w, t in zip(CW, ["Model", "PC name", "User", "Processor", "Ram", "PC age"])) + "</w:tr>"]
    cur_site = None
    for r in a["pcs"]:
        if a.get("pc_sites") and r["site"] != cur_site:
            cur_site = r["site"]
            n_site = a["pc_site_count"][cur_site]
            trs.append('<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:gridSpan w:val="6"/>'
                       '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr><w:p><w:pPr><w:keepNext/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' % sum(CW)
                       + _B(f"{cur_site} – {n_site} PC{'s' if n_site != 1 else ''}") + "</w:p></w:tc></w:tr>")
        hl = "yellow" if (r["age"] or 0) >= 5 else ("cyan" if (r["age"] or 0) >= 4 else None)
        w10 = '<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:highlight w:val="green"/></w:rPr><w:t>W10</w:t></w:r>' if r["w10"] else ""
        vals = [r["model"], (r["name"] or "") + (" †" if r["stale"] else ""), r["user"], r["cpu"], r["ram"],
                (f"{r['age']} years" + (" ◊" if r.get("age_src") == "cpu" else "")) if r["age"] is not None else "[age]"]
        trs.append("<w:tr><w:trPr><w:cantSplit/></w:trPr>" + cell(CW[0], vals[0], hl, extra=w10) +
                   "".join(cell(w, v, hl) for w, v in zip(CW[1:], vals[1:])) + "</w:tr>")
    borders = "<w:tblBorders>" + "".join(f'<w:{k} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for k in ["top", "left", "bottom", "right", "insideH", "insideV"]) + "</w:tblBorders>"
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(CW)}" w:type="dxa"/>{borders}<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
           "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in CW) + "</w:tblGrid>" + "".join(trs) + "</w:tbl>")
    SP = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    n_w10 = sum(r["w10"] for r in a["pcs"])
    match = "matching Halo" if a["halo_ws"] == len(a["pcs"]) else f"Halo lists {a['halo_ws']}"
    notes = [f"<w:p>{SP}</w:p>",
             f"<w:p>{SP}" + _N(f"{len(a['pcs'])} workstations in N-sight ({match})" +
                               (f" across {len(a['pc_sites'])} sites, grouped by site and sorted oldest to newest. " if a.get("pc_sites") else ". ") +
                               (f"{n_w10} still run Windows 10 (marked W10). " if n_w10 else "All run Windows 11, so no machines are flagged W10. ") +
                               "PC age is years since the PC went into service, taken from the date in its name (e.g. -0624 = June 2024). "
                               + ("◊ No in-service date in the name, so age is years since the processor generation was released."
                                  if any(r.get("age_src") == "cpu" for r in a["pcs"]) else "")) + "</w:p>"]
    if any(r["stale"] for r in a["pcs"]):
        notes.append(f"<w:p>{SP}" + _N("† Not checked in to N-sight for 30+ days. Confirm whether these are retired.") + "</w:p>")
    if not a["pcs"]:
        tbl, notes = "", [f"<w:p>{SP}" + _PH("[No N-sight workstations found]") + "</w:p>"]
    D.after("35DED51D", tbl + "".join(notes))
    for pid in ("5A81D4C2", "156E6B1E", "2A88CC48"):
        D.drop_empty(pid)

    # Sentence case for generic labels; product names keep their capitals. Section heading Server -> Servers and Network.
    for old, new in WORD_LABELS.items():
        D.relabel(old, new)
    files["word/document.xml"] = D.x.encode("utf-8")
    _brand_docx(files)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    return out_path


# --------------------------------------------------------------------------- #
# Portrait PDF
# --------------------------------------------------------------------------- #
CSS = """
@page { size: 8.5in 11in; margin: 0; }
* { box-sizing: border-box; margin: 0; padding: 0; }
:root { --navy:#14213D; --navy2:#1F2E50; --paper:#F7F5F0; --card:#FFFDF9; --line:#E2DDD3; --ink:#1E2A44; --soft:#4A5568; --muted:#5B6577; --accent:#B5531F; --accent2:#E8894F; --green:#1F6E4A; --yellow:#FFE97A; --cyan:#9EE6EF; }
body { font-family: Calibri, Carlito, 'Liberation Sans', Arial, sans-serif; color: var(--ink); -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.page { width: 8.5in; height: 11in; padding: 0.7in 0.7in 0.9in; position: relative; background: var(--paper); page-break-after: always; display: flex; flex-direction: column; gap: 22px; overflow: hidden; }
.page:last-child { page-break-after: auto; }
.dark { background: var(--navy); color: #F4F1EA; }
.eyebrow { font-size: 13px; font-weight: 700; letter-spacing: 2.5px; text-transform: uppercase; color: var(--accent); }
.dark .eyebrow { color: var(--accent2); }
h1, h2 { font-family: Cambria, 'Liberation Serif', Georgia, serif; font-weight: 700; line-height: 1.12; color: var(--navy); }
.dark h1, .dark h2 { color: #F4F1EA; }
h2 { font-size: 36px; margin-top: 8px; }
h3 { font-size: 19px; font-weight: 700; color: var(--navy); }
p { font-size: 15px; line-height: 1.45; }
.lead { font-size: 16.5px; color: var(--soft); }
.dark .lead { color: #D5DCE8; }
.foot { position: absolute; left: 0.7in; right: 0.7in; bottom: 0.45in; display: flex; justify-content: space-between; font-size: 11px; color: var(--muted); border-top: 1px solid var(--line); padding-top: 8px; }
.dark .foot { color: #BFC8D9; border-color: #3A4A6B; }
.grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
.grid2.g3c .em { grid-column: span 2; }
.grid3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; }
.card { background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 16px 18px; display: flex; flex-direction: column; gap: 6px; }
.lbl { font-size: 12.5px; color: var(--muted); }
.val { font-size: 18px; font-weight: 700; line-height: 1.3; }
.dcard { background: var(--navy2); border-radius: 10px; padding: 16px 18px; display: flex; flex-direction: column; gap: 6px; }
.dcard .lbl { color: #BFC8D9; } .dcard .val { font-size: 16px; }
.ph { background: #E6E2DA; color: var(--soft); padding: 1px 6px; border-radius: 4px; font-weight: 700; }
.dark .ph { background: #3A4A6B; color: #F4F1EA; }
table { width: 100%; border-collapse: collapse; font-size: 14.5px; }
th { text-align: left; font-size: 12px; color: var(--muted); font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; padding: 8px 10px; border-bottom: 2px solid var(--ink); }
td { padding: 9px 10px; border-bottom: 1px solid var(--line); vertical-align: top; }
.num { font-family: Cambria, 'Liberation Serif', Georgia, serif; font-size: 46px; font-weight: 700; line-height: 1; color: var(--navy); }
.num.acc { color: var(--accent); }
.status { display: flex; gap: 18px; align-items: center; background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 24px; }
.status.ok { border-left: 6px solid var(--green); } .status.na { border-left: 6px solid #9AA8C2; } .status.gap { border-left: 6px solid var(--accent); }
.q { background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 16px 20px 6px; }
.q p.question { font-size: 16px; font-weight: 700; }
.q .lines div { height: 30px; border-bottom: 1px solid var(--line); }
.callout { display: flex; flex-direction: column; gap: 6px; border-left: 5px solid var(--accent2); padding: 4px 0 4px 18px; }
.swrap { display: flex; flex-direction: column; gap: 22px; }
.legend { display: flex; gap: 10px; } .legend span { font-size: 13px; font-weight: 700; padding: 4px 10px; border-radius: 5px; }
.rec { display: flex; gap: 20px; align-items: baseline; border-top: 1px solid #3A4A6B; padding: 16px 0; }
.rec:last-child { border-bottom: 1px solid #3A4A6B; }
.rec .n { width: 28px; font-family: Cambria, 'Liberation Serif', Georgia, serif; font-size: 26px; font-weight: 700; color: var(--accent2); }
.rec p.t { flex: 1; font-size: 18px; }
.spacer { flex: 1; }
"""

ICON_OK = '<svg width="34" height="34" viewBox="0 0 24 24" aria-label="Covered"><circle cx="12" cy="12" r="11" fill="#1F6E4A"/><path d="M7 12.5l3.2 3.2L17.5 8.5" stroke="#FFFDF9" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>'
ICON_NA = '<svg width="34" height="34" viewBox="0 0 24 24" aria-label="Not applicable"><circle cx="12" cy="12" r="11" fill="#9AA8C2"/><rect x="6" y="11" width="12" height="2.4" rx="1" fill="#FFFDF9"/></svg>'
ICON_GAP = '<svg width="34" height="34" viewBox="0 0 24 24" aria-label="To confirm"><circle cx="12" cy="12" r="11" fill="#B5531F"/><rect x="11" y="6" width="2.2" height="8" fill="#FFFDF9"/><rect x="11" y="16" width="2.2" height="2.2" fill="#FFFDF9"/></svg>'


def _h(v) -> str:
    return html.escape(str(v))


def _ph(t) -> str:
    return f'<span class="ph">{_h(t)}</span>'


def _vv(v, ph) -> str:
    return _h(v) if v not in (None, "") else _ph(ph)


def _portal_line() -> str:
    url = BRAND["helpdesk_url"]
    if not url:
        return '<p class="lead">Log in to the helpdesk portal for full ticket details.</p>'
    return (f'<p class="lead">Log in to the <a href="{_h(url)}" style="color:var(--accent);text-decoration:underline;'
            f'font-weight:600">Helpdesk Portal</a> for full ticket details.</p>')


def _feedback_block() -> str:
    out = ""
    if BRAND["review_url"]:
        out += (f'<div class="card" style="padding:20px"><p style="font-size:17px;font-weight:700">Satisfied? Please leave us a '
                f'<a href="{_h(BRAND["review_url"])}" style="color:var(--accent)">review</a>.</p></div>')
    contact = " · ".join(x for x in (BRAND["phone"], BRAND["email"]) if x)
    out += f'<p class="lead">Thank you for trusting {_h(BRAND["company"])}.' + (f" Support: {_h(contact)}" if contact else "") + "</p>"
    return out


def _brand_docx(files: dict) -> None:
    """Fill the {{COMPANY}}/{{PHONE}}/{{EMAIL}}/{{HELPDESK_URL}}/{{REVIEW_URL}} tokens in the template and swap the logo."""
    vals = {"{{COMPANY}}": BRAND["company"], "{{PHONE}}": BRAND["phone"], "{{EMAIL}}": BRAND["email"],
            "{{HELPDESK_URL}}": BRAND["helpdesk_url"] or "https://example.com",
            "{{REVIEW_URL}}": BRAND["review_url"] or "https://example.com"}
    for n, b in list(files.items()):
        if n.startswith("word/") and (n.endswith(".xml") or n.endswith(".rels")):
            s = b.decode("utf-8") if isinstance(b, bytes) else b
            if "{{" in s:
                for k, v in vals.items():
                    s = s.replace(k, xesc(v, {'"': "&quot;"}))
                files[n] = s.encode("utf-8")
    logo = Path(BRAND["logo"]) if BRAND["logo"] else None
    if logo and not logo.is_absolute():
        logo = HERE / logo
    if logo and logo.exists() and "word/media/image1.png" in files:
        files["word/media/image1.png"] = logo.read_bytes()


def build_html(d: dict, a: dict, landscape: bool = False) -> str:
    max_pcs = 12 if landscape else 22
    short = d["name"].rstrip("*")
    title = d["legal"]
    _y, _q = d.get("review_q", d["quarter"]).upper().split("-Q")
    q = f"Q{_q} {_y}"
    page_no = [0]

    def foot(src: str) -> str:
        page_no[0] += 1
        return f'<div class="foot"><span>{_h(src)}</span><span>{_h(short)} · {_h(q)} · {page_no[0]}</span></div>'

    def pct(v): return f"{round(v)}%" if v is not None else "—"
    pages = []
    page_no[0] = 1
    pages.append(f"""<section class="page dark" style="justify-content:space-between">
<p class="eyebrow">Strategic Technology Review · {_h(q)}</p>
<div style="display:flex;flex-direction:column;gap:18px"><h1 style="font-size:56px">{_h(title)}</h1>
{('<p style="font-size:22px;color:#E8894F;font-weight:600;margin-top:-6px">' + _h(client_since_text(d)) + '</p>') if client_since_text(d) else ''}
<p style="font-size:20px;color:#BFC8D9">Prepared by {_h(BRAND['company'])} · {_h(f"{MONTHS[d['ref'].month - 1]} {d['ref'].day}, {d['ref'].year}")}</p></div>
<div class="grid2 g3c" style="border-top:1px solid #3A4A6B;padding-top:22px;gap:22px 28px">
<div><p class="lbl" style="color:#BFC8D9">Primary contact</p><p class="val">{_vv(", ".join(x for x in ((d.get('contact') or {}).get('name'), (d.get('contact') or {}).get('title')) if x) or None, '[Name]')}</p></div>
<div><p class="lbl" style="color:#BFC8D9">Phone</p><p class="val">{_vv(d.get('phone'), '[phone]')}</p></div>
<div class="em"><p class="lbl" style="color:#BFC8D9">Email</p><p class="val" style="overflow-wrap:anywhere">{_vv((d.get('contact') or {}).get('email'), '[Email]')}</p></div>
<div><p class="lbl" style="color:#BFC8D9">Review date</p><p class="val">{_h(a['review_month'])} {a['review_day'] if a.get('review_day') else _ph('[day]')}, {a['year']}</p></div>
<div><p class="lbl" style="color:#BFC8D9">Support plan</p><p class="val">{_vv(a['plan'], '[support plan]')}</p></div>
{"".join(f'<div><p class="lbl" style="color:#BFC8D9">Secondary contact</p><p class="val">{_h(", ".join(x for x in (cc.get("name"), cc.get("title")) if x))}</p></div><div class="em"><p class="lbl" style="color:#BFC8D9">Secondary contact email</p><p class="val" style="overflow-wrap:anywhere">{_vv(cc.get("email"), "[Email]")}</p></div>' for cc in d.get("co_contacts") or [])}
</div></section>""")

    m = a["m365"]
    seats = lambda v: f"{v[0]} seats · {v[1]} assigned" if v else "None"  # noqa: E731
    bp = m["Business Premium"]
    if bp and bp[1] >= bp[0]:
        lic_head = f"All {bp[0]} Business Premium seats are in use"
    elif bp:
        lic_head = f"{bp[1]} of {bp[0]} Business Premium seats in use"
    else:
        lic_head = "Microsoft 365 licensing"
    def lic_row(o):
        unsupported = o["renews"] == "Unsupported"
        renews = ("Unsupported – " + o["why"].replace("out of support ", "")) if unsupported and o.get("why") else o["renews"]
        row_style = ' style="background:var(--yellow);font-weight:700"' if unsupported else ""
        return ("<tr" + row_style + "><td>" + _h(o["name"]) + "</td><td>" + _h(o["seats"]) + "</td><td>" +
                _h(o["used"] if o["used"] is not None else "—") + "</td><td>" + _h(renews) + "</td></tr>")
    other_rows = "".join(lic_row(o) for o in a["also_lic"])
    def plan_card(c):
        v = c["seats"]
        line1 = (f"{v[0]} seats · {v[1]} assigned" + (" (nonprofit)" if c["nonprofit"] else "")) if v else "None"
        if not v:
            line2 = '<p style="font-size:13px;color:var(--soft);margin-top:4px">—</p>'
        elif c["renews"]:
            col = "#B7791F" if c["soon"] else "var(--soft)"
            line2 = f'<p style="font-size:13px;color:{col};margin-top:4px;{"font-weight:700" if c["soon"] else ""}">Renews {_h(_mdy(c["renews"]))}</p>'
        else:
            line2 = '<p style="font-size:13px;color:var(--soft);margin-top:4px">Free nonprofit grant</p>'
        accent = ' style="border-left:6px solid #D9622B"' if c["label"] == "Business Premium" and v else ""
        return f'<div class="card"{accent}><p class="lbl">{_h(c["label"])}</p><p class="val">{_h(line1)}</p>{line2}</div>'
    plan_cards_html = "".join(plan_card(c) for c in a["plan_cards"])
    lic_lead = ""
    if bp:
        lic_lead = f"Business Premium includes Defender for Office 365 Plan 1 and Entra ID P1 for all {bp[0]} users."
        if bp[1] >= bp[0]:
            lic_lead += " With every seat in use, the next hire will need one added."
    pages.append(f"""<section class="page"><div><p class="eyebrow">Microsoft 365 Licensing</p><h2>{_h(lic_head)}</h2></div>
<div class="grid2">{plan_cards_html}</div>
{('<p class="lead">' + _h(lic_lead) + '</p>') if lic_lead else ''}
{'<h3>Also licensed</h3><table><tr><th style="width:46%">Subscription</th><th>Seats</th><th>In use</th><th>Renews</th></tr>' + other_rows + '</table>' if other_rows else ''}
{('<h3>Other product licensing</h3><table><tr><th style="width:46%">Product</th><th>Subscription</th></tr>' + ''.join('<tr><td>' + _h(p) + '</td><td>Yes</td></tr>' for p in a['other_products']) + '</table>') if a['other_products'] else ''}
{foot('Source: Halo licences, synced from the Microsoft tenant and distributor')}</section>""")

    fw = a["firewall"]
    fws = a["firewalls"]
    fw_txt = "<br>".join(_h(firewall_text(f, len(fws) > 1)) for f in fws) if fws else _ph("[firewall]")
    inc = a["incidents"]
    callout = ""
    if inc:
        t = inc[0]
        is_open = t in a["incidents_open"]
        callout = f"""<div class="callout"><p style="font-size:17px;font-weight:700">{_h(incident_title(t))} · {_h(_mdy(_dt(t.get('opened'))))}</p>
<p class="lead">{_h(incident_text(t))} {'Clean-up is still open' if is_open else 'Resolved'} (#{_h(t.get('id'))}).</p></div>"""
        if len(inc) > 1:
            callout += f'<p class="lead" style="color:#BFC8D9">{len(inc) - 1} other security alert{"s" if len(inc) > 2 else ""} this quarter.</p>'
    sec_head = "Monitoring caught a real attack" if inc else "No security incidents this quarter"
    pages.append(f"""<section class="page dark"><div><h2>Security</h2></div>
<div class="grid2">
<div class="dcard"><p class="lbl">Current firewall / Wi-Fi</p><p class="val">{fw_txt}<br>Wi-Fi {_h(d['wifi']) if d.get('wifi') else _ph('[model]')}</p></div>
<div class="dcard"><p class="lbl">Defender for Office 365</p><p class="val">{'Yes, Plan 1 (Business Premium)' if bp else _vv(a['defender'], '[Yes/No]')}</p></div>
<div class="dcard"><p class="lbl">Microsoft 365 identity protection</p><p class="val">{'<br>'.join(_h(x) for x in a['identity'])}</p></div>
<div class="dcard"><p class="lbl">LastPass</p><p class="val">{_vv(a['lastpass'], '—')}</p></div>
<div class="dcard"><p class="lbl">User web-based security awareness training</p><p class="val">{_h(a['sat'])}</p></div>
<div class="dcard"><p class="lbl">Most recent organization cybersecurity awareness training</p><p class="val">{_h(training_text(d))}</p></div>
<div class="dcard"><p class="lbl">Endpoint security protection</p><p class="val">{_h(a['endpoint'])}</p></div>
</div><h3 style="color:#F4F1EA;margin-top:4px">{sec_head}</h3>{callout}
{foot('Source: Halo assets, licences and tickets')}</section>""")

    srv_rows = "".join(f"<tr><td>{_h(s['name'])} – {_ph('[OS]')}</td><td>{_h(s['vm'])}</td><td>{_h(s['age'])}</td><td>{_vv(s['obs'], '[OS support end]')}</td></tr>" for s in a["servers"]) \
        or "<tr><td>None on-site</td><td>None</td><td>—</td><td>—</td></tr>"
    equip = []
    for f in fws:
        equip.append(("Firewall" + (f" · {f['site']}" if len(fws) > 1 and f.get("site") else ""), f["model"],
                      f"Warranty to {f['warranty']}" if f["warranty"] else ""))
    for pr in a["printers"][:2]:
        equip.append(("Printer", pr, ""))
    equip_html = "".join(f'<div class="card"><p class="lbl">{_h(k)}</p><p class="val">{_h(v)}</p><p style="font-size:13px;color:var(--soft)">{_h(n)}</p></div>' for k, v, n in equip)
    n_srv = len(a["servers"])
    srv_head = "No servers, on site or in Azure" if not n_srv else (str(n_srv) + " server" + ("s" if n_srv > 1 else ""))
    pages.append(f"""<section class="page"><div><p class="eyebrow">Servers and Network</p><h2>{srv_head}</h2></div>
<table><tr><th>Server / OS</th><th>VM / cloud</th><th>Age / warranty</th><th>Obsolescence</th></tr>{srv_rows}</table>
{('<h3>Managed Network Devices and Printers</h3><div class="grid3">' + equip_html + '</div>') if equip_html else ''}

{foot('Source: Halo assets, N-sight')}</section>""")

    def status(title_, v, ph, na=False):
        if na:
            return f'<div class="status na">{ICON_NA}<div><h3>{_h(title_)}</h3><p style="color:var(--soft)">{_h(v)}</p></div></div>'
        if v:
            return f'<div class="status ok">{ICON_OK}<div><h3>{_h(title_)}</h3><p style="color:var(--soft)">Covered: {_h(v)}</p></div></div>'
        return f'<div class="status gap">{ICON_GAP}<div><h3>{_h(title_)}</h3><p style="color:var(--soft)">{_ph(ph)}</p></div></div>'
    srv_na = a["backup_server"] and a["backup_server"].startswith("N/A")
    all_ok = a["backup_m365"] and (a["backup_server"])
    pages.append(f"""<section class="page"><div><p class="eyebrow">Backup</p><h2>{'Microsoft 365 and the PCs are both backed up' if all_ok else 'Backup coverage'}</h2></div>
<div class="swrap">{status('Server / NAS', 'Not needed: no on-site server' if srv_na else a['backup_server'], '[Product / status]', na=bool(srv_na))}
{status('Microsoft 365', a['backup_m365'], '[Product / status]')}
{status('PCs', a['backup_pc'] + (' backup' if a['backup_pc'] == 'Cove Documents' else ''), '[Product / status]')}</div>
{'<p class="lead">Email, files and each PC' + "'" + 's documents are all backed up off-site.</p>' if all_ok else ''}
{foot('Source: Halo subscribed services')}</section>""")

    pq = "Q" + prev_quarter(d["quarter"]).split("-Q")[1]
    cat_rows = "".join(f"<tr><td>{_h(c)}</td><td>{n}</td></tr>" for c, n in a["categories"][:8])
    pages.append(f"""<section class="page"><div><p class="eyebrow">Helpdesk tickets</p><h2>{a['t_resolved']} tickets resolved in the last 3 months</h2></div>
<div class="grid2 g4">
<div class="card" style="padding:20px"><p class="num">{a['t_resolved']}</p><p class="val">Resolved</p><p class="lbl">{pq}: {a['t_resolved_prev']}</p></div>
<div class="card" style="padding:20px"><p class="num">{a['t_opened']}</p><p class="val">Opened</p><p class="lbl">{_h(a['period_label'])}</p></div>
<div class="card" style="padding:20px"><p class="num acc">{pct(a['sla_resp'])}</p><p class="val">Responded on time</p><p class="lbl">{pq}: {pct(a['sla_resp_prev'])}</p></div>
<div class="card" style="padding:20px"><p class="num acc">{pct(a['sla_res'])}</p><p class="val">Resolved on time</p><p class="lbl">{pq}: {pct(a['sla_res_prev'])}</p></div>
</div>
<h3>What came in this quarter</h3><table><tr><th style="width:75%">Request type</th><th>Tickets</th></tr>{cat_rows}</table>
{_portal_line()}
<p class="lbl">On-time figures use Halo's SLA status on each ticket, which allows for on-hold time. Security alerts and other tickets excluded from SLA are not counted.</p>
{foot('Source: HaloPSA')}</section>""")

    pc_rows = ""
    shown = a["pcs"]
    if len(shown) > max_pcs:
        keep = sorted(shown, key=lambda r: -(r["age"] if r["age"] is not None else -1))[:max_pcs]
        keep_ids = {id(r) for r in keep}
        shown = [r for r in shown if id(r) in keep_ids]
    if a.get("pc_sites"):
        max_pcs_rows = max_pcs - len({r["site"] for r in shown})  # leave room for the site heading rows
        if len(shown) > max_pcs_rows:
            keep = sorted(shown, key=lambda r: -(r["age"] if r["age"] is not None else -1))[:max_pcs_rows]
            keep_ids = {id(r) for r in keep}
            shown = [r for r in shown if id(r) in keep_ids]
    cur_site = None
    for r in shown:
        if a.get("pc_sites") and r["site"] != cur_site:
            cur_site = r["site"]
            n_site = a["pc_site_count"][cur_site]
            pc_rows += (f'<tr><td colspan="6" style="background:#1F2E50;color:#F4F1EA;font-weight:700;letter-spacing:.3px">'
                        f'{_h(cur_site)} <span style="font-weight:400;color:#BFC8D9">· {n_site} PC{"s" if n_site != 1 else ""}</span></td></tr>')
        bg = "var(--yellow)" if (r["age"] or 0) >= 5 else ("var(--cyan)" if (r["age"] or 0) >= 4 else "")
        style = f' style="background:{bg}"' if bg else ""
        age = (f"{r['age']} yrs" + (" ◊" if r.get("age_src") == "cpu" else "")) if r["age"] is not None else _ph("[age]")
        w10 = ' <span style="background:#7BE37B;padding:0 4px;font-weight:700">W10</span>' if r["w10"] else ""
        model = re.sub(r"^Lenovo (?=[TXPLE]\d|X1)", "ThinkPad ", r["model"]).replace("ThinkPad X1", "X1")
        cpu = re.sub(r"^(Core Ultra \d)\s+\d{3}[A-Z]?$", r"\1", r["cpu"])
        if r["age"] is not None and r["age"] >= 5:
            age += " · now"
        elif r["age"] is not None and r["age"] >= 4:
            age += f" · {d['ref'].year + 1}"
        pc_rows += f"<tr{style}><td>{_h(model)}{w10}</td><td>{_h(r['name'])}{' †' if r['stale'] else ''}</td><td>{_h(r['user'])}</td><td>{_h(cpu)}</td><td>{_h(r['ram'])}</td><td>{age}</td></tr>"
    more = f"<p class='lbl'>Showing the {len(shown)} oldest of {len(a['pcs'])} PCs. The full list is in the Word document.</p>" if len(shown) < len(a["pcs"]) else ""
    n_now, n_soon = len(a["pc_now"]), len(a["pc_soon"])
    head = (f"{n_now} PC" + ("s" if n_now != 1 else "") + f" due now, {n_soon} within a year") if a["pcs"] else "PC hardware"
    pages.append(f"""<section class="page"><div><p class="eyebrow">PC Hardware</p><h2>{_h(head)}</h2></div>
<div class="legend"><span style="background:var(--yellow)">Obsolete, recommend replace or upgrade now</span><span style="background:var(--cyan)">Recommend replace in the next 12 months</span></div>
<table style="font-size:13px"><tr><th>Model</th><th>PC name</th><th>User</th><th>Processor</th><th>Ram</th><th>Age · replace</th></tr>{pc_rows}</table>{more}
<p class="lbl">PC age is years since the PC went into service (the date in its name, e.g. 2023-07 = July 2023 or -0624 = June 2024).{' ◊ No in-service date, so age is from the processor release.' if any(r.get('age_src') == 'cpu' for r in a['pcs']) else ''}{' † not seen in N-sight for 30+ days.' if any(r['stale'] for r in a['pcs']) else ''}</p>
{foot('Source: N-sight')}</section>""")

    def qpage(eyebrow, heading, qs, foot_src):
        body = "".join(f'<div class="q"><p class="question">{_h(qq)}</p><div class="lines"><div></div><div></div><div></div></div></div>' for qq in qs)
        return f'<section class="page"><div><p class="eyebrow">{_h(eyebrow)}</p><h2>{_h(heading)}</h2></div><div class="qwrap" style="display:{"grid" if landscape else "flex"};flex-direction:column;gap:14px">{body}</div>{foot(foot_src)}</section>'
    has_copilot = any("copilot" in (o["name"] or "").lower() for o in a["other_lic"])
    cop = next((o for o in a["other_lic"] if "copilot" in (o["name"] or "").lower()), None)
    staff = (bp or (0, 0))[1]
    ai_head = (f"{cop['used'] or cop['seats']} of {staff} staff already have Copilot" if cop and staff else "AI at " + short)
    pages.append(qpage("AI", ai_head, ["Are employees already using AI tools?",
                       (f"How is Copilot working for the {cop['used'] or cop['seats']} people who have it?" if cop else "How is Copilot working for the people who have it?") if has_copilot else "Have you tried or considered Copilot?",
                       "If you could automate one thing tomorrow, what would it be?", "Do you have an AI acceptable use policy?"], "Discussion"))
    pages.append(qpage("Other Tech Planning", f"What's coming up for {short}?", ["Any upcoming changes we should be aware of?",
                       "Are you satisfied with your current phone system?", "Is there a better time in the year to do this review in the future?"], "Discussion"))
    fb = qpage(f"Feedback for {BRAND['company']}", "How are we doing?", ["Is there anything we can do to improve our services?",
               "Do you know of anyone that may be a good fit for our services?"], "Discussion")
    fb = fb.replace('<div class="foot">', _feedback_block() + '<div class="foot">')
    pages.append(fb)

    # Recommendations slide removed (Oct 2026): to-dos are agreed with the client in the meeting.
    css = CSS + (CSS_LANDSCAPE if landscape else "")
    return f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>{_h(short)} Strategic Technology Review – {_h(q)}</title><style>{css}</style></head><body>{"".join(pages)}</body></html>'


CSS_LANDSCAPE = """
@page { size: 11in 8.5in; margin: 0; }
.page { width: 11in; height: 8.5in; padding: 0.55in 0.8in 0.8in; gap: 16px; }
.foot { left: 0.8in; right: 0.8in; bottom: 0.35in; }
h2 { font-size: 34px; }
.grid2 { grid-template-columns: 1fr 1fr 1fr; }
.status { padding: 16px 22px; }
.qwrap { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
.q .lines div { height: 26px; }
table { font-size: 13.5px; } td { padding: 7px 10px; }
.rec { padding: 11px 0; }
.grid2.g4 { grid-template-columns: repeat(4, 1fr); }
.grid2.g3c { grid-template-columns: 1fr 1fr 1fr; }
.grid2.g3c .em { grid-column: auto; }
.swrap { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; }
.num { font-size: 40px; }
"""


def html_to_pdf(html_path: Path, pdf_path: Path, landscape: bool = False) -> Path:
    """Print with Microsoft Edge (Windows) or Chromium; falls back to Playwright."""
    candidates = [os.environ.get("STR_BROWSER"),
                  r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                  r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                  shutil.which("msedge"), shutil.which("chromium"), shutil.which("google-chrome")]
    exe = next((c for c in candidates if c and Path(c).exists()), None)
    if pdf_path.exists():
        pdf_path.unlink()  # never hand back an older PDF if printing fails
    if exe:
        # Edge sometimes writes the PDF but never exits. Watch for the finished file instead of waiting
        # for the process, then close Edge ourselves. One retry if nothing was written.
        for attempt in (1, 2):
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as prof:
                cmd = [exe, "--headless=new"] + (["--no-sandbox"] if hasattr(os, "geteuid") and os.geteuid() == 0 else []) + [ "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf-no-header", "--no-first-run",
                       "--no-default-browser-check", "--disable-extensions", "--disable-background-networking",
                       f"--user-data-dir={prof}", f"--print-to-pdf={pdf_path}", html_path.resolve().as_uri()]
                proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                deadline, last_size, stable = time.time() + 90, -1, 0
                while time.time() < deadline:
                    if proc.poll() is not None:
                        break
                    size = pdf_path.stat().st_size if pdf_path.exists() else -1
                    stable = stable + 1 if size > 1000 and size == last_size else 0
                    last_size = size
                    if stable >= 4:  # same size for ~2 s: printing is done
                        break
                    time.sleep(0.5)
                if proc.poll() is None:
                    if os.name == "nt":
                        subprocess.run(["taskkill", "/T", "/F", "/PID", str(proc.pid)], capture_output=True)
                    else:
                        proc.kill()
                    try:
                        proc.wait(timeout=10)
                    except Exception:  # noqa: BLE001
                        pass
            if pdf_path.exists() and pdf_path.stat().st_size > 1000:
                return pdf_path
            log.info("Edge did not produce %s (attempt %s)", pdf_path.name, attempt)
    from playwright.sync_api import sync_playwright  # optional fallback
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page()
        pg.goto(html_path.resolve().as_uri())
        pg.pdf(path=str(pdf_path), width="11in" if landscape else "8.5in", height="8.5in" if landscape else "11in", print_background=True)
        b.close()
    return pdf_path


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def sharepoint_folder(d: dict) -> str | None:
    """Client folder from Halo's website field, e.g. 'Clients/ACME' (inside the SHAREPOINT library)."""
    w = d.get("website") or ""
    qs = parse_qs(urlparse(w).query)
    if "id" in qs:
        path = unquote(qs["id"][0]).strip("/")      # SHAREPOINT/Clients/ACME
        parts = path.split("/")
        return "/".join(parts[1:]) if len(parts) > 1 else None
    return None


def generate(client_id: int, quarter: str | None = None, out_dir: Path | None = None, overrides: dict | None = None,
             main_contact: str | None = None, review_date: str | None = None) -> dict:
    cfg = load_config()
    # The Quarter box is the REVIEW quarter (blank = the current quarter).
    # Tickets, SLA and incidents cover the full quarter before it: a Q4 2026 review reports on July-September 2026.
    review_q = normalize_quarter(quarter) or current_quarter()
    data_q = prev_quarter(review_q)
    log.info("Review %s, reporting on %s", review_q, data_q)
    d = collect(client_id, data_q, overrides or (cfg.get("overrides") or {}).get(str(client_id)), main_contact)
    d["review_q"] = review_q
    d["review_date"] = parse_review_date(review_date)
    if review_date and not d["review_date"]:
        d["warnings"].append(f"Review date '{review_date}' not understood; left as [day]")
    a = analyse(d)
    out_dir = Path(out_dir or cfg["output_dir"]) / safe_name(d["name"].rstrip("*"))
    out_dir.mkdir(parents=True, exist_ok=True)
    base = safe_name(d["name"].rstrip("*"))
    docx = build_docx(d, a, out_dir / f"{base}_{a['review_month']}{a['year']}_Strategic_Technology_Review.docx", Path(cfg["template"]))
    html_path = out_dir / f"{base}_{d['review_q'].replace('-', '_')}_Strategic_Technology_Review.html"
    html_path.write_text(build_html(d, a), encoding="utf-8")
    pdf = html_to_pdf(html_path, html_path.with_suffix(".pdf"))
    land_html = out_dir / f"{base}_{d['review_q'].replace('-', '_')}_Strategic_Technology_Review_Landscape.html"
    land_html.write_text(build_html(d, a, landscape=True), encoding="utf-8")
    pdf_land = html_to_pdf(land_html, land_html.with_suffix(".pdf"), landscape=True)
    return {"client": d["name"], "docx": docx, "pdf": pdf, "pdf_landscape": pdf_land, "html": html_path,
            "sharepoint_folder": sharepoint_folder(d), "gaps": a["gaps"], "quarter": f"{d['review_q']} review (data {d['quarter']})"}
