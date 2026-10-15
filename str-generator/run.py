"""
STR generator command line.

  python run.py --client "ACME"            generate one STR locally (no SharePoint)
  python run.py --client "ACME" --upload   ...and upload to the client's SharePoint folder
  python run.py --setup                    create the SharePoint lists and fill the client dropdown
  python run.py --sync-clients             refresh the client dropdown from Halo
  python run.py --once                     process pending SharePoint requests once
  python run.py --poll                     keep processing requests (every 60 s; clients re-synced daily)
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
import traceback
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import str_generator as G  # noqa: E402

LOG = HERE / "str-generator.log"
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s",
                    handlers=[logging.FileHandler(LOG, encoding="utf-8"), logging.StreamHandler(sys.stdout)])
log = logging.getLogger("str.run")
# Keep request URLs (which can carry keys) out of the log.
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)


def find_client(name_or_id: str) -> dict:
    H = G.halo_mod()
    if name_or_id.isdigit():
        return {"id": int(name_or_id), "name": H.get_client(int(name_or_id)).get("name")}
    hits = H.list_clients(search=name_or_id)
    exact = [c for c in hits if c["name"].rstrip("*").lower() == name_or_id.lower()]
    if len(exact) == 1 or len(hits) == 1:
        return (exact or hits)[0]
    raise SystemExit(f"'{name_or_id}' matched {len(hits)} clients: " + ", ".join(c["name"] for c in hits[:15]))


def sp():
    from sharepoint import SharePoint
    cfg = G.load_config()
    if not cfg.get("graph"):
        raise SystemExit("No 'graph' section in config.json; run setup.ps1 first.")
    return SharePoint(cfg)


def generate_and_upload(client_id: int, quarter: str | None, s=None, main_contact: str | None = None,
                        review_date: str | None = None) -> dict:
    # Reload the generator each time so edits to str_generator.py apply without restarting the task.
    global G
    try:
        import importlib
        G = importlib.reload(G)
    except Exception as e:  # noqa: BLE001
        log.error("Reload of str_generator failed, using the loaded version: %s", e)
    res = G.generate(client_id, quarter or None, main_contact=main_contact or None, review_date=review_date or None)
    log.info("Generated %s, %s and %s", res["docx"].name, res["pdf"].name, res["pdf_landscape"].name)
    if s is not None:
        folder = res["sharepoint_folder"] or f"Clients/{G.safe_name(res['client'].rstrip('*'))}"
        if not res["sharepoint_folder"]:
            res["gaps"].append(f"No SharePoint folder on the Halo client record; used {folder}")
        s.upload(str(res["docx"]), folder, res["docx"].name)
        s.upload(str(res["pdf"]), folder, res["pdf"].name)
        s.upload(str(res["pdf_landscape"]), folder, res["pdf_landscape"].name)
        res["folder_url"] = s.folder_url(folder)
        log.info("Uploaded to %s", res["folder_url"] or folder)
    return res


def sync_clients(s) -> None:
    clients = G.halo_mod().list_clients(max_results=2000)
    added, updated = s.sync_clients([c for c in clients if not c.get("inactive")])
    log.info("Client dropdown synced from Halo: %s added, %s renamed", added, updated)


def process_requests(s) -> int:
    done = 0
    for item in s.pending_requests():
        f = item["fields"]
        iid = item["id"]
        try:
            s.update_request(iid, Status="Working", Message="Generating…")
            lookup = f.get("ClientLookupId")
            if not lookup:
                raise RuntimeError("No client selected. Open this request, pick a client and set Status back to Requested.")
            halo_id = s.halo_id_for_client_item(lookup)
            if not halo_id:
                raise RuntimeError("No client selected")
            res = generate_and_upload(halo_id, (f.get("Quarter") or "").strip() or None, s,
                                      (f.get("MainContact") or "").strip() or None,
                                      (f.get("ReviewDate") or "").strip() or None)
            msg = f"Done {time.strftime('%Y-%m-%d %H:%M')} · {res['quarter']} · {res['docx'].name}, portrait PDF and landscape PDF."
            if res["gaps"]:
                msg += " Check: " + "; ".join(res["gaps"])
            s.update_request(iid, Status="Done", Message=msg[:4000], Link=res.get("folder_url", "")[:255])
            done += 1
        except Exception as e:  # noqa: BLE001
            log.error("Request %s failed: %s\n%s", iid, e, traceback.format_exc())
            try:
                s.update_request(iid, Status="Failed", Message=f"{type(e).__name__}: {e}"[:4000])
            except Exception:  # noqa: BLE001
                pass
    return done


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--client")
    ap.add_argument("--quarter", help="review quarter, e.g. 2026-Q4; data covers the quarter before (default: current quarter)")
    ap.add_argument("--main-contact", help="name or email of the contact to list first")
    ap.add_argument("--review-date", help="review meeting date for the cover, e.g. 2026-10-20 (default: [day] placeholder)")
    ap.add_argument("--upload", action="store_true")
    ap.add_argument("--setup", action="store_true")
    ap.add_argument("--sync-clients", action="store_true")
    ap.add_argument("--once", action="store_true")
    ap.add_argument("--poll", action="store_true")
    args = ap.parse_args()
    G.load_config()

    if args.client:
        c = find_client(args.client)
        res = generate_and_upload(c["id"], args.quarter, sp() if args.upload else None, args.main_contact, args.review_date)
        print(f"\nWord: {res['docx']}\nPDF:  {res['pdf']}\nLandscape PDF: {res['pdf_landscape']}")
        if res.get("folder_url"):
            print(f"SharePoint: {res['folder_url']}")
        for g in res["gaps"]:
            print(f"Check: {g}")
        return 0
    if args.setup:
        s = sp()
        urls = s.ensure_lists()
        sync_clients(s)
        print("\nLists ready:")
        for n, u in urls.items():
            print(f"  {n}: {u}")
        print(f"\nButton link for your SharePoint page:\n  {urls['STR Requests'].rstrip('/')}/NewForm.aspx")
        return 0
    if args.sync_clients:
        sync_clients(sp())
        return 0
    if args.once:
        log.info("Processed %s request(s)", process_requests(sp()))
        return 0
    if args.poll:
        s, last_sync = sp(), 0.0
        try:
            added = s.ensure_columns()
            if added:
                log.info("Added column(s) to STR Requests: %s", ", ".join(added))
        except Exception as e:  # noqa: BLE001
            log.error("Could not check STR Requests columns: %s", e)
        log.info("Polling STR Requests every 60 s")
        while True:
            try:
                if time.time() - last_sync > 24 * 3600:
                    sync_clients(s)
                    last_sync = time.time()
                process_requests(s)
            except Exception as e:  # noqa: BLE001
                log.error("Poll cycle failed: %s", e)
            time.sleep(60)
    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
