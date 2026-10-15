# STR generator – changelog and how it works

Notes for anyone (or any Claude chat) picking this up. Last updated: October 9, 2026.

## How a report is made
1. Someone clicks **Generate STR** on the SharePoint page and fills in the form (list: **STR Requests** on the root site).
   - **Client** (required)
   - **Quarter** (optional): the *review* quarter, e.g. `Q4 2026`. The data covers the full quarter before it (Q4 2026 review → July–September 2026). Blank = current quarter.
   - **Main contact** (optional): name or email of who to list first.
2. The background task **QBR STR Generator** (Windows Task Scheduler, runs `run.py --poll` at sign-in) checks the list every 60 s.
3. `str_generator.py` pulls Halo + N-sight, builds the Word doc and two PDFs (portrait + landscape), and uploads them to `<library>/Clients/<Client>/Strategic Tech Review`.
4. Log: `str-generator.log` (each run logs the subscribed services it found).

`run.py` reloads `str_generator.py` on every request, so edits to `str_generator.py` apply without a restart.
Edits to `run.py` need a task restart:
`Stop-ScheduledTask -TaskName "QBR STR Generator"; Start-ScheduledTask -TaskName "QBR STR Generator"`

## Moving the generator to another PC (runs with nobody signed in)
1. Copy the whole `halo-mcp` folder to the new PC (any location).
2. In an elevated PowerShell on the new PC, from that folder: `powershell -ExecutionPolicy Bypass -File .\str-generator\setup-service.ps1`
   - Options: `-InstallDir C:\QBR\halo-mcp` (default), `-ServiceUser svc-qbr` (default; or `DOMAIN\user`), `-SkipKeys`.
   - It copies the files, creates the account, installs all-users Python and a fresh .venv, opens a window as the service account to store the 3 keys (Halo, N-sight, SharePoint app secret), tests all three, registers the task (at startup, runs whether signed in or not, password saved), disables sleep on AC, and starts the task.
3. On the old PC: `Disable-ScheduledTask -TaskName "QBR STR Generator"` so only one machine answers requests.
4. The script grants the account "Log on as a batch job" (without it the task never starts: error 2147943785). If a domain GPO controls that right, add the account there.
5. Halo on the new PC uses its own API application; its Client ID is in `str-generator\config.json` under `env.HALO_CLIENT_ID`. Re-run with `-KeysOnly` to change the Client ID or keys.
6. If the service account's password changes, re-run the script with `-SkipKeys` to re-register the task with the new password.

## Where each value comes from

| Report item | Source |
|---|---|
| Client name, legal name, phone | Halo client record |
| Client since | Halo client custom field **Client Since** (`CFClientSince`) |
| Primary contact | First Halo user with **Important contact** ticked (or the request's Main contact) |
| Secondary contact(s) | Other Important contacts (up to 2), labelled "Secondary Contact" |
| Job titles | Halo user field `other2` |
| Support plan | Subscribed services in `services.plan` (config.json) |
| Subscribed services | Halo client **Subscribed Services** (used for plan, backup, Wi-Fi, SAT, ITDR; no longer listed on the report) |
| Firewall(s) | Every active Halo asset of type Router/Firewall, one per site, with warranty end |
| Wi-Fi | Client custom field **Wi-Fi** if set, otherwise "UniFi" when subscribed to service **Unifi**; else `[model]` |
| Identity protection | "Entra ID P1 · Microsoft Security Defaults / Conditional Access" (+ "Huntress ITDR" on its own line if subscribed) |
| Endpoint protection | "Huntress MDR + Defender" |
| User web-based security awareness training | "Yes" if subscribed to **Huntress SAT** (or any service named "… SAT" / containing "Awareness"), else "No" |
| Most recent organization cybersecurity awareness training | Tab **Cybersecurity Awareness Training** → **Last Training Date** (`CFLastCyberTraining`); empty = `??`. Next Training Date is never used. |
| Backup | Subscribed services in `services.backup` (config.json); written for Cove names (M365, PC All-In, Server Basic/Standard/Premium) |
| Tickets, SLA, incidents | Halo tickets for the data quarter (previous quarter for comparison) |
| PCs | N-sight hardware inventory, all pages |

## PC table rules
- Grouped by N-sight **site** when a client has more than one (largest site first), oldest to newest within each site.
- **Age** = years since the in-service date in the PC name's last 4 digits (`-0624` = June 2024).
  Names ending in `9999` (unknown) or without a valid MMYY use the processor release date instead, marked **◊**.
- Highlight: yellow ≥ 5 years (replace now), cyan 4–5 years (within 12 months).
- † = not seen in N-sight for 30+ days.
- Lenovo model codes are mapped in `MODELS` (e.g. `20RY` = T490). Add new codes there.

## Changes made October 8–9, 2026
- Review quarter vs data quarter (Q4 review reports on Q3); cover shows "Prepared by <company> · <today's date>" (company from `branding` in config.json).
- Identity protection wording; Huntress ITDR as separate line; no "confirm MFA" placeholder anywhere.
- Subscribed services read from the client record (any service, not a fixed list), plus a backup check for the Unifi service.
- Primary + Secondary contacts, Main contact box on the request form, job titles.
- Client Since under the client name.
- Security Awareness Training Yes/No; Last Training Date or `??`.
- All firewalls listed by site; section renamed "Managed Network Devices and Printers".
- Wi-Fi UniFi from the Unifi service or Wi-Fi field.
- PC table grouped by site; age from in-service date; "Dell Dell" model name fix.
- Helpdesk Portal link on the tickets page (`branding.helpdesk_url`).
- PDF printing no longer times out when Edge hangs after writing the file (watches for the finished PDF, closes Edge, retries once).
- Empty Client box on a request gives a clear error message.
- Recommendations slide removed; to-dos are agreed with the client in the meeting. The thank-you and support line moved to the Feedback page (now the last slide).

## Changes made October 12, 2026 (client feedback email)
- Duplicate licences merged: when the Microsoft tenant and the distributor list the same product (e.g. "Power BI Pro" and "NCE Power BI Pro (Non-Profit Pricing)"), only the active distributor row is shown, with "In use" taken from the tenant. Seats shown as whole numbers.
- Security slide: main heading is now "Security"; "Monitoring caught a real attack" / "No security incidents this quarter" sits above the incident details.
- "Last cybersecurity awareness training" renamed "Most recent organization cybersecurity awareness training".
- "Security Awareness Training" (Huntress SAT Yes/No) renamed "User web-based security awareness training".
- Sentence case for generic labels in slides and Word (product names keep capitals). Word template labels are relabelled in the generated copy only (`WORD_LABELS`); the template file is untouched.
- "Server" section renamed "Servers and Network" (slides and Word).
- Removed the "Email and files live in Microsoft 365…" sentence from the servers page.
- Removed the "subscribed services" list from the backup slide and from the Word licensing section.
- Local admin accounts (administrator, admin, localadmin, defaultuser0, plus any in `branding.local_admin_users`) show as "Unknown" in the PC user column (`LOCAL_ADMIN_USERS`).

## Changes made October 14, 2026 (licensing page redesign)
- Four plan cards: Business Premium, Business Standard, Business Basic, Microsoft 365 Copilot. Cards with seats first, then "None". Business Premium has the accent colour when it has seats.
- Each card: "N seats · N assigned" (+ "(nonprofit)" for Premium when applicable) and "Renews <date>" from that plan's own distributor record. Within 90 days of today = amber.
- A plan with seats but no distributor record shows "Free nonprofit grant" (slides and Word) and is not listed as a data gap.
- Removed the "Contract renewal date", "365 Apps for Business" and "Defender for Office 365" cards.
- "Also licensed": Defender for Office 365 Plan 1 first, 365 Apps if it has seats, then other seat-based subscriptions. Distributor names are tidied (`clean_lic_name`: "NCE Power BI Pro (Non-Profit Pricing)- Annual Paid Monthly" -> "Power BI Pro (nonprofit)").
- New "Other product licensing" table for third-party products without seats: non-Microsoft Halo licence records with no seat count, plus subscribed services listed in `OTHER_PRODUCT_SERVICES`. Omitted when empty.
- Word: the licensing table uses the same four boxes as the slides, each with seats and its own "Renews <date>" line.
- PC age: the in-service date is also read from names ending in YYYY-MM (e.g. PC2023-07 = July 2023; '-', '_' or '.' separator), as well as -MMYY.
- Review date: new optional "Review date" field (column `ReviewDate`) on the STR Requests form. When filled, the cover shows the full date and the Word file name uses that month; blank keeps the `[day]` placeholder. The service adds the column itself at start-up (`SharePoint.ensure_columns`). CLI: `--review-date 2026-10-20`. Changes to run.py/sharepoint.py need a task restart.

## Related files
- `..\templates\Client_MonthDATE_Strategic_Technology_Review_Template.docx` – the Word template (never edited by the script).
