---
name: "msp-qbr"
description: "Build a client QBR from Halo and N-sight: a Slides deck plus the filled Strategic Technology Review Word document. Use when asked for a QBR or technology review for a client."
---

# Client QBR (HaloPSA + N-sight)

Produces two deliverables for one client and one quarter:
1. A QBR slide deck (Slides artifact) in the layout of your reference deck.
2. The filled "Strategic Technology Review" Word document, named `<Client>_<Month><Year>_Strategic_Technology_Review.docx`.

Default period: the last completed calendar quarter, compared with the quarter before it. Say which quarter you used.

## Sources and tools

- **Halo** (local connector `halo`, read-only): `list_clients`, `get_client`, `ticket_summary(client_id, quarter="YYYY-Qn")`, `search_tickets(status="open")`, `list_assets`, `list_contracts`, `list_licences`, `list_recurring_invoices`.
- **N-sight** (local connector `nsight`, read-only): `list_clients(search)`, `hardware_inventory(clientid, offset, limit=25)`. Page with `next_offset` until it is null. The N-sight client id often matches Halo's `gficlientid` field on the client record.
- Both connectors live in the install folder (default `%USERPROFILE%\Documents\halo-mcp`) (N-sight in the `nsight-mcp` subfolder). If their tools are missing, run RefreshMcpTools, then ask the user to fully quit and reopen Claude desktop.
- **Word template**: `<install folder>\templates\Client_MonthDATE_Strategic_Technology_Review_Template.docx`. Request access to that folder if needed and stage the file. Never edit the template in place.
- **Reference deck**: `<link to your reference QBR Slides artifact>` (set this to a finished deck of your own). Read its `project/deck.json` and slide files and reuse the same typefaces, palette, layouts and slide order for the new client. Create a NEW Slides artifact for each client; never overwrite the reference deck.

Never put API keys in chat, logs or files.

## Steps

1. **Find the client** in Halo (`list_clients`) and N-sight (`list_clients`). Confirm the match if names differ.
2. **Pull data**:
   - `ticket_summary` for this quarter and last quarter.
   - Open tickets (`search_tickets status=open`), to see open projects and recurring issues.
   - `get_client` for the phone number and legal name (`accountsid`).
   - `list_assets` (servers, firewall, certificates and warranties), `list_contracts`, `list_licences`, `list_recurring_invoices`.
   - N-sight `hardware_inventory` for all workstations.
3. **Analyse**:
   - Ticket volume by month, SLA response and resolution % against last quarter, median and mean resolution time, top categories, top requesters.
   - Recurring themes: search ticket summaries for repeats.
   - Licensing: use the Microsoft tenant-synced licences (type 0, `count` and `consumedcount`) for seats and in-use numbers. Use the distributor records (Synnex, type 1) for renewal dates. Ignore distributor records whose `end_date` has passed, and flag them as stale.
   - Hardware: see the rules below.
4. **Build the deck** from the Slides type, following the reference deck. Slides, in order:
   - cover, glance, volume, sla, categories, recurring, security, licensing, environment, lifecycle, projects, next steps, internal data-gaps.
   - Put page numbers in each footer, and speaker notes on every slide.
   - The data-gaps slide is labelled "Internal · remove before sending".
5. **Fill the Word template** (docx skill: unzip, edit `word/document.xml`, rezip, validate, then render to check the layout):
   - Values go in the empty bold paragraph under each label.
   - Anything not in the data becomes a gray-highlighted placeholder like `[qty]`, `[OS]` or `[date]`. Never guess.
   - **Header:** Client = Halo name plus legal name. Phone = `main_phonenumber`. Review date = month and year with a `[day]` placeholder. Contact, email and support plan stay placeholders unless a contract exists.
   - **Microsoft 365 Licensing:**
     - Business Premium, Standard, Basic and Apps: seats with in-use count in brackets, or "None".
     - Contract Renewal Date: the annual NCE term end.
     - Defender for Office 365: "Plan 1 – included in Business Premium" when the client has Business Premium.
     - Add one line under the table listing other subscriptions (Copilot, Teams Phone, Acrobat and so on).
   - **Security:**
     - Firewall: from Halo assets, with its warranty end.
     - Defender: as in the licensing section.
     - Identity protection: "Entra ID P1 (Business Premium)" plus the placeholder `[confirm MFA / Conditional Access]`.
     - LastPass: only if tickets or licences show it.
     - Training date: placeholder.
     - Leave "Huntress MDR + Defender" as it is.
   - **Server:** one row per server from Halo assets: name – `[OS]`; physical model or "VM"; purchase year / warranty end; obsolescence. Keep the table short enough to stay on one page.
   - **Backup:** placeholders.
   - **Tickets resolved in the last 3 months:** `tickets_closed` for the quarter, with the month range.
   - **PC Hardware:** the table described below.
   - Leave the AI, Other Tech Planning and Feedback sections untouched. They are discussion questions for the meeting.
6. **Deliver**:
   - Send the .docx.
   - Summarise in a few lines what was filled, what is still a placeholder, and the data problems found.
   - Remind the user the deck is private until shared, and that it downloads as .pptx.

## PC Hardware rules

Table columns: Model | PC name | User | Processor | Ram | PC age. Sort oldest first. Include every N-sight workstation.

- **Model:** the friendly name from the Lenovo or Dell model code:
  - 20N4 = T590
  - 20S6 = T15 Gen 1
  - 20W4 = T15 Gen 2
  - 20XW = X1 Carbon Gen 9
  - 21BV = T16 Gen 1
  - 21HH = T16 Gen 2
  - 21HM = X1 Carbon Gen 11
  - 21MN = T16 Gen 3
  - 21KC = X1 Carbon Gen 12
  - 21KS = P16s Gen 3
  - Unknown codes: "Lenovo <model> (<code>)".
  - Dell: "Dell <model>".
- **User:** the last user with the domain prefix removed.
- **Processor:** short form, like "8th Gen i7" or "Core Ultra 7 155U".
- **Ram:** bytes ÷ 2^30, in GB.
- **PC age:** years since the processor generation was released, counted to the review month, to one decimal. Not the purchase date. Release months:
  - i7-8565U 2018-08
  - i7-10510U 2019-08
  - i7-1165G7 2020-09
  - i5-11500 2021-03
  - i7-1260P 2022-02
  - i5-12500 / i7-12700 2022-01
  - i7-1355U / i5-13500 2023-01
  - Core Ultra 7 155U / 155H 2023-12
  - Core Ultra 7 258V 2024-09
  - Look up any processor not listed and add it here.
- **Highlighting** (on the text runs, as in the template legend):
  - Yellow = 5+ years: obsolete, replace now.
  - Cyan = 4–5 years: replace within 12 months.
  - Green highlight on "W10" after the model for any Windows 10 machine.
- Mark with † any PC that hasn't checked in to N-sight for 30+ days (scan_time). Add a note below the table asking to confirm whether those are retired.
- Below the table, add one line saying age is years since processor release, and whether any machine runs Windows 10.

## Things to always flag

- Halo asset count vs N-sight workstation count.
- Business Premium (or any plan) with every seat in use.
- Expired certificates or warranties shown in Halo.
- Stale distributor subscription records.
- Uncategorized or unassigned tickets.
- Tickets all logged to one site.
- No contract or recurring invoice on the client record.