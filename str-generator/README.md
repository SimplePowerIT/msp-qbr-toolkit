# QBR STR generator

Branding (company name, phone, email, helpdesk and review links, logo) and your Halo service IDs
are set in `config.json`; copy `config.example.json` to start.

A **Generate STR** button on a SharePoint page opens a short form with a
client dropdown. Picking a client and saving creates a request; this generator (running on
a Windows PC) picks it up within a minute, pulls the data from Halo and N-sight, and
saves two files to the client's folder:

`<library>\Clients\<Client>\Strategic Tech Review\`
- `<Client>_<Month><Year>_Strategic_Technology_Review.docx` – the Word STR template, filled
- `<Client>_<Year>_Q<n>_Strategic_Technology_Review.pdf` – the portrait summary
- `<Client>_<Year>_Q<n>_Strategic_Technology_Review_Landscape.pdf` – the same pages as landscape slides

The request row then shows **Done** with a link to the folder, or **Failed** with the reason.
Anything Halo and N-sight don't record (contact name, email, Wi-Fi, training date, server OS)
stays a gray placeholder to fill in by hand.

It runs only while the laptop is on and signed in. Requests made while it's off are
processed when it's next running.

## 1. Register the app in Microsoft Entra (one time, admin)

1. https://entra.microsoft.com → **App registrations** → **New registration**.
   Name: `QBR STR Generator`. Single tenant. No redirect URI. Register.
2. Copy the **Application (client) ID** and **Directory (tenant) ID**.
3. **Certificates & secrets** → **New client secret** (24 months). Copy the secret **Value**.
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
   → `Sites.ReadWrite.All` and `Sites.Manage.All` → **Grant admin consent**.
   (`Sites.Manage.All` is only needed to create the two lists during setup; you can remove it afterwards.)

   Tighter alternative: use `Sites.Selected` instead and grant the app write access to just
   the one SharePoint site. Ask Claude if you want the steps.

## 2. Run setup on the laptop

Open PowerShell in this folder (`Documents\halo-mcp\str-generator`) and run:

```
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

It will:
1. Generate a test STR locally for a client you name. Check it in `output\` first.
2. Ask for the tenant ID, client ID, SharePoint site and library, and the secret. The
   secret goes into Windows Credential Manager, not a file.
3. Create two SharePoint lists on that site and fill the client list from Halo:
   - **STR Clients** feeds the dropdown.
   - **STR Requests** holds the button's requests and their status.
4. Install a background task that starts when you sign in and checks for requests every
   minute. It also re-syncs the client dropdown from Halo once a day.

At the end it prints the **button link**, which ends in `/Lists/STR Requests/NewForm.aspx`.

## 3. Add the button to a SharePoint page

1. Open the page → **Edit**.
2. Add a **Button** web part. Label: `Generate STR`. Link: the button link from setup.
3. **Republish**.

Optional: add a **List** web part showing **STR Requests**, so people can see the status
and folder link without leaving the page.

## Everyday use

- **From SharePoint:** click **Generate STR**, pick the client, and save. Leave Quarter
  blank for the last completed quarter, or type something like `2026-Q2`.
- **From the laptop:**
  - `..\.venv\Scripts\python.exe run.py --client "ACME"` creates the files locally.
  - Add `--upload` to also save them to SharePoint.
- **If it stops:** check the log in `str-generator.log`. To restart the background task,
  run `Start-ScheduledTask "QBR STR Generator"`.

## Where each field comes from

| Section | Source |
|---|---|
| Client, phone | Halo client record |
| Support plan | Halo Subscribed Services listed under `services.plan` in config.json |
| Microsoft 365 licensing | Halo licences: seats and in-use from the Microsoft tenant sync, renewal from the distributor |
| Firewall, servers, printer, software | Halo assets |
| Backup | Microsoft 365 and server backup from Subscribed Services listed under `services.backup` (written for N-able Cove service names). PCs are Cove Documents unless Cove PC All-In is subscribed |
| Tickets resolved, SLA | Halo tickets for the quarter |
| PC Hardware | N-sight. Age is years since the processor generation's release: 5+ years yellow, 4–5 years cyan, W10 flagged green |

The support plan and backup are found by looking up this client in the Halo service
records. If Halo's API doesn't expose that for a service, the field stays a placeholder and
the request's message says so.
