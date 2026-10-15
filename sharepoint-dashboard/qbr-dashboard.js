define("__WEBPART_ID__", ["@microsoft/sp-webpart-base"], function (wpBase) {
  "use strict";
  var Base = wpBase.BaseClientSideWebPart;

  var CSS = [
    ".qbrd{--ink:#1b2631;--muted:#5c6b78;--line:#dde3e8;--sunk:#f3f5f7;--surface:#fff;--accent:#0a6a8f;--accent-soft:#dcedf4;",
    "--crit:#b42318;--crit-soft:#fbe4e1;--warn:#9a5800;--warn-soft:#fdf0d8;--info:#2f5bd3;--info-soft:#e2e9fb;--ok:#2e7d4f;--ok-soft:#e0f1e6;--idle:#66727f;--idle-soft:#eceff2;",
    "font-family:'Segoe UI',system-ui,-apple-system,sans-serif;color:var(--ink);font-size:14px;line-height:1.45;display:grid;gap:18px}",
    ".qbrd *{box-sizing:border-box}",
    ".qbrd h2,.qbrd h3{margin:0;font-weight:700;letter-spacing:-.01em}",
    ".qbrd h2{font-size:20px}.qbrd h3{font-size:16px}",
    ".qbrd .eyebrow{font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}",
    ".qbrd .muted{color:var(--muted)}",
    ".qbrd .mono{font-variant-numeric:tabular-nums}",
    ".qbrd button,.qbrd input,.qbrd select,.qbrd textarea{font:inherit;color:inherit}",
    ".qbrd button{cursor:pointer}",
    ".qbrd :focus-visible{outline:2px solid var(--accent);outline-offset:2px}",
    ".qbrd .top{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:10px 20px}",
    ".qbrd .ctrls{display:flex;flex-wrap:wrap;gap:8px;align-items:center}",
    ".qbrd .seg{display:inline-flex;border:1px solid var(--line);border-radius:999px;padding:2px;background:var(--surface)}",
    ".qbrd .seg button{border:0;background:none;padding:4px 11px;border-radius:999px;font-size:12.5px;color:var(--muted)}",
    ".qbrd .seg button[aria-pressed=true]{background:var(--ink);color:#fff}",
    ".qbrd .kpis{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}",
    "@media (max-width:640px){.qbrd .kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}",
    ".qbrd .kpi{border:1px solid var(--line);border-radius:8px;padding:10px 14px;display:grid;gap:1px;background:var(--surface);text-align:left;border-top:3px solid var(--line)}",
    ".qbrd button.kpi:hover{border-color:var(--muted)}",
    ".qbrd .kpi .v{font-size:28px;font-weight:700;font-variant-numeric:tabular-nums;line-height:1.1}",
    ".qbrd .kpi .l{font-size:12px;color:var(--muted)}",
    ".qbrd .kpi.crit{border-top-color:var(--crit)}.qbrd .kpi.crit .v{color:var(--crit)}",
    ".qbrd .kpi.warn{border-top-color:var(--warn)}.qbrd .kpi.warn .v{color:var(--warn)}",
    ".qbrd .kpi.info{border-top-color:var(--accent)}.qbrd .kpi.ok{border-top-color:var(--ok)}",
    ".qbrd .notice{border:1px solid var(--line);border-radius:10px;overflow:hidden;background:var(--surface)}",
    ".qbrd .nh{display:flex;flex-wrap:wrap;gap:8px 16px;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid var(--line)}",
    ".qbrd .tabs{display:flex;flex-wrap:wrap;gap:6px}",
    ".qbrd .tab{border:1px solid var(--line);background:var(--surface);border-radius:999px;padding:3px 10px;font-size:12.5px}",
    ".qbrd .tab[aria-pressed=true]{border-color:var(--ink);box-shadow:inset 0 0 0 1px var(--ink)}",
    ".qbrd .queue{list-style:none;margin:0;padding:0;max-height:560px;overflow:auto}",
    ".qbrd .qi{display:grid;grid-template-columns:4px minmax(0,1fr) auto;gap:0 14px;align-items:center;padding:10px 16px 10px 0;border-bottom:1px solid var(--line)}",
    ".qbrd .qi:last-child{border-bottom:0}",
    ".qbrd .stripe{align-self:stretch;border-radius:0 3px 3px 0}",
    ".qbrd .qm{min-width:0;display:grid;gap:2px}",
    ".qbrd .qn{font-weight:600;font-size:15px;display:flex;flex-wrap:wrap;gap:6px;align-items:center}",
    ".qbrd .link{all:unset;cursor:pointer;border-bottom:1px dotted var(--muted)}",
    ".qbrd .meta{font-size:12.5px;color:var(--muted);display:flex;flex-wrap:wrap;gap:2px 12px}",
    ".qbrd .acts{display:flex;flex-wrap:wrap;gap:6px;justify-content:flex-end;align-items:center}",
    ".qbrd .btn{border:1px solid var(--line);background:var(--surface);border-radius:6px;padding:4px 10px;font-size:12.5px;font-weight:600}",
    ".qbrd .btn:hover{border-color:var(--muted)}",
    ".qbrd .btn.primary{background:var(--accent);border-color:var(--accent);color:#fff}",
    ".qbrd .btn.ghost{border-color:transparent;background:none;color:var(--muted)}",
    ".qbrd .btn[disabled]{opacity:.5;cursor:default}",
    ".qbrd input[type=date],.qbrd select,.qbrd input[type=text],.qbrd textarea{border:1px solid var(--line);border-radius:6px;padding:4px 7px;background:var(--surface)}",
    ".qbrd .chip{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:1px 8px;font-size:11.5px;font-weight:600;white-space:nowrap}",
    ".qbrd .chip:before{content:'';width:6px;height:6px;border-radius:50%;background:currentColor}",
    ".qbrd .c-overdue{background:var(--crit-soft);color:var(--crit)}.qbrd .s-overdue{background:var(--crit)}",
    ".qbrd .c-soon{background:var(--warn-soft);color:var(--warn)}.qbrd .s-soon{background:var(--warn)}",
    ".qbrd .c-confirm{background:var(--info-soft);color:var(--info)}.qbrd .s-confirm{background:var(--info)}",
    ".qbrd .c-nodate{background:var(--idle-soft);color:var(--idle)}.qbrd .s-nodate{background:var(--idle)}",
    ".qbrd .c-scheduled{background:var(--accent-soft);color:var(--accent)}",
    ".qbrd .c-upcoming{background:var(--sunk);color:var(--muted)}",
    ".qbrd .c-completed{background:var(--ok-soft);color:var(--ok)}",
    ".qbrd .c-declined,.qbrd .c-asneeded{background:var(--idle-soft);color:var(--idle)}",
    ".qbrd .pri{font-size:10.5px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:var(--muted);border:1px solid var(--line);border-radius:4px;padding:0 5px}",
    ".qbrd .pri.Primary{color:var(--ink);border-color:var(--ink)}",
    ".qbrd .two{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.3fr);gap:14px}",
    "@media (max-width:760px){.qbrd .two{grid-template-columns:minmax(0,1fr)}.qbrd .qi{grid-template-columns:4px minmax(0,1fr)}.qbrd .acts{grid-column:2;justify-content:flex-start;padding-top:6px}}",
    ".qbrd .panel{border:1px solid var(--line);border-radius:8px;padding:14px;display:grid;gap:10px;align-content:start;min-width:0;background:var(--surface)}",
    ".qbrd .ph{display:flex;justify-content:space-between;align-items:baseline;gap:8px;flex-wrap:wrap}",
    ".qbrd .next{list-style:none;margin:0;padding:0}",
    ".qbrd .next li{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:1px 10px;padding:8px 0;border-top:1px solid var(--line)}",
    ".qbrd .next li:first-child{border-top:0}",
    ".qbrd .r{text-align:right}",
    ".qbrd svg{width:100%;height:auto;display:block}",
    ".qbrd .bar{cursor:pointer}.qbrd .bar:hover rect{opacity:.85}",
    ".qbrd .empty{padding:24px 16px;text-align:center;color:var(--muted)}",
    ".qbrd .err{padding:12px 16px;border:1px solid var(--crit);background:var(--crit-soft);color:var(--crit);border-radius:8px}",
    ".qbrd-scrim{position:fixed;inset:0;background:rgba(10,18,26,.4);z-index:100000}",
    ".qbrd-drawer{position:fixed;top:0;right:0;bottom:0;width:min(440px,100%);background:#fff;z-index:100001;border-left:1px solid #dde3e8;display:flex;flex-direction:column;font-family:'Segoe UI',system-ui,sans-serif;color:#1b2631;font-size:14px;box-shadow:-8px 0 24px rgba(0,0,0,.08)}",
    ".qbrd-drawer .btn{border:1px solid #dde3e8;background:#fff;color:#1b2631;border-radius:6px;padding:5px 12px;font:600 12.5px 'Segoe UI',system-ui,sans-serif;cursor:pointer}",
    ".qbrd-drawer .btn:hover{border-color:#5c6b78}",
    ".qbrd-drawer .btn.primary{background:#0a6a8f;border-color:#0a6a8f;color:#fff}",
    ".qbrd-drawer .btn.ghost{border-color:transparent;background:none;color:#5c6b78}",
    ".qbrd-drawer .btn[disabled]{opacity:.55;cursor:default}",
    ".qbrd-drawer header{padding:16px 20px 12px;border-bottom:1px solid #dde3e8;display:grid;gap:6px}",
    ".qbrd-drawer .body{padding:14px 20px;overflow:auto;display:grid;gap:14px;align-content:start}",
    ".qbrd-drawer .quick{display:grid;gap:8px;background:#f3f5f7;border-radius:8px;padding:12px}",
    ".qbrd-drawer .row{display:flex;flex-wrap:wrap;gap:6px;align-items:center}",
    ".qbrd-drawer .fields{display:grid;grid-template-columns:1fr 1fr;gap:10px}",
    ".qbrd-drawer label{display:grid;gap:3px;font-size:12px;color:#5c6b78}",
    ".qbrd-drawer label.full{grid-column:1/-1}",
    ".qbrd-drawer input,.qbrd-drawer select,.qbrd-drawer textarea{border:1px solid #dde3e8;border-radius:6px;padding:5px 8px;font:inherit;color:#1b2631;min-width:0;background:#fff}",
    ".qbrd-drawer textarea{min-height:90px;resize:vertical}",
    ".qbrd-drawer footer{padding:12px 20px;border-top:1px solid #dde3e8;display:flex;justify-content:space-between;gap:8px;align-items:center}",
    ".qbrd-drawer .copybox{width:100%;min-height:160px;font-family:Consolas,monospace;font-size:12px}",
    ".qbrd-toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#1b2631;color:#fff;padding:8px 14px;border-radius:8px;font:13px 'Segoe UI',sans-serif;z-index:100002}"
  ].join("\n");

  var STATUSES = ["Open", "Contacted", "No response", "Scheduled", "Completed", "Declined", "As needed"];
  var TIMING = {
    overdue: { label: "Overdue", cls: "overdue" }, soon: { label: "Due soon", cls: "soon" },
    confirm: { label: "Confirm it happened", cls: "confirm" }, nodate: { label: "Needs a due date", cls: "nodate" },
    scheduled: { label: "Scheduled", cls: "scheduled" }, upcoming: { label: "Upcoming", cls: "upcoming" },
    completed: { label: "Completed", cls: "completed" }, declined: { label: "Declined", cls: "declined" }, asneeded: { label: "As needed", cls: "asneeded" }
  };
  var NOTICE = ["overdue", "soon", "confirm", "nodate"];
  var PRI = { Primary: 0, Secondary: 1, "As Needed": 2 };
  var QEND = { Q1: [2, 31], Q2: [5, 30], Q3: [8, 30], Q4: [11, 31] };
  var FIELDS = "Id,Title,Priority,Frequency,PreferredWindow,QBRStatus,LastCompleted,PlannedDue,ScheduledDate,ManualDue,LastContact,Plan,QBRNotes,SPFolder,NextDue,Modified,Editor/Title";

  function pad(n) { return String(n).padStart(2, "0"); }
  function iso(d) { return d ? d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) : null; }
  function fromSP(v) { if (!v) return null; var d = new Date(v); if (isNaN(d)) return null; return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function parseISO(s) { if (!s) return null; var p = s.split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function toSP(s) { return s ? s + "T12:00:00Z" : null; }
  function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
  function edate(d, m) { var y = d.getFullYear(), mo = d.getMonth() + m, last = new Date(y, mo + 1, 0).getDate(); return new Date(y, mo, Math.min(d.getDate(), last)); }
  function dayDiff(a, b) { return Math.round((b - a) / 864e5); }
  function fmt(d) { return d ? d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—"; }
  function quarterOf(d) { return "Q" + (Math.floor(d.getMonth() / 3) + 1) + " " + d.getFullYear(); }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }

  function App(el, webUrl, listTitle) {
    this.el = el; this.web = webUrl.replace(/\/$/, ""); this.listTitle = listTitle;
    this.rows = []; this.tab = "all"; this.lead = 90; this.openId = null; this.listUrl = null; this.loading = true; this.error = null;
    try { var l = +window.localStorage.getItem("qbrd-lead"); if ([30, 60, 90].indexOf(l) >= 0) this.lead = l; } catch (e) { }
    this.today = (function () { var t = new Date(); return new Date(t.getFullYear(), t.getMonth(), t.getDate()); })();
  }
  App.prototype.api = function (path) { return this.web + "/_api/web/lists/getbytitle('" + encodeURIComponent(this.listTitle.replace(/'/g, "''")) + "')" + path; };
  App.prototype.digest = function () {
    var self = this;
    if (this._dg && this._dgExp > Date.now()) return Promise.resolve(this._dg);
    return fetch(this.web + "/_api/contextinfo", { method: "POST", credentials: "same-origin", headers: { Accept: "application/json;odata=nometadata" } })
      .then(function (r) { return r.json(); }).then(function (j) { self._dg = j.FormDigestValue; self._dgExp = Date.now() + (j.FormDigestTimeoutSeconds - 60) * 1000; return self._dg; });
  };
  App.prototype.load = function () {
    var self = this;
    var h = { Accept: "application/json;odata=nometadata" };
    var p1 = fetch(this.api("/items?$select=" + FIELDS + "&$expand=Editor&$top=1000"), { credentials: "same-origin", headers: h })
      .then(function (r) { if (!r.ok) throw new Error("Couldn't read the " + self.listTitle + " list (" + r.status + ")."); return r.json(); });
    var p2 = this.listUrl ? Promise.resolve(null) : fetch(this.api("/RootFolder?$select=ServerRelativeUrl"), { credentials: "same-origin", headers: h }).then(function (r) { return r.ok ? r.json() : null; });
    return Promise.all([p1, p2]).then(function (res) {
      if (res[1]) self.listUrl = res[1].ServerRelativeUrl;
      self.rows = res[0].value.map(function (x) {
        return {
          id: x.Id, client: x.Title || "(no name)", priority: x.Priority, freq: x.Frequency, window: x.PreferredWindow, status: x.QBRStatus || "Open",
          lastCompleted: fromSP(x.LastCompleted), plannedDue: fromSP(x.PlannedDue), scheduledDate: fromSP(x.ScheduledDate), manualDue: fromSP(x.ManualDue),
          lastContact: fromSP(x.LastContact), plan: x.Plan, notes: x.QBRNotes, folder: x.SPFolder, modified: x.Modified ? new Date(x.Modified) : null, editor: x.Editor && x.Editor.Title
        };
      });
      self.loading = false; self.error = null; self.render();
    }).catch(function (e) { self.loading = false; self.error = e.message || String(e); self.render(); });
  };
  App.prototype.save = function (id, patch, msg) {
    var self = this, body = {};
    Object.keys(patch).forEach(function (k) { body[k] = patch[k]; });
    return this.digest().then(function (dg) {
      return fetch(self.api("/items(" + id + ")"), {
        method: "POST", credentials: "same-origin",
        headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata", "X-RequestDigest": dg, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
        body: JSON.stringify(body)
      });
    }).then(function (r) {
      if (!r.ok) return r.text().then(function (t) { throw new Error(r.status === 403 ? "You don't have permission to edit this list." : "SharePoint didn't save the change (" + r.status + ")."); });
      self.toast(msg || "Saved"); return self.load();
    }).catch(function (e) { self.toast(e.message || "Couldn't save. Try again."); });
  };

  /* Same rule as the list's Next due column */
  App.prototype.nextDue = function (r) {
    if (r.manualDue) return r.manualDue;
    if (r.scheduledDate) return r.scheduledDate;
    var F = r.plannedDue, E = r.lastCompleted, C = Number(r.freq) || 0;
    if (F) { if (!E || F > E) return F; if (!C) return null; var m = (E.getFullYear() - F.getFullYear()); return new Date(F.getFullYear() + C * (Math.floor(m / C) + 1), F.getMonth(), F.getDate()); }
    if (E && C) return new Date(E.getFullYear() + C, E.getMonth(), E.getDate());
    return null;
  };
  App.prototype.timing = function (r) {
    var s = r.status, n = this.nextDue(r);
    if (s === "Completed") return "completed"; if (s === "Declined") return "declined"; if (s === "As needed") return "asneeded";
    if (s === "Scheduled") return r.scheduledDate && r.scheduledDate < this.today ? "confirm" : "scheduled";
    if (!n) return "nodate";
    if (n < this.today) return "overdue";
    if (n <= addDays(this.today, this.lead)) return "soon";
    return "upcoming";
  };
  App.prototype.rel = function (r) {
    var t = this.timing(r), n = this.nextDue(r);
    if (t === "confirm") return "Was scheduled " + fmt(r.scheduledDate);
    if (t === "scheduled") return "Meeting " + fmt(r.scheduledDate);
    if (!n) return "No due date set";
    var d = dayDiff(this.today, n);
    if (d < 0) return (-d) + " day" + (d === -1 ? "" : "s") + " overdue · due " + fmt(n);
    if (d === 0) return "Due today";
    return "Due in " + d + " days · " + fmt(n);
  };
  App.prototype.nextPlanned = function (r, doneISO) {
    var E = parseISO(doneISO), C = Number(r.freq) || 1, y = E.getFullYear() + C, q = String(r.window || "").match(/Q[1-4]/);
    if (q) return iso(new Date(y, QEND[q[0]][0], QEND[q[0]][1]));
    return iso(edate(E, 12 * C));
  };
  App.prototype.chip = function (t) { var m = TIMING[t]; return '<span class="chip c-' + m.cls + '">' + m.label + "</span>"; };
  App.prototype.pri = function (p) { return p ? '<span class="pri ' + esc(p) + '">' + esc(p) + "</span>" : ""; };
  App.prototype.sortNotice = function (a, b) {
    var ka = NOTICE.indexOf(this.timing(a)), kb = NOTICE.indexOf(this.timing(b)); if (ka !== kb) return ka - kb;
    var pa = PRI[a.priority] == null ? 3 : PRI[a.priority], pb = PRI[b.priority] == null ? 3 : PRI[b.priority]; if (pa !== pb) return pa - pb;
    var da = this.nextDue(a) || a.scheduledDate || new Date(9e12), db = this.nextDue(b) || b.scheduledDate || new Date(9e12); return da - db;
  };

  App.prototype.render = function () {
    var self = this, T = function (r) { return self.timing(r); };
    if (this.loading) { this.el.innerHTML = '<div class="qbrd"><div class="empty">Loading the QBR Tracker…</div></div>'; return; }
    if (this.error) { this.el.innerHTML = '<div class="qbrd"><div class="err">' + esc(this.error) + ' Check that a list named “' + esc(this.listTitle) + '” exists on this site.</div></div>'; return; }
    var rows = this.rows, due = rows.filter(function (r) { return NOTICE.indexOf(T(r)) >= 0; }).sort(function (a, b) { return self.sortNotice(a, b); });
    var cnt = {}; NOTICE.forEach(function (k) { cnt[k] = due.filter(function (r) { return T(r) === k; }).length; });
    var yr = this.today.getFullYear();
    var kpis = [
      ["Clients", rows.length, "", "all"],
      ["Overdue", cnt.overdue, "crit", "overdue"],
      ["Due in " + this.lead + " days", cnt.soon, "warn", "soon"],
      ["Scheduled", rows.filter(function (r) { return T(r) === "scheduled" || T(r) === "confirm"; }).length, "info", "confirm"],
      ["Completed in " + yr, rows.filter(function (r) { return r.lastCompleted && r.lastCompleted.getFullYear() === yr && r.lastCompleted <= self.today; }).length, "ok", null]
    ];
    var list = this.tab === "all" ? due : due.filter(function (r) { return T(r) === self.tab; });
    var tabs = [["all", "All", due.length]].concat(NOTICE.map(function (k) { return [k, TIMING[k].label, cnt[k]]; })).filter(function (t) { return t[0] === "all" || t[2]; });

    var h = '<div class="qbrd">';
    h += '<div class="top"><div><div class="eyebrow">Strategic Technology Reviews</div><h2>Review dashboard</h2><div class="muted">' +
      this.today.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" }) + "</div></div>";
    h += '<div class="ctrls"><span class="eyebrow">Notify when due within</span><div class="seg" role="group" aria-label="Notice lead time">' +
      [30, 60, 90].map(function (d) { return '<button data-lead="' + d + '" aria-pressed="' + (d === self.lead) + '">' + d + " days</button>"; }).join("") +
      '</div><button class="btn" data-act="refresh">Refresh</button><button class="btn primary" data-act="add">+ Add new client</button></div></div>';
    h += '<div class="kpis">' + kpis.map(function (k) {
      var tag = k[3] ? "button" : "div";
      return "<" + tag + ' class="kpi ' + k[2] + '"' + (k[3] ? ' data-tab="' + k[3] + '"' : "") + '><span class="v">' + k[1] + '</span><span class="l">' + k[0] + "</span></" + tag + ">";
    }).join("") + "</div>";

    // Next 5 + quarter chart
    var next = rows.filter(function (r) { var t = T(r); return t === "upcoming" || t === "scheduled"; }).sort(function (a, b) { return self.nextDue(a) - self.nextDue(b); }).slice(0, 5);
    h += '<div class="two"><div class="panel"><div class="ph"><h3>Next 5 reviews</h3><span class="eyebrow">Not yet due</span></div><ul class="next">' +
      (next.length ? next.map(function (r) {
        var n = self.nextDue(r);
        return '<li><span><button class="link" data-act="open" data-id="' + r.id + '" style="font-weight:600">' + esc(r.client) + "</button> " + self.pri(r.priority) + '</span><span class="r mono">' + fmt(n) + '</span><span class="muted" style="font-size:12.5px">' + TIMING[T(r)].label + (r.window ? " · " + esc(r.window) : "") + '</span><span class="muted r" style="font-size:12.5px">in ' + dayDiff(self.today, n) + " days</span></li>";
      }).join("") : '<li class="empty">No upcoming reviews.</li>') + "</ul></div>";
    h += '<div class="panel"><div class="ph"><h3>Open reviews by quarter</h3><span class="eyebrow">By next due date</span></div>' + this.chart() + "</div></div>";
    var nh = '';
    nh += '<section class="notice"><div class="nh"><div><h3>QBRs to schedule: ' + due.length + '</h3><div class="muted">' +
      (due.length ? cnt.overdue + " overdue, " + cnt.soon + " due in the next " + this.lead + " days" + (cnt.confirm ? ", " + cnt.confirm + " to confirm" : "") + (cnt.nodate ? ", " + cnt.nodate + " missing a date" : "") : "Nothing needs scheduling in the next " + this.lead + " days") +
      '</div></div><div class="tabs">' + tabs.map(function (t) { return '<button class="tab" data-tab="' + t[0] + '" aria-pressed="' + (self.tab === t[0]) + '">' + t[1] + " <b>" + t[2] + "</b></button>"; }).join("") + "</div></div>";
    nh += '<ul class="queue">' + (list.length ? list.map(function (r) {
      var t = T(r), acts;
      if (t === "confirm") acts = '<button class="btn primary" data-act="complete" data-id="' + r.id + '">Mark completed</button><button class="btn" data-act="open" data-id="' + r.id + '">Reschedule</button>';
      else acts = '<button class="btn" data-act="email" data-id="' + r.id + '">Copy email</button><button class="btn" data-act="contact" data-id="' + r.id + '">Log contact</button>' +
        '<input type="date" id="qbrd-sd-' + r.id + '" aria-label="Meeting date for ' + esc(r.client) + '"><button class="btn primary" data-act="schedule" data-id="' + r.id + '">Schedule</button>';
      return '<li class="qi"><span class="stripe s-' + TIMING[t].cls + '"></span><div class="qm"><div class="qn"><button class="link" data-act="open" data-id="' + r.id + '">' + esc(r.client) + "</button>" +
        self.chip(t) + self.pri(r.priority) + (r.status !== "Open" && t !== "confirm" ? '<span class="muted" style="font-size:12px;font-weight:500">' + esc(r.status) + "</span>" : "") + "</div>" +
        '<div class="meta"><span class="mono">' + self.rel(r) + "</span><span>" + (r.freq ? "Every " + (r.freq === 1 ? "year" : r.freq + " years") : "") + (r.window ? " · prefers " + esc(r.window) : "") + (r.plan ? " · " + esc(r.plan) + " plan" : "") +
        (r.lastContact ? " · last contact " + fmt(r.lastContact) : "") + "</span></div></div><div class=\"acts\">" + acts + "</div></li>";
    }).join("") : '<li class="empty">Every client is scheduled or not yet due.</li>') + "</ul></section>";

    h += nh;
    h += "</div>";
    this.el.innerHTML = h;
  };

  App.prototype.chart = function () {
    var self = this;
    var open = this.rows.filter(function (r) { return ["overdue", "soon", "upcoming", "scheduled", "confirm"].indexOf(self.timing(r)) >= 0 && self.nextDue(r); });
    var qs = new Date(this.today.getFullYear(), Math.floor(this.today.getMonth() / 3) * 3, 1);
    var b = [{ label: "Overdue", p: 0, s: 0, over: true }], later = 0, i;
    for (i = 0; i < 6; i++) b.push({ label: quarterOf(new Date(qs.getFullYear(), qs.getMonth() + i * 3, 1)), p: 0, s: 0 });
    open.forEach(function (r) {
      var t = self.timing(r), n = self.nextDue(r), k = t === "overdue" ? b[0] : b.filter(function (x) { return !x.over && x.label === quarterOf(n); })[0];
      if (!k) { later++; return; } if (r.priority === "Primary") k.p++; else k.s++;
    });
    var W = 560, H = 230, L = 34, B = 36, T = 16, R = 8, max = Math.max.apply(null, [4].concat(b.map(function (x) { return x.p + x.s; })));
    var step = max <= 8 ? 2 : max <= 20 ? 5 : 10, top = Math.ceil(max / step) * step, y = function (v) { return T + (H - T - B) * (1 - v / top); }, bw = (W - L - R) / b.length;
    var g = "", v;
    for (v = 0; v <= top; v += step) g += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="#dde3e8"/><text x="' + (L - 6) + '" y="' + (y(v) + 4) + '" text-anchor="end" font-size="11" fill="#5c6b78">' + v + "</text>";
    var bars = b.map(function (k, i) {
      var x = L + i * bw + bw * 0.18, w = bw * 0.64, yp = y(k.p), ys = y(k.p + k.s), pc = k.over ? "#b42318" : "#0a6a8f", sc = k.over ? "#fbe4e1" : "#dcedf4";
      var lbl = k.over ? "Overdue" : k.label.replace(/ (\d\d)(\d\d)$/, " ’$2");
      return '<g class="bar"' + (k.over ? ' data-tab="overdue"' : "") + ">" + (k.p ? '<rect x="' + x + '" y="' + yp + '" width="' + w + '" height="' + (y(0) - yp) + '" fill="' + pc + '" rx="2"/>' : "") +
        (k.s ? '<rect x="' + x + '" y="' + ys + '" width="' + w + '" height="' + (yp - ys) + '" fill="' + sc + '" stroke="' + pc + '" rx="2"/>' : "") +
        (k.p + k.s ? '<text x="' + (x + w / 2) + '" y="' + (ys - 5) + '" text-anchor="middle" font-size="12" font-weight="600" fill="#1b2631">' + (k.p + k.s) + "</text>" : "") +
        '<text x="' + (x + w / 2) + '" y="' + (H - B + 16) + '" text-anchor="middle" font-size="11" fill="' + (k.over ? "#b42318" : "#5c6b78") + '" font-weight="' + (k.over ? 600 : 400) + '">' + lbl + "</text></g>";
    }).join("");
    var lg = '<g font-size="11" fill="#5c6b78"><rect x="' + L + '" y="' + (H - 11) + '" width="10" height="10" fill="#0a6a8f" rx="2"/><text x="' + (L + 14) + '" y="' + (H - 2) + '">Primary</text>' +
      '<rect x="' + (L + 76) + '" y="' + (H - 11) + '" width="10" height="10" fill="#dcedf4" stroke="#0a6a8f" rx="2"/><text x="' + (L + 90) + '" y="' + (H - 2) + '">Secondary / other</text>' +
      (later ? '<text x="' + (W - R) + '" y="' + (H - 2) + '" text-anchor="end">+' + later + " later</text>" : "") + "</g>";
    return '<svg viewBox="0 0 ' + W + " " + H + '" role="img" aria-label="Open reviews by quarter">' + g + bars + lg + "</svg>";
  };

  App.prototype.emailText = function (r) {
    var q = String(r.window || "").match(/Q[1-4]/), when = q ? { Q1: "January–March", Q2: "April–June", Q3: "July–September", Q4: "October–December" }[q[0]] : "the next few weeks";
    var name = r.client.replace(/\*+/g, "").trim();
    return "Subject: Time to schedule your Strategic Technology Review\n\nHi,\n\nIt's time for " + name + "'s Strategic Technology Review with " + this.company + ". We'll walk through your current environment, security posture, hardware lifecycle and budget items for the year ahead. It usually takes about an hour.\n\nYour preferred window is " + when + ". Could you send a couple of dates and times that work for you?\n\nThanks,\n" + this.company;
  };
  App.prototype.toast = function (msg) {
    var t = document.querySelector(".qbrd-toast"); if (!t) { t = document.createElement("div"); t.className = "qbrd-toast"; document.body.appendChild(t); }
    t.textContent = msg; t.hidden = false; clearTimeout(this._tt); this._tt = setTimeout(function () { t.hidden = true; }, 2600);
  };
  App.prototype.find = function (id) { id = +id; return this.rows.filter(function (r) { return r.id === id; })[0]; };


  App.prototype.create = function (body, msg) {
    var self = this;
    return this.digest().then(function (dg) {
      return fetch(self.api("/items"), {
        method: "POST", credentials: "same-origin",
        headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata", "X-RequestDigest": dg },
        body: JSON.stringify(body)
      });
    }).then(function (r) {
      if (!r.ok) throw new Error(r.status === 403 ? "You don't have permission to add to this list." : "SharePoint didn't add the client (" + r.status + ").");
      self.toast(msg); return self.load();
    });
  };
  /* Suggested next due: end of the preferred quarter, frequency years after the last review; with no last review, the next upcoming preferred quarter */
  App.prototype.suggestDue = function (freq, win, lastISO) {
    var q = String(win || "").match(/Q[1-4]/), C = Number(freq) || 1, t = this.today;
    if (lastISO) {
      var E = parseISO(lastISO);
      if (q) return iso(new Date(E.getFullYear() + C, QEND[q[0]][0], QEND[q[0]][1]));
      return iso(edate(E, 12 * C));
    }
    if (q) { var d = new Date(t.getFullYear(), QEND[q[0]][0], QEND[q[0]][1]); if (d < t) d = new Date(t.getFullYear() + 1, QEND[q[0]][0], QEND[q[0]][1]); return iso(d); }
    return null;
  };
  App.prototype.openAdd = function () {
    var self = this; this.closeDrawer(); this.openId = -1;
    var plans = []; this.rows.forEach(function (r) { var p = (r.plan || "").trim(); if (p && plans.map(function (x) { return x.toLowerCase(); }).indexOf(p.toLowerCase()) < 0) plans.push(p); });
    ["Complete", "NFP/Complete", "Simple", "Basic"].forEach(function (p) { if (plans.indexOf(p) < 0) plans.push(p); });
    plans.sort();
    var opt = function (arr, v) { return arr.map(function (o) { var val = Array.isArray(o) ? o[0] : o, lab = Array.isArray(o) ? o[1] : o; return '<option value="' + esc(val) + '"' + (val === v ? " selected" : "") + ">" + esc(lab) + "</option>"; }).join(""); };
    var root = document.createElement("div"); root.id = "qbrd-drawer-root";
    root.innerHTML = '<div class="qbrd-scrim" data-qclose></div><aside class="qbrd-drawer" role="dialog" aria-modal="true" aria-label="Add new client">' +
      '<header><div class="row" style="justify-content:space-between"><h2 style="margin:0;font-size:20px">Add new client</h2><button class="btn ghost" data-qclose style="border:0;background:none;cursor:pointer;color:#5c6b78">Close</button></div>' +
      '<div style="color:#5c6b78;font-size:12.5px">Adds a row to the QBR Tracker list. Next due fills in from the quarter and frequency; change it if you need a different date.</div></header>' +
      '<form class="body" id="qbrd-add-form" novalidate><div class="fields">' +
      '<label class="full">Client name<input type="text" id="qbrd-n-client" required autocomplete="off" placeholder="e.g. Cascade Dental"></label>' +
      '<label>Priority<select id="qbrd-n-priority">' + opt(["Primary", "Secondary", "As Needed"], "Primary") + "</select></label>" +
      '<label>Review frequency<select id="qbrd-n-freq">' + opt([["1", "Every year"], ["2", "Every 2 years"], ["3", "Every 3 years"]], "1") + "</select></label>" +
      '<label>Quarter preference<select id="qbrd-n-window">' + opt([["", "No preference"], ["Q1", "Q1 (Jan–Mar)"], ["Q2", "Q2 (Apr–Jun)"], ["Q3", "Q3 (Jul–Sep)"], ["Q4", "Q4 (Oct–Dec)"]], "") + "</select></label>" +
      '<label>Plan type<select id="qbrd-n-plan"><option value="">Not set</option>' + opt(plans, "") + "</select></label>" +
      '<label>Last review date<input type="date" id="qbrd-n-last"></label>' +
      '<label>Next due date<input type="date" id="qbrd-n-due"></label>' +
      '<label>Status<select id="qbrd-n-status">' + opt(["Open", "Contacted", "Scheduled", "As needed"], "Open") + "</select></label>" +
      '<label>Scheduled meeting<input type="date" id="qbrd-n-sched"></label>' +
      '<label class="full">SharePoint folder<input type="text" id="qbrd-n-folder" autocomplete="off"></label>' +
      '<label class="full">Notes<textarea id="qbrd-n-notes"></textarea></label></div>' +
      '<div id="qbrd-n-err" style="color:#b42318;font-size:12.5px" hidden></div></form>' +
      '<footer><span style="font-size:12px;color:#5c6b78" id="qbrd-n-hint"></span><div class="row"><button class="btn" data-qclose>Cancel</button><button class="btn primary" data-qact="addsave">Add client</button></div></footer></aside>';
    document.body.appendChild(root);
    var $ = function (k) { return document.getElementById("qbrd-n-" + k); }, dueTouched = false;
    var refresh = function () {
      if (!dueTouched) { var s = self.suggestDue($("freq").value, $("window").value, $("last").value); $("due").value = s || ""; }
      var d = $("due").value; $("hint").textContent = d ? "Next due " + fmt(parseISO(d)) : "No due date yet";
    };
    ["freq", "window", "last"].forEach(function (k) { $(k).addEventListener("change", refresh); });
    $("due").addEventListener("input", function () { dueTouched = !!$("due").value; refresh(); });
    $("status").addEventListener("change", function () { if ($("status").value === "Scheduled") $("sched").focus(); });
    $("priority").addEventListener("change", function () { if ($("priority").value === "As Needed" && $("freq").value === "1") { $("freq").value = "3"; refresh(); } });
    refresh(); $("client").focus();
    var submit = function () {
      var err = $("err"), name = $("client").value.trim();
      var fail = function (m) { err.textContent = m; err.hidden = false; };
      if (!name) { fail("Enter the client's name."); $("client").focus(); return; }
      if (self.rows.some(function (r) { return r.client.toLowerCase() === name.toLowerCase(); })) { fail("A client named “" + name + "” is already in the tracker."); return; }
      if ($("status").value === "Scheduled" && !$("sched").value) { fail("Pick the scheduled meeting date, or change Status."); return; }
      err.hidden = true;
      var body = { Title: name, Priority: $("priority").value, Frequency: Number($("freq").value), QBRStatus: $("status").value };
      if ($("window").value) body.PreferredWindow = $("window").value;
      if ($("plan").value) body.Plan = $("plan").value;
      if ($("last").value) body.LastCompleted = toSP($("last").value);
      if ($("due").value) body.PlannedDue = toSP($("due").value);
      if ($("sched").value) body.ScheduledDate = toSP($("sched").value);
      if ($("folder").value.trim()) body.SPFolder = $("folder").value.trim();
      body.QBRNotes = ($("notes").value.trim() ? $("notes").value.trim() + " | " : "") + "Added from QBR Board " + iso(self.today);
      var btn = root.querySelector('[data-qact="addsave"]'); btn.disabled = true; btn.textContent = "Adding…";
      self.create(body, "Added " + name).then(function () { self.closeDrawer(); }, function (e) { btn.disabled = false; btn.textContent = "Add client"; fail(e.message || "Couldn't add the client. Try again."); });
    };
    root.addEventListener("click", function (e) {
      if (e.target.closest("[data-qclose]")) { self.closeDrawer(); return; }
      if (e.target.closest('[data-qact="addsave"]')) submit();
    });
    $("client").form.addEventListener("submit", function (e) { e.preventDefault(); submit(); });
  };

  App.prototype.closeDrawer = function () { var d = document.getElementById("qbrd-drawer-root"); if (d) d.remove(); this.openId = null; };
  App.prototype.openDrawer = function (id, emailText) {
    var self = this, r = this.find(id); if (!r) return; this.closeDrawer(); this.openId = r.id;
    var opt = function (arr, v) { return arr.map(function (o) { return "<option" + (o === v ? " selected" : "") + ">" + esc(o) + "</option>"; }).join(""); };
    var di = function (k, label, val) { return "<label>" + label + '<input type="date" id="qbrd-e-' + k + '" value="' + (val ? iso(val) : "") + '"></label>'; };
    var heldDefault = r.scheduledDate && r.scheduledDate <= this.today ? iso(r.scheduledDate) : iso(this.today);
    var root = document.createElement("div"); root.id = "qbrd-drawer-root";
    root.innerHTML = '<div class="qbrd-scrim" data-qclose></div><aside class="qbrd-drawer" role="dialog" aria-modal="true" aria-label="' + esc(r.client) + '">' +
      '<header><div class="row" style="justify-content:space-between"><h2 style="margin:0;font-size:20px">' + esc(r.client) + '</h2><button class="btn ghost" data-qclose style="border:0;background:none;cursor:pointer;color:#5c6b78">Close</button></div>' +
      "<div>" + this.chip(this.timing(r)).replace('class="chip', 'style="display:inline-flex;gap:5px;border-radius:999px;padding:1px 8px;font-size:11.5px;font-weight:600" class="chip') + ' <span style="color:#5c6b78">' + this.rel(r) + "</span></div></header>" +
      '<div class="body">' +
      (emailText ? '<div><div style="font-size:12px;color:#5c6b78;margin-bottom:6px">Scheduling email. Select all and copy.</div><textarea class="copybox" readonly id="qbrd-copybox">' + esc(emailText) + "</textarea></div>" : "") +
      '<div class="quick"><div style="font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#5c6b78">Quick actions</div>' +
      '<div class="row"><button class="btn" data-qact="email">Copy email</button><button class="btn" data-qact="contact">Log contact today</button><button class="btn" data-qact="noresp">No response</button></div>' +
      '<div class="row"><span style="font-size:12px;color:#5c6b78">Meeting</span><input type="date" id="qbrd-d-sched" value="' + (r.scheduledDate ? iso(r.scheduledDate) : "") + '"><button class="btn primary" data-qact="schedule">Schedule</button></div>' +
      '<div class="row"><span style="font-size:12px;color:#5c6b78">Held on</span><input type="date" id="qbrd-d-done" value="' + heldDefault + '"><button class="btn" data-qact="complete">Mark completed</button>' +
      (r.status === "Declined" || r.status === "Completed" ? '<button class="btn ghost" data-qact="reopen">Reopen</button>' : '<button class="btn ghost" data-qact="decline">Client declined</button>') + "</div></div>" +
      '<div class="fields"><label>Status<select id="qbrd-e-status">' + opt(STATUSES, r.status) + '</select></label><label>Priority<select id="qbrd-e-priority">' + opt(["Primary", "Secondary", "As Needed"], r.priority) + "</select></label>" +
      '<label>Every (years)<select id="qbrd-e-freq">' + opt(["1", "2", "3"], String(r.freq || "")) + '</select></label><label>Preferred window<select id="qbrd-e-window"><option value=""></option>' + opt(["Q1", "Q2", "Q3", "Q4"], r.window) + "</select></label>" +
      di("lastCompleted", "Last completed", r.lastCompleted) + di("plannedDue", "Planned due", r.plannedDue) + di("scheduledDate", "Scheduled date", r.scheduledDate) + di("manualDue", "Manual due (override)", r.manualDue) + di("lastContact", "Last contact", r.lastContact) +
      '<label>Plan<input type="text" id="qbrd-e-plan" value="' + esc(r.plan || "") + '"></label><label class="full">Notes<textarea id="qbrd-e-notes">' + esc(r.notes || "") + "</textarea></label></div>" +
      '<div style="font-size:12px;color:#5c6b78">' + (r.modified ? "Last changed " + fmt(r.modified) + (r.editor ? " by " + esc(r.editor) : "") + ". " : "") +
      (this.listUrl ? '<a href="' + esc(this.listUrl) + "/DispForm.aspx?ID=" + r.id + '" target="_blank" rel="noopener">Open in the list</a>' : "") + "</div></div>" +
      '<footer><span></span><div class="row"><button class="btn" data-qclose>Cancel</button><button class="btn primary" data-qact="save">Save changes</button></div></footer></aside>';
    document.body.appendChild(root);
    root.addEventListener("click", function (e) {
      var c = e.target.closest("[data-qclose]"); if (c) { self.closeDrawer(); return; }
      var b = e.target.closest("[data-qact]"); if (!b) return;
      var a = b.getAttribute("data-qact"), get = function (k) { var el = document.getElementById("qbrd-e-" + k); return el && el.value ? el.value : null; };
      if (a === "email") { self.copyEmail(r.id); return; }
      if (a === "save") {
        self.save(r.id, {
          QBRStatus: get("status"), Priority: get("priority"), Frequency: get("freq") ? Number(get("freq")) : null, PreferredWindow: get("window"),
          LastCompleted: toSP(get("lastCompleted")), PlannedDue: toSP(get("plannedDue")), ScheduledDate: toSP(get("scheduledDate")), ManualDue: toSP(get("manualDue")),
          LastContact: toSP(get("lastContact")), Plan: get("plan"), QBRNotes: document.getElementById("qbrd-e-notes").value
        }); self.closeDrawer(); return;
      }
      if (a === "schedule") { var sd = document.getElementById("qbrd-d-sched").value; if (!sd) { self.toast("Pick a meeting date first"); return; } self.act("schedule", r.id, sd); self.closeDrawer(); return; }
      if (a === "complete") { self.act("complete", r.id, document.getElementById("qbrd-d-done").value); self.closeDrawer(); return; }
      self.act(a, r.id); self.closeDrawer();
    });
    var cb = document.getElementById("qbrd-copybox"); if (cb) { cb.focus(); cb.select(); }
  };
  App.prototype.copyEmail = function (id) {
    var self = this, r = this.find(id), txt = this.emailText(r);
    var fallback = function () { self.openDrawer(id, txt); };
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(txt).then(function () { self.toast("Scheduling email copied"); }, fallback); else fallback();
  };
  App.prototype.act = function (a, id, dateISO) {
    var r = this.find(id), today = iso(this.today);
    if (a === "contact") return this.save(id, { QBRStatus: "Contacted", LastContact: toSP(today) }, "Contact logged");
    if (a === "noresp") return this.save(id, { QBRStatus: "No response", LastContact: toSP(today) }, "Marked no response");
    if (a === "schedule") return this.save(id, { QBRStatus: "Scheduled", ScheduledDate: toSP(dateISO), ManualDue: null }, "Scheduled for " + fmt(parseISO(dateISO)));
    if (a === "complete") {
      var d = dateISO || (r.scheduledDate && r.scheduledDate <= this.today ? iso(r.scheduledDate) : today);
      return this.save(id, { QBRStatus: "Open", LastCompleted: toSP(d), PlannedDue: toSP(this.nextPlanned(r, d)), ScheduledDate: null, ManualDue: null, LastContact: null }, "Completed. Next review due " + fmt(parseISO(this.nextPlanned(r, d))));
    }
    if (a === "decline") return this.save(id, { QBRStatus: "Declined" }, "Marked declined");
    if (a === "reopen") return this.save(id, { QBRStatus: "Open" }, "Reopened");
  };
  App.prototype.bind = function () {
    var self = this;
    this.el.addEventListener("click", function (e) {
      var b = e.target.closest("[data-lead],[data-tab],[data-act]"); if (!b || !self.el.contains(b)) return;
      if (b.hasAttribute("data-lead")) { self.lead = +b.getAttribute("data-lead"); try { window.localStorage.setItem("qbrd-lead", self.lead); } catch (er) { } self.render(); return; }
      if (b.hasAttribute("data-tab")) { self.tab = b.getAttribute("data-tab"); self.render(); var q = self.el.querySelector(".notice"); if (q && (b.classList.contains("kpi") || b.classList.contains("bar"))) q.scrollIntoView({ behavior: "smooth", block: "start" }); return; }
      var a = b.getAttribute("data-act"), id = b.getAttribute("data-id");
      if (a === "refresh") { self.load(); return; }
      if (a === "add") { self.openAdd(); return; }
      if (a === "open") { self.openDrawer(id); return; }
      if (a === "email") { self.copyEmail(id); return; }
      if (a === "schedule") { var inp = document.getElementById("qbrd-sd-" + id); if (!inp || !inp.value) { self.toast("Pick a meeting date first"); return; } self.act("schedule", id, inp.value); return; }
      self.act(a, id);
    });
    document.addEventListener("keydown", this._esc = function (e) { if (e.key === "Escape" && self.openId) self.closeDrawer(); });
  };

  class QbrDashboardWebPart extends Base {
    render() {
      if (!document.getElementById("qbrd-css")) { var s = document.createElement("style"); s.id = "qbrd-css"; s.textContent = CSS; document.head.appendChild(s); }
      if (!this._app) {
        this._app = new App(this.domElement, this.context.pageContext.web.absoluteUrl, (this.properties && this.properties.listTitle) || "QBR Tracker");
        this._app.company = (this.properties && this.properties.companyName) || "your IT team";
        this._app.bind(); this._app.render(); this._app.load();
      } else { this._app.render(); }
    }
    onDispose() { if (this._app) { this._app.closeDrawer(); if (this._app._esc) document.removeEventListener("keydown", this._app._esc); } }
  }
  return { __esModule: true, default: QbrDashboardWebPart };
});
