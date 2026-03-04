"use strict";

const API = "/api";

const SHEET_MAP = {
  overview: "profiles",
  employees: "profiles",
  engagements: "Engagements",
  sites: "Sites_List",
  schedule: "Seti_Updated",
  gantt: "Seti_Updated",
};

const ID_PREFIX_MAP = {
  employees: "EMP",
  engagements: "ENG",
  sites: "SITE",
};

const store = {
  view: "overview",
  page: 1,
  pageSize: 10,
  globalSearch: "",
  columnFilters: {},
  sortCol: null,
  sortDir: "asc",
  data: [],
  allEmployees: [],
  allEngagements: [],
  allSites: [],
  allSchedule: [],
};
const SITE_STATUS_OPTIONS = ["Active", "Completed", "TBD"];

// ================================================================
// AUTH
// ================================================================
function showLogin() {
  const loginShell = document.getElementById("loginShell");
  const appShell = document.getElementById("appShell");
  if (loginShell) loginShell.style.display = "grid";
  if (appShell) appShell.style.display = "none";
  const btn = document.getElementById("btnLogout");
  if (btn) btn.style.display = "none";
}

function showApp() {
  const loginShell = document.getElementById("loginShell");
  const appShell = document.getElementById("appShell");
  if (loginShell) loginShell.style.display = "none";
  if (appShell) appShell.style.display = "";
  const btn = document.getElementById("btnLogout");
  if (btn) btn.style.display = "";
}

async function apiFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  if (
    options.body &&
    !(options.body instanceof FormData) &&
    !headers.has("Content-Type")
  ) {
    headers.set("Content-Type", "application/json");
  }

  const res = await fetch(url, {
    ...options,
    headers,
    credentials: "include",
  });

  if (res.status === 401) {
    showLogin();
  }
  return res;
}

async function apiLogin(email, password) {
  const res = await apiFetch(`${API}/auth/login`, {
    method: "POST",
    body: JSON.stringify({ email, password }),
  });
  const json = await res.json().catch(() => ({}));
  if (!json.ok) throw new Error(json.error || "Login failed");
  return json;
}

async function apiLogout() {
  await apiFetch(`${API}/auth/logout`, { method: "POST" }).catch(() => {});
}

async function apiMe() {
  const res = await apiFetch(`${API}/auth/me`);
  const json = await res.json().catch(() => ({}));
  if (!json.ok) return null;
  return json;
}

async function boot() {
  const logoutBtn = document.getElementById("btnLogout");
  if (logoutBtn) {
    logoutBtn.onclick = async () => {
      await apiLogout();
      showLogin();
      toast("info", "Logged out");
    };
  }

  showLogin();

  const me = await apiMe();
  if (!me?.ok) {
    window.location.href = "/login";
    return;
  }

  const form = document.getElementById("loginForm");
  const err = document.getElementById("loginError");
  const emailEl = document.getElementById("loginEmail");
  const passEl = document.getElementById("loginPassword");

  if (!form) return;

  const setErr = (msg) => {
    if (!err) return;
    err.textContent = msg || "";
    err.style.display = msg ? "" : "none";
  };

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    setErr("");

    const email = String(emailEl?.value || "").trim();
    const password = String(passEl?.value || "");

    if (!email || !password) {
      setErr("Enter your email and password.");
      return;
    }

    try {
      await apiLogin(email, password);
      showApp();
      toast("success", "Welcome back");
      await loadAllData();
      const v = getViewFromHash() || DEFAULT_VIEW;
      await navigate(v, { pushHash: true });
    } catch (e) {
      setErr(e.message || "Login failed");
      toast("error", "Login failed", e.message || "");
    }
  });
}

document.addEventListener("DOMContentLoaded", boot);

function findKey(obj, re) {
  return Object.keys(obj || {}).find((k) => re.test(String(k)));
}

function normalizeName(s) {
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function buildEmployeeIndex() {
  const emps = getEmployeesForPicker();
  const byName = new Map();
  for (const e of emps) {
    if (!e.id || !e.name) continue;
    byName.set(normalizeName(e.name), e.id);
  }
  return { byName };
}

function enrichEngagementRowsWithEmployeeId(rows) {
  if (!Array.isArray(rows) || !rows.length) return rows;

  const sample = rows[0] || {};
  const empNameKey =
    findKey(sample, /employee\s*name/i) || findKey(sample, /\bname\b/i);

  const empIdKey = findKey(sample, /employee\s*id/i) || "Employee ID";

  if (!empNameKey) return rows;

  const { byName } = buildEmployeeIndex();

  return rows.map((r) => {
    const cur = String(r[empIdKey] ?? "").trim();
    if (cur) return r;

    const name = normalizeName(r[empNameKey]);
    const id = byName.get(name) || "";
    if (!id) return r;

    return { ...r, [empIdKey]: id };
  });
}

const views = {
  overview: { title: "Dashboard Overview", columns: [] },
  employees: { title: "Employees", columns: [] },
  engagements: { title: "Engagements", columns: [] },
  sites: { title: "Sites", columns: [] },
  schedule: { title: "Schedule", columns: [] },
  gantt: { title: "Gantt", columns: [] },
};

let chartInstances = {};

const el = (id) => document.getElementById(id);
const qsa = (sel) => [...document.querySelectorAll(sel)];

const esc = (str) =>
  String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

let __toastMounted = false;

function tableEls() {
  const thead = document.getElementById("theadDefault");
  const tbody = document.getElementById("tbodyDefault");
  return { thead, tbody };
}

function applyViewLayout() {
  const isOverview = store.view === "overview";
  const isSchedule = store.view === "schedule";
  const isGantt = store.view === "gantt";

  const overviewView = document.getElementById("overviewView");
  const tableView = document.getElementById("tableView");

  const stats = document.getElementById("stats");
  const filterBar = document.getElementById("filterBar");
  const advancedFilters = document.getElementById("advancedFilters");
  const paginationBar = document.getElementById("paginationBar");
  const rowsWrap = document.querySelector(".page-size-wrap");

  const scheduleHeader = document.getElementById("scheduleHeader");
  const scheduleLayout = document.getElementById("scheduleLayout");
  const defaultWrap = document.getElementById("defaultTableWrap");

  const tableCard = document.querySelector("#tableView .table-card");

  // main sections
  if (overviewView) overviewView.style.display = isOverview ? "" : "none";
  if (tableView) tableView.style.display = isOverview ? "none" : "";

  // table card MUST stay visible for schedule + gantt (because scheduleLayout is inside it)
  if (tableCard) tableCard.style.display = isOverview ? "none" : "";

  // schedule header + layout visible for schedule and gantt
  if (scheduleHeader) scheduleHeader.style.display = isSchedule ? "" : "none";
  if (scheduleLayout)
    scheduleLayout.style.display = isSchedule || isGantt ? "" : "none";

  // default table visible only for non-schedule views
  if (defaultWrap)
    defaultWrap.style.display =
      !isOverview && !isSchedule && !isGantt ? "" : "none";

  // gantt page should be “full” -> hide filters/stats/pagination/rows selector
  if (stats) stats.style.display = isGantt ? "none" : "";
  if (filterBar) filterBar.style.display = isGantt ? "none" : "";
  if (advancedFilters) advancedFilters.style.display = "none";
  if (paginationBar) paginationBar.style.display = isGantt ? "none" : "";
  if (rowsWrap) rowsWrap.style.display = isGantt ? "none" : "";
}

function mountScheduleGanttButton() {
  const btn = document.getElementById("btnOpenGantt");
  if (!btn) return;

  btn.style.display = store.view === "schedule" ? "" : "none";
  btn.onclick = async () => {
    store.view = "gantt";
    applyViewLayout();
    await loadCurrentViewData();
    await renderGanttPage();
  };
}

function wireGanttControls() {
  const back = document.getElementById("btnBackToSchedule");
  if (back) {
    back.onclick = async () => {
      store.view = "schedule";
      applyViewLayout();
      await loadCurrentViewData();
      await renderTableView();
    };
  }

  const zoom = document.getElementById("ganttZoomSelect");
  if (zoom) zoom.onchange = () => renderGanttPage();
}

function renderGanttPage() {
  // Gantt uses the PDF-like sheet that mounts in #scheduleSheet
  if (!store.allSchedule || !store.allSchedule.length) {
    const mount = document.getElementById("scheduleSheet");
    if (mount) {
      mount.innerHTML = `
        <div class="gantt-empty" style="padding:18px">
          <div style="font-weight:900">No data available</div>
          <div style="color:#60708a;margin-top:4px">Load schedule data first.</div>
        </div>`;
    }
    return;
  }
  const zoomSel = document.getElementById("ganttZoom");
  store.ganttCellW = Number(zoomSel?.value || store.ganttCellW || 22);

  renderScheduleSheet();
}

function mountToaster() {
  if (__toastMounted) return;
  __toastMounted = true;

  const style = document.createElement("style");
  style.textContent = `
    .toast-wrap {
      position:fixed; right:16px; top:16px; z-index:9999;
      display:flex; flex-direction:column; gap:10px; width:min(360px, calc(100vw - 32px));
      pointer-events:none;
    }
    .toast {
      pointer-events:auto;
      border:1px solid rgba(0,0,0,.08);
      background:#ffffff;
      box-shadow:0 12px 30px rgba(0,0,0,.14);
      border-radius:14px;
      padding:12px 16px;
      display:flex; gap:12px; align-items:flex-start;
      transform: translateY(-6px);
      opacity:0;
      transition: all .22s ease;
    }
    .toast.show { transform: translateY(0); opacity:1; }
    .toast .ico { width:24px; height:24px; display:grid; place-items:center; border-radius:8px; font-size:14px; flex:0 0 auto; }
    .toast .msg { font: 13.5px/1.4 system-ui, -apple-system, Segoe UI, Roboto; color:#1a1916; flex:1; }
    .toast .sub { font-size:12px; color:#6b6964; margin-top:3px; }
    .toast .x {
      margin-left:auto; border:0; background:transparent; cursor:pointer;
      color:#8a8883; font-size:17px; line-height:1; padding:2px 6px;
    }
    .toast.success .ico { background:rgba(45,122,79,.14); color:#2d7a4f; }
    .toast.error   .ico { background:rgba(192,57,43,.12); color:#c0392b; }
    .toast.info    .ico { background:rgba(122,119,114,.14); color:#4a4844; }
  `;
  document.head.appendChild(style);

  const wrap = document.createElement("div");
  wrap.className = "toast-wrap";
  wrap.id = "toastWrap";
  document.body.appendChild(wrap);
}
// ==============================
// NAVIGATION
// ==============================

document.querySelectorAll(".nav-link").forEach((btn) => {
  btn.addEventListener("click", async () => {
    const view = btn.dataset.view;
    await setView(view);
  });
});

function toast(type, message, sub = "", ttl = 2800) {
  mountToaster();
  const wrap = el("toastWrap");
  if (!wrap) return;

  const t = document.createElement("div");
  t.className = `toast ${type || "info"}`;
  const ico = type === "success" ? "✓" : type === "error" ? "!" : "i";

  t.innerHTML = `
    <div class="ico">${ico}</div>
    <div class="msg">
      <div>${esc(message)}</div>
      ${sub ? `<div class="sub">${esc(sub)}</div>` : ""}
    </div>
    <button class="x" title="Close">×</button>
  `;

  const remove = () => {
    t.classList.remove("show");
    setTimeout(() => t.remove(), 220);
  };

  t.querySelector(".x").onclick = remove;
  wrap.appendChild(t);
  requestAnimationFrame(() => t.classList.add("show"));
  if (ttl > 0) setTimeout(remove, ttl);
}

window.addEventListener("hashchange", () => {
  const v = getViewFromHash() || DEFAULT_VIEW;
  navigate(v, { pushHash: false });
});

function isStatusKey(key) {
  const k = String(key || "").toLowerCase();
  return k.includes("status") || k.includes("phase");
}

function statusBadgeHTML(value) {
  const raw = String(value ?? "").trim();
  const v = raw
    .toLowerCase()
    .replace(/\./g, "")
    .replace(/-/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const map = [
    {
      cls: "active",
      keys: ["active", "in progress", "in-progress", "ongoing", "running"],
    },
    {
      cls: "completed",
      keys: ["completed", "complete", "done", "finished", "closed"],
    },
    {
      cls: "pending",
      keys: [
        "pending",
        "waiting",
        "queued",
        "not started",
        "not-started",
        "todo",
        "to do",
        "to be decided",
        "tbd",
      ],
    },
    { cls: "hold", keys: ["on hold", "hold", "paused", "blocked", "stuck"] },
    {
      cls: "cancelled",
      keys: ["cancelled", "canceled", "rejected", "failed", "stopped"],
    },
  ];
  const hit = map.find((m) => m.keys.includes(v));
  const cls = hit ? hit.cls : "unknown";
  const label = raw ? raw.replace(/\s+/g, " ").trim() : "Unknown";

  return `<span class="badge ${cls}"><span class="badge-dot"></span>${esc(label)}</span>`;
}

function isSiteKey(key) {
  const k = String(key || "").toLowerCase();
  return k.includes("site") || k.includes("location");
}

function isEmployeeNameKey(key) {
  const k = String(key || "").toLowerCase();
  return k.includes("employee") && k.includes("name");
}

function getSiteNames() {
  const rows = store.allSites || [];
  if (!rows.length) return [];

  const sample = rows[0] || {};
  const nameKey =
    Object.keys(sample).find((k) =>
      /site.*name|sitename|location|name|title/i.test(k),
    ) || "id";

  const names = rows
    .map((r) => String(r[nameKey] ?? "").trim())
    .filter(Boolean);

  return Array.from(new Set(names)).sort((a, b) => a.localeCompare(b));
}

function getEmployeesForPicker() {
  const rows = store.allEmployees || [];
  if (!rows.length) return [];

  const sample = rows[0] || {};
  const idKey = Object.keys(sample).find((k) => /^id$/i.test(k)) || "id";
  const nameKey =
    Object.keys(sample).find((k) => /employee\s*name/i.test(k)) ||
    Object.keys(sample).find((k) => /^name$/i.test(k)) ||
    Object.keys(sample).find((k) => /name/i.test(k)) ||
    null;
  const desigKey =
    Object.keys(sample).find((k) => /designation/i.test(k)) || null;

  const items = rows
    .map((r) => {
      const id = String(r[idKey] ?? "").trim();
      const name = nameKey ? String(r[nameKey] ?? "").trim() : "";
      const desig = desigKey ? String(r[desigKey] ?? "").trim() : "";
      return { id, name, desig };
    })
    .filter((x) => x.id && x.name);

  items.sort((a, b) => a.name.localeCompare(b.name));
  return items;
}

function resolveEmployeeFromInput(inputText) {
  const val = String(inputText || "").trim();
  if (!val) return null;

  const employees = getEmployeesForPicker();

  const byId = employees.find((e) => e.id.toLowerCase() === val.toLowerCase());
  if (byId) return byId;

  const byName = employees.find(
    (e) => e.name.toLowerCase() === val.toLowerCase(),
  );
  if (byName) return byName;

  const loose = employees.find((e) =>
    e.name.toLowerCase().includes(val.toLowerCase()),
  );
  return loose || null;
}

// ================================================================
// API CALLS
// ================================================================

async function fetchAllRows(sheet) {
  const pageSize = 500;
  let page = 1;
  let all = [];

  while (true) {
    const res = await fetch(
      `${API}/rows?sheet=${encodeURIComponent(sheet)}&page=${page}&pageSize=${pageSize}`,
    );
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || "Failed to load rows");

    const items = json.items || [];
    all = all.concat(items);
    if (items.length < pageSize) break;
    page++;
  }

  return all;
}

async function apiCreateRow(sheet, view, data) {
  const idPrefix = view === "schedule" ? "" : (ID_PREFIX_MAP[view] || "ROW");
  const res = await fetch(
    `${API}/rows?sheet=${encodeURIComponent(sheet)}&idPrefix=${encodeURIComponent(idPrefix)}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    },
  );
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Create failed");
  return json.item;
}

async function apiUpdateRow(sheet, id, patch) {
  const key = sheet === "Sites_List" ? "Location" : "id";

  const res = await fetch(
    `${API}/rows/${encodeURIComponent(id)}?sheet=${encodeURIComponent(sheet)}&key=${encodeURIComponent(key)}`,
    {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(patch),
    },
  );

  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Update failed");
  return json.item;
}

async function apiDeleteRow(sheet, id) {
  const res = await fetch(
    `${API}/rows/${encodeURIComponent(id)}?sheet=${encodeURIComponent(sheet)}`,
    { method: "DELETE" },
  );
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Delete failed");
  return true;
}
function renderGantt(container, rows, opts = {}) {
  container.innerHTML = "";

  const dayWidth = Number(opts.dayW || 18);
  const rowHeight = 38;
  const padDays = Number(opts.padDays ?? 14);

  const startOfDay = (d) => {
    const x = new Date(d);
    x.setHours(0, 0, 0, 0);
    return x;
  };

  const parse = (v) => {
    if (!v) return null;
    const d = new Date(v);
    return isNaN(d) ? null : startOfDay(d);
  };

  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

  const toISO = (d) => d.toISOString().slice(0, 10);

  // ISO week number (to match PDF style W4, W5...)
  const isoWeek = (date) => {
    const d = new Date(
      Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()),
    );
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  };

  const tasks = (Array.isArray(rows) ? rows : [])
    .map((r, i) => {
      const s = parse(r.Start);
      const f = parse(r.Finish);
      const durStr = String(r.Duration ?? "");
      const pct =
        parseFloat(String(r["% Complete"] ?? 0).replace("%", "")) || 0;

      // milestone if start==finish AND duration looks like 0 days (safe)
      const isMilestone = !!(s && f && +s === +f && /\b0\b/.test(durStr));

      // summary heuristic (optional)
      const name = String(r["Task Name"] ?? "");
      const isSummary = /scope|project|procurement|design|powerhouse/i.test(
        name,
      );

      return {
        id: r.ID || i + 1,
        name,
        start: s,
        finish: f,
        pct: clamp(pct, 0, 100),
        duration: durStr,
        predecessors: r.Predecessors || "",
        isSummary,
        isMilestone,
      };
    })
    .filter((t) => t.start && t.finish);

  if (!tasks.length) {
    container.innerHTML =
      "<div style='padding:18px;font-weight:800'>No schedule data</div>";
    return;
  }

  // range
  let min = tasks.reduce((a, b) => (a < b.start ? a : b.start), tasks[0].start);
  let max = tasks.reduce(
    (a, b) => (a > b.finish ? a : b.finish),
    tasks[0].finish,
  );

  min = new Date(min);
  min.setDate(min.getDate() - padDays);
  min = startOfDay(min);

  max = new Date(max);
  max.setDate(max.getDate() + padDays);
  max = startOfDay(max);

  const totalDays = Math.max(1, Math.ceil((max - min) / 86400000) + 1);
  const timelineWidth = totalDays * dayWidth;

  const addDays = (d, n) => {
    const x = new Date(d);
    x.setDate(x.getDate() + n);
    return x;
  };

  // --- wrapper
  const wrap = document.createElement("div");
  wrap.className = "gantt-container";

  // ================= LEFT TABLE =================
  const left = document.createElement("div");
  left.className = "gantt-left";

  left.innerHTML = `
    <div class="gantt-left-head">
      <div>ID</div>
      <div>Task Name</div>
      <div>Start</div>
      <div>Finish</div>
      <div>%</div>
    </div>
  `;

  const leftBody = document.createElement("div");
  leftBody.className = "gantt-left-body";

  tasks.forEach((t) => {
    const row = document.createElement("div");
    row.className = "gantt-left-row";
    row.style.height = rowHeight + "px";

    row.innerHTML = `
      <div>${t.id}</div>
      <div class="gantt-taskname">${t.name}</div>
      <div>${toISO(t.start)}</div>
      <div>${toISO(t.finish)}</div>
      <div>${Math.round(t.pct)}%</div>
    `;
    leftBody.appendChild(row);
  });

  left.appendChild(leftBody);

  // ================= RIGHT SIDE =================
  const rightOuter = document.createElement("div");
  rightOuter.className = "gantt-right-outer";

  // --- sticky timeline header (YEAR + WEEKS)
  const head = document.createElement("div");
  head.className = "gantt-right-head";

  // YEAR row (like PDF “2025”)
  const yearsRow = document.createElement("div");
  yearsRow.className = "gantt-head-years";
  yearsRow.style.width = timelineWidth + "px";

  // WEEKS row (W4, W5...)
  const weeksRow = document.createElement("div");
  weeksRow.className = "gantt-head-weeks";
  weeksRow.style.width = timelineWidth + "px";

  // Build year blocks (accurate across multiple years)
  let yCursor = new Date(min.getFullYear(), 0, 1);
  yCursor = startOfDay(yCursor);

  while (yCursor <= max) {
    const yStart = yCursor < min ? min : yCursor;
    const nextYear = new Date(yCursor.getFullYear() + 1, 0, 1);
    const yEnd = nextYear > max ? max : addDays(nextYear, -1);

    const startOffset = Math.max(0, Math.floor((yStart - min) / 86400000));
    const endOffset = Math.min(
      totalDays,
      Math.floor((yEnd - min) / 86400000) + 1,
    );

    const w = Math.max(1, (endOffset - startOffset) * dayWidth);

    const cell = document.createElement("div");
    cell.className = "gantt-year-cell";
    cell.style.width = w + "px";
    cell.textContent = `${yStart.getFullYear()}`;
    yearsRow.appendChild(cell);

    yCursor = nextYear;
  }

  // Build ISO week labels every 7 days
  for (let d = 0; d < totalDays; d += 7) {
    const dt = addDays(min, d);
    const wk = isoWeek(dt);

    const cell = document.createElement("div");
    cell.className = "gantt-week-cell";
    cell.style.width = 7 * dayWidth + "px";
    cell.textContent = `W${wk}`;
    weeksRow.appendChild(cell);
  }

  head.appendChild(yearsRow);
  head.appendChild(weeksRow);

  // --- scrollable timeline body
  const right = document.createElement("div");
  right.className = "gantt-right";
  right.style.width = timelineWidth + "px";

  // grid lines (daily thin, weekly stronger)
  const grid = document.createElement("div");
  grid.className = "gantt-grid";
  grid.style.width = timelineWidth + "px";

  for (let d = 0; d < totalDays; d++) {
    const line = document.createElement("div");
    line.className = "gantt-grid-line" + (d % 7 === 0 ? " week" : "");
    line.style.left = d * dayWidth + "px";
    grid.appendChild(line);
  }

  right.appendChild(grid);

  // today line
  const today = startOfDay(new Date());
  if (today >= min && today <= max) {
    const offset = Math.floor((today - min) / 86400000);
    const tline = document.createElement("div");
    tline.className = "gantt-today-line";
    tline.style.left = offset * dayWidth + "px";
    right.appendChild(tline);
  }

  // bars
  tasks.forEach((t, i) => {
    const offset = Math.floor((t.start - min) / 86400000);
    const duration = Math.max(
      1,
      Math.floor((t.finish - t.start) / 86400000) + 1,
    );

    const top = i * rowHeight + 10;

    if (t.isMilestone) {
      const m = document.createElement("div");
      m.className = "gantt-milestone";
      m.style.top = top + 2 + "px";
      m.style.left = offset * dayWidth + "px";
      right.appendChild(m);
      return;
    }

    const bar = document.createElement("div");
    bar.className = "gantt-bar" + (t.isSummary ? " summary" : "");
    bar.style.top = top + "px";
    bar.style.left = offset * dayWidth + "px";
    bar.style.width = duration * dayWidth + "px";

    const prog = document.createElement("div");
    prog.className = "gantt-progress";
    prog.style.width = `${t.pct}%`;

    const label = document.createElement("div");
    label.className = "gantt-bar-label";
    label.textContent = t.name;

    bar.appendChild(prog);
    bar.appendChild(label);
    right.appendChild(bar);
  });

  // legend
  const legend = document.createElement("div");
  legend.className = "gantt-legend";
  legend.innerHTML = `
    <div class="gl-item"><span class="gl-swatch bar"></span>Task</div>
    <div class="gl-item"><span class="gl-swatch summary"></span>Summary</div>
    <div class="gl-item"><span class="gl-swatch milestone"></span>Milestone</div>
    <div class="gl-item"><span class="gl-swatch progress"></span>Progress</div>
    <div class="gl-item"><span class="gl-swatch today"></span>Today</div>
  `;

  // assemble right side
  rightOuter.appendChild(head);

  const bodyWrap = document.createElement("div");
  bodyWrap.className = "gantt-right-bodywrap";
  bodyWrap.appendChild(right);

  rightOuter.appendChild(bodyWrap);
  rightOuter.appendChild(legend);

  wrap.appendChild(left);
  wrap.appendChild(rightOuter);
  container.appendChild(wrap);

  bodyWrap.addEventListener("scroll", () => {
    leftBody.scrollTop = bodyWrap.scrollTop;
    head.scrollLeft = bodyWrap.scrollLeft;
  });
}

function inferColumnsFromData(rows, viewName) {
  const first = rows[0] || {};
  if (viewName === "schedule") {
    const keys = Object.keys(first || {});
    const find = (re) => keys.find((k) => re.test(String(k))) || null;

    const kID = find(/^id$/i) || "ID";
    const kMode = find(/task\s*mode/i) || "Task Mode";
    const kPct = find(/%\s*complete/i) || "% Complete";
    const kName = find(/task\s*name/i) || "Task Name";
    const kDur = find(/^duration/i) || "Duration";
    const kStart = find(/^start$/i) || "Start";
    const kFinish = find(/^finish$/i) || "Finish";
    const kPred = find(/predecess/i) || "Predecessors";

    return [
      {
        key: kID,
        label: "ID",
        type: "number",
        filterable: true,
        sortable: true,
      },
      {
        key: kMode,
        label: "Task Mode",
        type: "text",
        filterable: true,
        sortable: true,
      },
      {
        key: kPct,
        label: "% Complete",
        type: "number",
        filterable: true,
        sortable: true,
      },
      {
        key: kName,
        label: "Task Name",
        type: "text",
        filterable: true,
        sortable: true,
      },
      {
        key: kDur,
        label: "Duration",
        type: "text",
        filterable: true,
        sortable: true,
      },
      {
        key: kStart,
        label: "Start",
        type: "date",
        filterable: true,
        sortable: true,
      },
      {
        key: kFinish,
        label: "Finish",
        type: "date",
        filterable: true,
        sortable: true,
      },
      {
        key: kPred,
        label: "Predecessors",
        type: "text",
        filterable: true,
        sortable: true,
      },
    ];
  }

  return Object.keys(first)
    .filter((key) => {
      const lower = key.toLowerCase();
      if (viewName === "sites" && lower === "id") return false;
      if (viewName === "schedule" && isStatusKey(key)) return false;
      return true;
    })
    .map((key) => {
      const lk = key.toLowerCase();
      const type =
        lk.includes("date") ||
        lk.includes("createdat") ||
        lk.includes("updatedat")
          ? "date"
          : "text";

      return {
        key,
        label: key.replace(/_/g, " ").replace(/\b\w/g, (m) => m.toUpperCase()),
        type,
        filterable: key !== "id",
        sortable: true,
      };
    });
}

function generateTaskMode(row) {
  const name = String(row["Task Name"] ?? row["TASK NAME"] ?? "").trim();
  const start = row["Start"] ?? row["START"];
  const finish = row["Finish"] ?? row["FINISH"];

  const durRaw =
    row["Duration (days)"] ??
    row["Duration"] ??
    row["DURATION (DAYS)"] ??
    row["DURATION"];

  const durNum = toNum(durRaw);
  const dur = durNum != null ? durNum : calcDurationDays(start, finish);

  // If it looks like a heading/section row (common in schedules)
  if (name && !start && !finish && (dur === "" || dur === 0)) {
    return "Summary";
  }

  // Milestone: 0-day duration OR same start & finish
  const s = toDateOnly(start);
  const f = toDateOnly(finish);
  if (dur === 0 || (s && f && s.getTime() === f.getTime())) {
    return "Milestone";
  }

  // If no dates at all but has a name
  if (name && !start && !finish) {
    return "Task";
  }

  return "Task";
}
function toDateOnly(v) {
  if (!v) return null;
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function calcDurationDays(start, finish) {
  const s = toDateOnly(start);
  const f = toDateOnly(finish);
  if (!s || !f) return "";
  return Math.max(0, Math.round((f - s) / 86400000));
}

function getScheduleField(row, names) {
  for (const n of names) {
    if (row[n] != null && String(row[n]).trim() !== "") return row[n];
  }
  return row[names[0]]; // fallback
}

function generateTaskMode(row) {
  const name = String(
    getScheduleField(row, ["Task Name", "TASK NAME"]) ?? "",
  ).trim();
  const start = getScheduleField(row, [
    "Start",
    "START",
    "Start Date",
    "START DATE",
  ]);
  const finish = getScheduleField(row, [
    "Finish",
    "FINISH",
    "Finish Date",
    "FINISH DATE",
  ]);
  const durRaw = getScheduleField(row, [
    "Duration (days)",
    "Duration",
    "DURATION (DAYS)",
    "DURATION",
  ]);

  const durNum = Number(String(durRaw ?? "").trim());
  const dur = Number.isFinite(durNum)
    ? durNum
    : calcDurationDays(start, finish);

  // "Summary" heuristic (group/header row)
  if (name && !start && !finish && (dur === "" || dur === 0)) return "Summary";

  const s = toDateOnly(start);
  const f = toDateOnly(finish);

  // milestone: 0 duration OR same day
  if (dur === 0 || (s && f && s.getTime() === f.getTime())) return "Milestone";

  return "Task";
}

function clamp01(n) {
  n = Number(n);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(1, n));
}

async function loadAllData() {
  const statusPill = el("statusPill");
  try {
    const [employees, engagements, sites, schedule] = await Promise.all([
      fetchAllRows(SHEET_MAP.employees).catch(() => []),
      fetchAllRows(SHEET_MAP.engagements).catch(() => []),
      fetchAllRows(SHEET_MAP.sites).catch(() => []),
      fetchAllRows(SHEET_MAP.schedule).catch(() => []),
    ]);

    store.allEmployees = employees;
    store.allEngagements = engagements;
    store.allSites = sites;

    store.allSchedule = schedule;
    views.employees.columns = inferColumnsFromData(employees, "employees");
    views.engagements.columns = inferColumnsFromData(
      engagements,
      "engagements",
    );
    views.sites.columns = inferColumnsFromData(sites, "sites");
    views.schedule.columns = inferColumnsFromData(schedule, "schedule");
    const hasData =
      employees.length || engagements.length || sites.length || schedule.length;

    if (statusPill) {
      statusPill.classList.remove("error");
      statusPill.classList.add("ok");
      statusPill.innerHTML = `<span class="status-dot live"></span>${
        hasData ? "Data connected" : "Connected (no rows yet)"
      }`;
    }
  } catch (e) {
    console.error("Failed to load data:", e);
    if (statusPill) {
      statusPill.classList.remove("ok");
      statusPill.classList.add("error");
      statusPill.innerHTML = `<span class="status-dot"></span>Disconnected`;
    }
    toast(
      "error",
      "Disconnected",
      e.message || "Failed to load data from sheets",
    );
  }
}
const DEFAULT_VIEW = "overview";

function isValidView(v) {
  return Object.prototype.hasOwnProperty.call(views, v);
}

function getViewFromHash() {
  const raw = (location.hash || "").replace("#", "").trim().toLowerCase();
  return isValidView(raw) ? raw : null;
}

function setHashView(view) {
  const v = isValidView(view) ? view : DEFAULT_VIEW;
  // avoid adding duplicate history entries when already same
  if (location.hash.replace("#", "") !== v) location.hash = `#${v}`;
}

function setActiveNav(view) {
  document.querySelectorAll(".nav-link").forEach((b) => {
    b.classList.toggle("active", b.dataset.view === view);
  });
}

async function navigate(view, { pushHash = true } = {}) {
  const v = isValidView(view) ? view : DEFAULT_VIEW;
  store.view = v;
  setActiveNav(v);
  if (pushHash) setHashView(v);

  applyViewLayout();
  await loadCurrentViewData();
  if (store.view === "gantt") await renderGanttPage();
}

async function loadCurrentViewData() {
  const sheet = SHEET_MAP[store.view];
  if (!sheet) {
    store.data = [];
    render();
    return;
  }

  const titleBase = views[store.view].title;

  try {
    let rows = await fetchAllRows(sheet);
    if (store.view === "schedule" || store.view === "gantt") {
      rows = rows.map(computeScheduleFields);
    }
    store.data = rows;

    if (store.view === "employees") store.allEmployees = rows;
    if (store.view === "engagements") store.allEngagements = rows;
    if (store.view === "sites") store.allSites = rows;
    if (store.view === "schedule" || store.view === "gantt")
      store.allSchedule = rows;
    views[store.view].columns = inferColumnsFromData(rows, store.view);

    store.page = 1;
    render();
  } catch (err) {
    toast("error", "Failed to load current view", err.message);
  }
}

// ================================================================
// FILTERING & SORTING
// ================================================================

function deriveRows() {
  const cols = views[store.view].columns;
  const query = store.globalSearch.trim().toLowerCase();

  let rows = store.data.slice();

  if (query && cols.length) {
    rows = rows.filter((row) =>
      cols.some((col) =>
        String(row[col.key] ?? "")
          .toLowerCase()
          .includes(query),
      ),
    );
  }

  for (const [key, raw] of Object.entries(store.columnFilters)) {
    const col = cols.find((c) => c.key === key);
    if (!col || raw === "" || raw == null) continue;

    const rawStr = String(raw).trim();
    const val = rawStr.toLowerCase();

    rows = rows.filter((row) => {
      const cell = String(row[key] ?? "");
      const cellLower = cell.toLowerCase();

      if (col.type === "number") {
        const num = parseFloat(cell);
        const r = rawStr.trim();

        if (/^>=/.test(r)) return num >= parseFloat(r.slice(2));
        if (/^<=/.test(r)) return num <= parseFloat(r.slice(2));
        if (/^>/.test(r)) return num > parseFloat(r.slice(1));
        if (/^</.test(r)) return num < parseFloat(r.slice(1));
        if (/^-?\d+(\.\d+)?\s*-\s*-?\d+(\.\d+)?$/.test(r)) {
          const [lo, hi] = r.split("-").map((x) => parseFloat(x.trim()));
          return num >= lo && num <= hi;
        }
        return cellLower.includes(val);
      }

      if (col.type === "date") return cellLower.startsWith(val);

      return cellLower.includes(val);
    });
  }

  if (store.sortCol) {
    const col = cols.find((c) => c.key === store.sortCol);
    rows.sort((a, b) => {
      let av = a[store.sortCol] ?? "";
      let bv = b[store.sortCol] ?? "";

      if (col?.type === "number") {
        av = parseFloat(av);
        bv = parseFloat(bv);
        av = Number.isFinite(av) ? av : -Infinity;
        bv = Number.isFinite(bv) ? bv : -Infinity;
      } else if (col?.type === "date") {
        av = Date.parse(av) || 0;
        bv = Date.parse(bv) || 0;
      } else {
        av = String(av).toLowerCase();
        bv = String(bv).toLowerCase();
      }

      if (av < bv) return store.sortDir === "asc" ? -1 : 1;
      if (av > bv) return store.sortDir === "asc" ? 1 : -1;
      return 0;
    });
  }

  return rows;
}

function pageRows(rows) {
  const start = (store.page - 1) * store.pageSize;
  return rows.slice(start, start + store.pageSize);
}

function highlight(text, term) {
  if (!term) return esc(text);
  const safeText = esc(text);
  const safeTerm = esc(term).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return safeText.replace(
    new RegExp(`(${safeTerm})`, "gi"),
    '<mark class="hl">$1</mark>',
  );
}

// ================================================================
// DATE + DURATION HELPERS (short date, auto duration)
// ================================================================

function pad2(n) {
  return String(n).padStart(2, "0");
}

function startOfDay(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function toISODate(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function parseAnyDate(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime()))
    return startOfDay(value);

  const s = String(value).trim();
  if (!s) return null;

  // Excel serial number (very common when dates come from ExcelJS)
  const num = Number(s);
  if (!Number.isNaN(num) && num > 20000 && num < 60000) {
    const base = new Date(Date.UTC(1899, 11, 30));
    return startOfDay(new Date(base.getTime() + num * 86400000));
  }

  // ISO yyyy-mm-dd
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(d.getTime()) ? null : startOfDay(d);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return startOfDay(d);

  return null;
}

function formatDateShort(value) {
  const d = parseAnyDate(value);
  return d ? toISODate(d) : String(value ?? "").slice(0, 10);
}

function daysInclusive(a, b) {
  const ms = startOfDay(b).getTime() - startOfDay(a).getTime();
  return Math.floor(ms / 86400000) + 1;
}

function computeScheduleFields(row) {
  const obj = { ...(row || {}) };
  const keys = Object.keys(obj);

  const startKey = keys.find((k) => /^start$/i.test(k)) || "Start";
  const finishKey = keys.find((k) => /^finish$/i.test(k)) || "Finish";

  const durationKey =
    keys.find((k) => /^duration\s*\(days\)$/i.test(k)) ||
    keys.find((k) => /^duration/i.test(k)) ||
    "Duration (days)";

  const taskModeKey = keys.find((k) => /^task\s*mode$/i.test(k)) || "Task Mode";

  const startVal = obj[startKey];
  const finishVal = obj[finishKey];

  // Auto duration if blank
  const rawDur = obj[durationKey];
  if (String(rawDur ?? "").trim() === "") {
    obj[durationKey] = calcDurationDays(startVal, finishVal);
  }

  // Generate Task Mode if blank
  const rawTM = obj[taskModeKey];
  if (String(rawTM ?? "").trim() === "") {
    obj[taskModeKey] = generateTaskMode(obj);
  }

  return obj;
}

function toDateOnly(v) {
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function calcDurationDays(start, finish) {
  const s = toDateOnly(start);
  const f = toDateOnly(finish);
  if (!s || !f) return "";
  return Math.max(0, Math.round((f - s) / 86400000));
}

function toNum(v) {
  const n = Number(String(v ?? "").trim());
  return Number.isFinite(n) ? n : null;
}

function calcDurationDays(start, finish) {
  const s = new Date(start);
  const f = new Date(finish);
  if (isNaN(s.getTime()) || isNaN(f.getTime())) return "";

  const s0 = new Date(s.getFullYear(), s.getMonth(), s.getDate());
  const f0 = new Date(f.getFullYear(), f.getMonth(), f.getDate());
  return Math.max(0, Math.round((f0 - s0) / 86400000));
}
function formatDateMDY(v) {
  if (v == null || v === "") return "";

  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Intl.DateTimeFormat("en-US", {
      month: "short",
      day: "2-digit",
      year: "numeric",
    }).format(v);
  }

  const s = String(v).trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [yy, mm, dd] = s.split("-").map(Number);
    const d = new Date(yy, mm - 1, dd);
    return new Intl.DateTimeFormat("en-US", {
      month: "short",
      day: "2-digit",
      year: "numeric",
    }).format(d);
  }
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return new Intl.DateTimeFormat("en-US", {
        month: "short",
        day: "2-digit",
        year: "numeric",
      }).format(d);
    }
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return new Intl.DateTimeFormat("en-US", {
      month: "short",
      day: "2-digit",
      year: "numeric",
    }).format(d);
  }
  return s;
}
// ================================================================
// RENDER MAIN
// ================================================================
function formatDatePDF(v) {
  const d = parseAnyDate(v);
  if (!d) return "";
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = String(d.getFullYear()).slice(-2);
  return `${days[d.getDay()]} ${dd}-${mm}-${yy}`;
}

function formatSchedulePercent(v) {
  if (v == null || v === "") return "";
  const n = Number(String(v).replace(/[^\d.]/g, ""));
  if (isNaN(n)) return String(v);
  const pct = n <= 1 ? Math.round(n * 100) : Math.round(n);
  return `${pct}%`;
}

function formatScheduleDuration(v, start, finish) {
  if (v != null && String(v).trim() !== "") {
    const s = String(v).trim();
    if (/day/i.test(s)) return s;
    const n = Number(s.replace(/[^\d.]/g, ""));
    if (!isNaN(n)) return `${n} days`;
    return s;
  }
  const d = calcDurationDays(start, finish);
  if (d == null) return "";
  return `${d} days`;
}

function scheduleNameCellHTML(row) {
  const nameKey =
    Object.keys(row || {}).find((k) => /task\s*name/i.test(k)) || "Task Name";
  const modeKey =
    Object.keys(row || {}).find((k) => /task\s*mode/i.test(k)) || "Task Mode";
  const name = String(row[nameKey] ?? "").trim();
  const mode = String(row[modeKey] ?? "").trim();

  const isSummary =
    /summary/i.test(mode) ||
    /project\s*summary/i.test(mode) ||
    /^upper\s*myagdi/i.test(name);
  const leading = (String(row[nameKey] ?? "").match(/^(\s+)/) || [])[1] || "";
  const indent = Math.min(3, Math.floor(leading.length / 2));

  return `<span class="sch-name ${isSummary ? "is-summary" : ""}" style="--indent:${indent}">${esc(name)}</span>`;
}

// ================================================================
// SCHEDULE SHEET (PDF-like table + gantt in ONE scroller)
// ================================================================

function startOfWeekMonday(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const day = x.getDay(); // 0 Sun .. 6 Sat
  const diff = day === 0 ? -6 : 1 - day;
  x.setDate(x.getDate() + diff);
  x.setHours(0, 0, 0, 0);
  return x;
}

function addDays(d, n) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}

function daysBetween(a, b) {
  const a0 = new Date(a.getFullYear(), a.getMonth(), a.getDate());
  const b0 = new Date(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((b0 - a0) / 86400000);
}

function safeIdFromRow(row) {
  return (
    row?.id ??
    row?.ID ??
    row?.Id ??
    row?.["ID "] ??
    row?.[findKey(row, /^id$/i) || ""] ??
    ""
  );
}

function scheduleKeysFromRow(sample) {
  const k = (re, fallback) => findKey(sample, re) || fallback;

  return {
    id: k(/^id$/i, "ID"),
    taskMode: k(/^task\s*mode$/i, "Task Mode"),
    pct: k(/^%?\s*complete$/i, "% Complete"),
    taskName: k(/task\s*name/i, "Task Name"),
    duration: k(/^duration/i, "Duration"),
    start: k(/^start$/i, "Start"),
    finish: k(/^finish$/i, "Finish"),
    pred: k(/^predecessors?/i, "Predecessors"),
  };
}

function parseAnyDateSched(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v).trim();
  if (!s) return null;

  // Already ISO
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [yy, mm, dd] = s.split("-").map(Number);
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  // "Thu 17-07-25"
  const m = s.match(/^[A-Za-z]{3}\s+(\d{2})-(\d{2})-(\d{2})$/);
  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]);
    const yy = 2000 + Number(m[3]);
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function renderScheduleSheet() {
  const mount = el("scheduleSheet");
  if (!mount) return;

  const isFullGantt = store.view === "gantt";

  // Pull schedule rows even on the dedicated Gantt page
  const __prevView = store.view;
  const __prevData = store.data;
  store.view = "schedule";
  store.data = (store.allSchedule || []).slice();
  const rows = deriveRows(); // filtered + sorted using Schedule rules
  store.view = __prevView;
  store.data = __prevData;
  el("rowCount").textContent = rows.length.toLocaleString();

  if (!rows.length) {
    mount.innerHTML = `<div class="gantt-empty">No schedule records</div>`;
    return;
  }

  const sample = rows[0] || {};
  const keys = scheduleKeysFromRow(sample);

  // Timeline range (PDF-like: start a bit earlier, end a bit later)
  let minStart = null;
  let maxFinish = null;

  for (const r of rows) {
    const s = parseAnyDateSched(r[keys.start]);
    const f = parseAnyDateSched(r[keys.finish]);
    if (s && (!minStart || s < minStart)) minStart = s;
    if (f && (!maxFinish || f > maxFinish)) maxFinish = f;
  }

  if (!minStart || !maxFinish) {
    mount.innerHTML = `<div class="gantt-empty">No schedule dates found</div>`;
    return;
  }

  const timelineStart = startOfWeekMonday(addDays(minStart, -56)); // -8 weeks
  const timelineEnd = startOfWeekMonday(addDays(maxFinish, 28)); // +4 weeks
  const weeks = [];
  for (let d = new Date(timelineStart); d <= timelineEnd; d = addDays(d, 7)) {
    weeks.push(new Date(d));
  }

  const cellW = store.ganttCellW || 22;
  const ganttWidth = weeks.length * cellW;

  // Build year spans
  const yearSpans = [];
  let curYear = weeks[0].getFullYear();
  let spanStart = 0;
  for (let i = 0; i < weeks.length; i++) {
    const y = weeks[i].getFullYear();
    if (y !== curYear) {
      yearSpans.push({ year: curYear, from: spanStart, to: i });
      curYear = y;
      spanStart = i;
    }
  }
  yearSpans.push({ year: curYear, from: spanStart, to: weeks.length });

  const weekLabels = weeks.map((_, i) => `W${i - 8}`);

  // Sticky left column widths (match PDF vibe)
  const leftCols = isFullGantt
    ? [
        { key: keys.id, label: "ID", w: 64, cls: "mono" },
        { key: keys.taskName, label: "TASK NAME", w: 460, cls: "sch-task" },
        { key: keys.start, label: "START", w: 130, cls: "mono" },
        { key: keys.finish, label: "FINISH", w: 130, cls: "mono" },
      ]
    : [
        { key: keys.id, label: "ID", w: 60, cls: "mono" },
        { key: keys.taskMode, label: "TASK MODE", w: 120, cls: "" },
        { key: keys.pct, label: "% COMPLETE", w: 110, cls: "mono sch-pct" },
        { key: keys.taskName, label: "TASK NAME", w: 360, cls: "sch-task" },
        { key: keys.duration, label: "DURATION", w: 120, cls: "mono" },
        { key: keys.start, label: "START", w: 120, cls: "mono" },
        { key: keys.finish, label: "FINISH", w: 120, cls: "mono" },
        { key: keys.pred, label: "PREDECESSORS", w: 150, cls: "mono" },
        { key: "__actions", label: "ACTIONS", w: 140, cls: "mono" },
      ];

  // Compute sticky left offsets
  let acc = 0;
  for (const c of leftCols) {
    c.left = acc;
    acc += c.w;
  }
  const leftTotal = acc;

  const headLeft = leftCols
    .map(
      (c) =>
        `<div class="sch-h sch-sticky" style="width:${c.w}px; left:${c.left}px">${esc(
          c.label,
        )}</div>`,
    )
    .join("");

  const headYears = yearSpans
    .map((ys) => {
      const w = (ys.to - ys.from) * cellW;
      return `<div class="sch-year" style="width:${w}px">${ys.year}</div>`;
    })
    .join("");

  const headWeeks = weekLabels
    .map((lab) => `<div class="sch-week" style="width:${cellW}px">${lab}</div>`)
    .join("");

  const today = startOfWeekMonday(new Date());
  const todayIdx = Math.max(
    0,
    Math.min(
      weeks.length - 1,
      Math.round(daysBetween(timelineStart, today) / 7),
    ),
  );

  const bodyRows = rows
    .map((r) => {
      const idVal = safeIdFromRow(r);
      const mode = String(r[keys.taskMode] ?? "").trim();
      const isMilestone = /milestone/i.test(mode);
      const isSummary =
        /summary/i.test(mode) || /project\s*summary/i.test(mode);

      const s = parseAnyDateSched(r[keys.start]);
      const f = parseAnyDateSched(r[keys.finish]);

      const startIdx = s
        ? Math.round(daysBetween(timelineStart, startOfWeekMonday(s)) / 7)
        : null;
      const endIdx = f
        ? Math.round(daysBetween(timelineStart, startOfWeekMonday(f)) / 7)
        : startIdx;

      const leftPx = startIdx == null ? 0 : startIdx * cellW;
      const spanWeeks =
        startIdx == null || endIdx == null
          ? 0
          : Math.max(1, endIdx - startIdx + 1);
      const widthPx = spanWeeks * cellW;

      const pct = Number(String(r[keys.pct] ?? "").replace(/[^\d.]/g, ""));
      const pct01 = isNaN(pct) ? 0 : pct > 1 ? pct / 100 : pct;

      const nameText = esc(String(r[keys.taskName] ?? "").trim());
      const dateTag = s
        ? `${String(String(r[keys.start] ?? "")).slice(0, 5)}`
        : "";

      const barHTML = isMilestone
        ? `<div class="sch-item milestone" style="left:${leftPx}px">
             <span class="sch-dia"></span>
             <span class="sch-milabel">${nameText}</span>
           </div>`
        : `<div class="sch-item ${isSummary ? "summary" : "task"}" style="left:${leftPx}px; width:${widthPx}px">
            <span class="sch-bar">
              <span class="sch-progress" style="width:${Math.round(pct01 * 100)}%"></span>
            </span>
            <span class="sch-label">${nameText}</span>
          </div>`;

      const ganttCell = `
        <div class="sch-gantt-cell" style="width:${ganttWidth}px">
          <div class="sch-gantt-grid" style="--cw:${cellW}px">
            <div class="sch-today" style="left:${todayIdx * cellW}px"></div>
            ${barHTML}
          </div>
        </div>`;

      const leftCells = leftCols
        .map((c) => {
          if (c.key === "__actions") {
            return `<div class="sch-c sch-sticky" style="width:${c.w}px; left:${c.left}px">
              <button class="btn ghost btn-xs" data-act="edit" data-id="${esc(idVal)}">Edit</button>
              <button class="btn ghost btn-xs" data-act="del" data-id="${esc(idVal)}">Delete</button>
            </div>`;
          }

          if (
            String(c.key).toLowerCase() === String(keys.taskName).toLowerCase()
          ) {
            return `<div class="sch-c sch-sticky ${c.cls}" style="width:${c.w}px; left:${c.left}px">${scheduleNameCellHTML(
              r,
            )}</div>`;
          }

          if (
            String(c.key).toLowerCase() === String(keys.start).toLowerCase()
          ) {
            return `<div class="sch-c sch-sticky ${c.cls}" style="width:${c.w}px; left:${c.left}px">${esc(
              formatDatePDF(r[keys.start]),
            )}</div>`;
          }

          if (
            String(c.key).toLowerCase() === String(keys.finish).toLowerCase()
          ) {
            return `<div class="sch-c sch-sticky ${c.cls}" style="width:${c.w}px; left:${c.left}px">${esc(
              formatDatePDF(r[keys.finish]),
            )}</div>`;
          }

          if (
            String(c.key).toLowerCase() === String(keys.duration).toLowerCase()
          ) {
            return `<div class="sch-c sch-sticky ${c.cls}" style="width:${c.w}px; left:${c.left}px">${esc(
              formatScheduleDuration(
                r[keys.duration],
                r[keys.start],
                r[keys.finish],
              ),
            )}</div>`;
          }

          if (String(c.key).toLowerCase() === String(keys.pct).toLowerCase()) {
            return `<div class="sch-c sch-sticky ${c.cls}" style="width:${c.w}px; left:${c.left}px">${esc(
              formatSchedulePercent(r[keys.pct]),
            )}</div>`;
          }

          const raw = r[c.key] ?? "";
          return `<div class="sch-c sch-sticky ${c.cls}" style="width:${c.w}px; left:${c.left}px">${esc(
            raw,
          )}</div>`;
        })
        .join("");

      const rowType = isMilestone
        ? "milestone"
        : isSummary
          ? "summary"
          : "task";
      return `<div class="sch-row sch-row--${rowType}" style="--leftw:${leftTotal}px">${leftCells}${ganttCell}</div>`;
    })
    .join("");

  // Apply/remove gantt-fullscreen class on the schedule layout
  const schLayout = el("scheduleLayout");
  if (schLayout) {
    if (isFullGantt) schLayout.classList.add("gantt-fullscreen");
    else schLayout.classList.remove("gantt-fullscreen");
  }

  const ganttProTopbar = isFullGantt
    ? `
    <div class="gantt-pro-topbar">
      <button class="gantt-pro-backbtn" id="ganttProBackBtn" type="button">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M9 1L3 7L9 13" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
        Back to Schedule
      </button>
      <div class="gantt-pro-info">
        <div class="gantt-pro-title">Schedule Gantt</div>
        <div class="gantt-pro-subtitle">Upper Myagdi-1 Hydro Power Project &middot; Troyer ITA &amp; Troyer Hydro</div>
      </div>
      <div class="gantt-pro-right">
        <div class="gantt-pro-legend">
          <span class="gpl-item"><span class="gpl-swatch task"></span>Task</span>
          <span class="gpl-item"><span class="gpl-swatch summary"></span>Summary</span>
          <span class="gpl-item"><span class="gpl-swatch milestone"></span>Milestone</span>
          <span class="gpl-item"><span class="gpl-swatch today"></span>Today</span>
        </div>
        <div class="gantt-zoom-ctrl">
          <label class="gantt-zoom-lbl" for="ganttProZoom">Zoom</label>
          <select class="gantt-zoom-sel" id="ganttProZoom">
            <option value="14">Compact</option>
            <option value="18">Normal</option>
            <option value="24">Wide</option>
            <option value="32">Extra Wide</option>
          </select>
        </div>
      </div>
    </div>
  `
    : "";

  mount.innerHTML =
    ganttProTopbar +
    `
    <div class="sch-sheet ${isFullGantt ? "sch-sheet--gantt" : ""}" style="--cw:${cellW}px">
      <div class="sch-head">
        <div class="sch-head-left" style="width:${leftTotal}px">${headLeft}</div>
        <div class="sch-head-right" style="width:${ganttWidth}px">
          <div class="sch-year-row">${headYears}</div>
          <div class="sch-week-row">${headWeeks}</div>
        </div>
      </div>
      <div class="sch-body">${bodyRows}</div>
    </div>`;

  // Wire up the inline gantt controls
  if (isFullGantt) {
    const backBtn = el("ganttProBackBtn");
    if (backBtn)
      backBtn.onclick = () => navigate("schedule", { pushHash: true });
    const zoomSel = el("ganttProZoom");
    if (zoomSel) {
      zoomSel.value = String(cellW);
      zoomSel.onchange = () => {
        store.ganttCellW = Number(zoomSel.value) || 18;
        renderScheduleSheet();
      };
    }
  }
}

function render() {
  if (store.view === "overview") {
    renderOverviewDashboard();
    return;
  }

  if (store.view === "gantt") {
    renderGanttPage();
    return;
  }

  renderTableView();
}

// ================================================================
// OVERVIEW DASHBOARD
// ================================================================

function renderOverviewDashboard() {
  el("overviewView").style.display = "";
  el("tableView").style.display = "none";

  const chartsGrid = document.querySelector(".charts-grid");
  if (chartsGrid) chartsGrid.style.display = "";

  const sitesCanvas = el("sitesChart");
  if (sitesCanvas) {
    const card =
      sitesCanvas.closest(".chart-card") || sitesCanvas.closest(".card");
    if (card) card.style.display = "";
  }

  const engCanvas = el("engagementsChart");
  if (engCanvas) {
    const card = engCanvas.closest(".chart-card") || engCanvas.closest(".card");
    if (card) card.style.display = "";
  }

  const statusCanvas = el("statusChart");
  if (statusCanvas) {
    const card =
      statusCanvas.closest(".chart-card") || statusCanvas.closest(".card");
    if (card) card.style.display = "";
  }

  if (chartInstances.status) {
    chartInstances.status.destroy();
    chartInstances.status = null;
  }
  if (store.view === "schedule" && /^task\s*mode$/i.test(key)) {
    const v = String(row[key] ?? "").trim() || "Task";
    td.innerHTML = `<span class="pill">${v}</span>`;
    return;
  }

  renderOverviewStats();
  renderPeoplePerProjectChart();
  renderResponsibilityChart();

  const items = document.querySelectorAll(
    ".stat-card, .chart-card, .activity-item",
  );

  items.forEach((item, i) => {
    item.classList.remove("show");
    item.classList.add("anim-in");
    item.style.setProperty("--d", `${i * 80}ms`);

    requestAnimationFrame(() => {
      item.classList.add("show");
    });
  });
}

function renderPeoplePerProjectChart() {
  const canvas = el("sitesChart");
  if (!canvas) return;

  if (chartInstances.sites) chartInstances.sites.destroy();

  const rows = store.allEngagements || [];
  if (!rows.length) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No engagement data available</div>';
    return;
  }

  const sample = rows[0] || {};
  const siteKey =
    Object.keys(sample).find((k) => /site\s*engaged/i.test(k)) ||
    Object.keys(sample).find((k) => /site|location/i.test(k)) ||
    null;

  if (!siteKey) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No "Site Engaged" field found</div>';
    return;
  }

  const map = new Map();
  for (const r of rows) {
    const site = String(r[siteKey] || "").trim();
    if (!site) continue;
    map.set(site, (map.get(site) || 0) + 1);
  }

  const items = Array.from(map.entries())
    .map(([site, count]) => ({ site, count }))
    .sort((a, b) => b.count - a.count);

  const labels = items.map((x) => x.site);
  const data = items.map((x) => x.count);

  chartInstances.sites = new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "People",
          data,
          backgroundColor: "#4f8ef7",
          borderRadius: 8,
          barPercentage: 0.7,
          categoryPercentage: 0.8,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: true, text: "People per Project (Site Engaged)" },
      },
      scales: {
        x: {
          title: { display: true, text: "Project / Site" },
          ticks: { maxRotation: 25, minRotation: 0 },
        },
        y: {
          beginAtZero: true,
          ticks: { precision: 0 },
          title: { display: true, text: "Number of People" },
        },
      },
    },
  });
}

function renderResponsibilityChart() {
  const canvas = el("engagementsChart");
  if (!canvas) return;

  if (chartInstances.engagements) chartInstances.engagements.destroy();

  const rows = store.allSchedule || [];
  if (!rows.length) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No schedule data available</div>';
    return;
  }

  const sample = rows[0] || {};
  const respKey =
    Object.keys(sample).find((k) => /responsibility|owner|assigned/i.test(k)) ||
    null;

  if (!respKey) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No "Responsibility" column found</div>';
    return;
  }

  let troyerCount = 0;
  let customerCount = 0;

  rows.forEach((r) => {
    const resp = String(r[respKey] || "")
      .trim()
      .toLowerCase();

    if (!resp) return; // ignore empty

    if (resp === "troyer") troyerCount++;
    else if (resp === "customer") customerCount++;
  });

  chartInstances.engagements = new Chart(canvas, {
    type: "bar",
    data: {
      labels: ["TROYER", "CUSTOMER"],
      datasets: [
        {
          label: "Tasks",
          data: [troyerCount, customerCount],
          backgroundColor: ["#2d7a4f", "#4f8ef7"],
          borderRadius: 8,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: {
          display: true,
          text: "Tasks by Responsibility",
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { precision: 0 },
          title: { display: true, text: "Number of Tasks" },
        },
      },
    },
  });
}

function renderEngagementReportChart() {
  const canvas = el("statusChart");
  if (!canvas) return;

  if (chartInstances.status) chartInstances.status.destroy();

  const rows = store.allEngagements || [];
  if (!rows.length) return;

  const sample = rows[0] || {};
  const statusKey = Object.keys(sample).find((k) => /status|phase/i.test(k));

  if (!statusKey) return;

  let active = 0;
  let completed = 0;

  for (const r of rows) {
    const s = String(r[statusKey] || "")
      .trim()
      .toLowerCase();
    if (s === "active") active++;
    else if (s === "completed") completed++;
  }

  chartInstances.status = new Chart(canvas, {
    type: "bar",
    data: {
      labels: ["Active", "Completed"],
      datasets: [
        {
          data: [active, completed],
          backgroundColor: ["#28a745", "#6c757d"],
          borderRadius: 8,
          barPercentage: 0.5,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: {
          display: true,
          text: "Engagement Status Overview",
        },
      },
      scales: {
        x: {
          title: {
            display: true,
            text: "Engagement Status",
          },
        },
        y: {
          beginAtZero: true,
          ticks: { precision: 0 },
          title: {
            display: true,
            text: "Number of People",
          },
        },
      },
    },
  });
}

function renderOverviewStats() {
  const container = el("overviewStats");

  const empCount = store.allEmployees.length;
  const engCount = store.allEngagements.length;

  container.innerHTML = `
    <div class="stat-card">
      <div class="stat-label">Total Employees</div>
      <div class="stat-value" data-count="${empCount}">0</div>
      <div class="stat-sub">Active workforce</div>
    </div>

    <div class="stat-card">
      <div class="stat-label">Engagements</div>
      <div class="stat-value" data-count="${engCount}">0</div>
      <div class="stat-sub">Total projects</div>
    </div>
  `;
  document.querySelectorAll(".stat-value").forEach((el) => {
    const target = parseInt(el.dataset.count || "0", 10);
    animateCount(el, target, 800);
  });
}
function animateCount(el, to, duration = 800) {
  const start = performance.now();
  const from = 0;

  function update(time) {
    const progress = Math.min((time - start) / duration, 1);
    const value = Math.floor(progress * to);
    el.textContent = value.toLocaleString();

    if (progress < 1) {
      requestAnimationFrame(update);
    } else {
      el.textContent = to.toLocaleString();
    }
  }

  requestAnimationFrame(update);
}

function renderCharts() {
  renderSitesChart();
  renderEngagementsChart();
  renderStatusChart();
}

function renderSitesChart() {
  const canvas = el("sitesChart");
  if (!canvas) return;

  if (chartInstances.sites) chartInstances.sites.destroy();

  if (!store.allEngagements.length) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No engagement data available</div>';
    return;
  }

  // Detect site field
  const sample = store.allEngagements[0] || {};
  const siteField = Object.keys(sample).find((k) => /site|location/i.test(k));

  if (!siteField) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No site field found</div>';
    return;
  }

  const counts = {};

  store.allEngagements.forEach((eng) => {
    const site = String(eng[siteField] || "").trim();
    if (!site) return;
    counts[site] = (counts[site] || 0) + 1;
  });

  const labels = Object.keys(counts);
  const data = Object.values(counts);

  chartInstances.sites = new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Employees",
          data,
          backgroundColor: "#2d7a4f",
          borderRadius: 6,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      scales: {
        y: {
          beginAtZero: true,
          ticks: { precision: 0 },
        },
      },
      plugins: {
        legend: { display: false },
      },
    },
  });
}

function renderEngagementsChart() {
  const canvas = el("engagementsChart");
  if (!canvas) return;

  if (chartInstances.engagements) chartInstances.engagements.destroy();

  if (!store.allEngagements.length) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No engagement data available</div>';
    return;
  }

  const dateField = Object.keys(store.allEngagements[0]).find(
    (k) =>
      k.toLowerCase().includes("start") ||
      k.toLowerCase().includes("date") ||
      k.toLowerCase().includes("created"),
  );

  if (!dateField) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No date field found</div>';
    return;
  }

  const monthlyCounts = {};
  const now = new Date();

  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    monthlyCounts[key] = 0;
  }

  store.allEngagements.forEach((eng) => {
    const dateStr = eng[dateField];
    if (!dateStr) return;
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    if (monthlyCounts.hasOwnProperty(key)) {
      monthlyCounts[key]++;
    }
  });

  const labels = Object.keys(monthlyCounts).map((k) => {
    const [y, m] = k.split("-");
    return new Date(y, m - 1).toLocaleDateString("en-US", { month: "short" });
  });
  const data = Object.values(monthlyCounts);

  chartInstances.engagements = new Chart(canvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Engagements",
          data,
          borderColor: "#2d7a4f",
          backgroundColor: "rgba(45, 122, 79, 0.1)",
          borderWidth: 2,
          fill: true,
          tension: 0.3,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            font: { family: "DM Mono", size: 10 },
            color: "#7a7772",
            precision: 0,
          },
          grid: { color: "#e2e0da" },
        },
        x: {
          ticks: { font: { family: "DM Sans", size: 10 }, color: "#7a7772" },
          grid: { display: false },
        },
      },
      plugins: { legend: { display: false } },
    },
  });
}

function renderStatusChart() {
  const canvas = el("statusChart");
  if (!canvas) return;

  if (chartInstances.status) chartInstances.status.destroy();

  if (!store.allEngagements.length) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No engagement data available</div>';
    return;
  }

  const statusField = Object.keys(store.allEngagements[0]).find(
    (k) =>
      k.toLowerCase().includes("status") || k.toLowerCase().includes("phase"),
  );

  if (!statusField) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No status field found</div>';
    return;
  }

  const statusCounts = {};
  store.allEngagements.forEach((eng) => {
    const status = eng[statusField] || "Unknown";
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });

  const labels = Object.keys(statusCounts);
  const data = Object.values(statusCounts);

  chartInstances.status = new Chart(canvas, {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data,
          backgroundColor: [
            "#2d7a4f",
            "#b25a00",
            "#1a1916",
            "#7a7772",
            "#c0392b",
          ],
          borderWidth: 2,
          borderColor: "#f5f4f0",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            font: { family: "DM Sans", size: 11 },
            color: "#4a4844",
            padding: 12,
          },
        },
      },
    },
  });
}

function idSeq(id) {
  const m = String(id || "").match(/-(\d+)\s*$/);
  return m ? parseInt(m[1], 10) : 0;
}

function findFirstKey(obj, patterns) {
  const keys = Object.keys(obj || {});
  return keys.find((k) => patterns.some((re) => re.test(String(k)))) || null;
}

function activityTimeValue(row) {
  if (!row) return 0;

  // Prefer createdAt/updatedAt if you ever add them later
  const timeKey = findFirstKey(row, [
    /updated\s*at/i,
    /created\s*at/i,
    /\bcreated\b/i,
    /\bupdated\b/i,
  ]);
  if (timeKey) {
    const t = Date.parse(row[timeKey]);
    if (!Number.isNaN(t)) return t;
  }

  // Fallback to any meaningful date field (engagement/schedule)
  const dateKey = findFirstKey(row, [
    /starting\s*date/i,
    /start\s*date/i,
    /end\s*date/i,
    /\bdate\b/i,
  ]);
  if (dateKey) {
    const t = Date.parse(String(row[dateKey]).slice(0, 10));
    if (!Number.isNaN(t)) return t;
  }
  return idSeq(row.id);
}

function formatActivityTime(row) {
  const t = activityTimeValue(row);
  if (t > 1000000000) return new Date(t).toLocaleDateString();
  return "Recently";
}

// ================================================================
// TABLE VIEW RENDERING
// ================================================================

function renderTableView() {
  el("overviewView").style.display = "none";
  el("tableView").style.display = "";

  const isSchedule = store.view === "schedule";
  const isGantt = store.view === "gantt";

  renderStats();
  renderFilterChips();
  renderAdvancedFilterPanel();
  mountToolbarButtons();

  // toolbar toggles
  const btnGantt = document.getElementById("btnGantt");
  const btnBack = document.getElementById("btnBackFromGantt");
  const zoomWrap = document.getElementById("ganttZoomWrap");
  const btnAdd = document.getElementById("btnAddRow");
  const btnDl = document.getElementById("btnDownloadExcel");

  if (btnGantt) btnGantt.style.display = isSchedule ? "" : "none";
  if (btnBack) btnBack.style.display = isGantt ? "" : "none";
  if (zoomWrap) zoomWrap.style.display = isGantt ? "flex" : "none";
  if (btnAdd) btnAdd.style.display = isSchedule || isGantt ? "" : "";
  if (btnDl) btnDl.style.display = isGantt ? "none" : "";

  applyViewLayout();

  // Schedule + Gantt both render the schedule sheet
  if (isSchedule || isGantt) {
    renderScheduleSheet();
    return;
  }

  // other views use normal table
  renderTable();
  renderPagination();
}

function mountToolbarButtons() {
  const right = document.querySelector(".table-toolbar-right");
  if (!right || right.dataset.mounted === "1") return;
  right.dataset.mounted = "1";

  const wrap = document.createElement("div");
  wrap.style.display = "flex";
  wrap.style.gap = "8px";
  wrap.style.alignItems = "center";

  const backBtn = document.createElement("button");
  backBtn.id = "btnBackFromGantt";
  backBtn.className = "btn ghost btn-sm";
  backBtn.textContent = "Back to Schedule";
  backBtn.style.display = "none";
  backBtn.onclick = async () => {
    await navigate("schedule", { pushHash: true });
  };

  const ganttBtn = document.createElement("button");
  ganttBtn.id = "btnGantt";
  ganttBtn.className = "btn ghost btn-sm";
  ganttBtn.textContent = "Gantt";
  ganttBtn.style.display = "none";
  ganttBtn.onclick = async () => {
    // Ensure schedule data exists
    if (!store.allSchedule?.length) await loadAllData();
    await navigate("gantt", { pushHash: true });
  };

  const zoomWrap = document.createElement("div");
  zoomWrap.id = "ganttZoomWrap";
  zoomWrap.style.display = "none";
  zoomWrap.style.alignItems = "center";
  zoomWrap.style.gap = "8px";
  zoomWrap.style.marginLeft = "6px";

  const zoomLabel = document.createElement("span");
  zoomLabel.textContent = "Zoom";
  zoomLabel.style.fontWeight = "800";
  zoomLabel.style.fontSize = "12px";
  zoomLabel.style.color = "var(--text-2)";

  const zoomSel = document.createElement("select");
  zoomSel.id = "ganttZoom";
  zoomSel.className = "page-size-select";
  zoomSel.innerHTML = `
    <option value="18">Dense</option>
    <option value="22" selected>Normal</option>
    <option value="28">Large</option>
    <option value="36">Extra</option>
  `;
  zoomSel.onchange = () => {
    store.ganttCellW = Number(zoomSel.value) || 22;
    if (store.view === "gantt") renderScheduleSheet();
  };

  zoomWrap.appendChild(zoomLabel);
  zoomWrap.appendChild(zoomSel);

  const downloadBtn = document.createElement("button");
  downloadBtn.id = "btnDownloadExcel";
  downloadBtn.className = "btn ghost btn-sm";
  downloadBtn.textContent = "Download Excel";
  downloadBtn.onclick = () => (window.location.href = `${API}/excel/download`);

  const addBtn = document.createElement("button");
  addBtn.id = "btnAddRow";
  addBtn.className = "btn primary btn-sm";
  addBtn.textContent = "Add Row";
  addBtn.onclick = () => openFormModal("add", null);

  wrap.appendChild(backBtn);
  wrap.appendChild(ganttBtn);
  wrap.appendChild(zoomWrap);
  wrap.appendChild(downloadBtn);
  wrap.appendChild(addBtn);

  right.prepend(wrap);
}

function renderStats() {
  const cfg = views[store.view];
  const total = store.data.length;

  if (!total) {
    el("stats").innerHTML = `
      <div class="empty-stats">
        <div class="empty-stats-title">No data available</div>
        <div class="empty-stats-desc">
          ${cfg.title} will appear here once data is loaded.
        </div>
      </div>`;
    return;
  }

  if (store.view === "schedule") {
    const rows = store.data;

    // find status column name safely (status / phase)
    const sample = rows[0] || {};
    const statusKey =
      Object.keys(sample).find((k) => /status|phase/i.test(k)) || "status";

    const completed = rows.filter(
      (r) => String(r[statusKey] || "").toLowerCase() === "completed",
    ).length;

    const active = rows.filter(
      (r) => String(r[statusKey] || "").toLowerCase() === "active",
    ).length;

    const progress = total ? Math.round((completed / total) * 100) : 0;

    el("stats").innerHTML = `
      <div class="stat-card">
        <div class="stat-label">Total Tasks</div>
        <div class="stat-value">${total.toLocaleString()}</div>
        <div class="stat-sub">All schedule items</div>
      </div>

      <div class="stat-card">
        <div class="stat-label">Completed</div>
        <div class="stat-value">${completed.toLocaleString()}</div>
        <div class="stat-sub">Finished tasks</div>
      </div>

      <div class="stat-card">
        <div class="stat-label">Active</div>
        <div class="stat-value">${active.toLocaleString()}</div>
        <div class="stat-sub">Ongoing tasks</div>
      </div>

      <div class="stat-card progress-card">
        <div class="stat-label">Progress Report</div>
        <div class="stat-value">${progress}%</div>
        <div class="progress-bar">
          <div class="progress-fill" style="width:${progress}%"></div>
        </div>
        <div class="stat-sub">${completed}/${total} completed</div>
      </div>
    `;
    return;
  }

  // Default for other views
  el("stats").innerHTML = `
    <div class="stat-card">
      <div class="stat-label">Total Records</div>
      <div class="stat-value">${total.toLocaleString()}</div>
      <div class="stat-sub">All ${cfg.title.toLowerCase()}</div>
    </div>`;
}

function renderFilterChips() {
  const chips = el("filterChips");
  const clearAll = el("btnClearAll");
  const badge = el("filterCountBadge");
  const clearSrch = el("clearSearch");

  const activeColFilters = Object.entries(store.columnFilters).filter(
    ([, v]) => v !== "" && v != null,
  );

  const hasSearch = store.globalSearch.trim() !== "";
  const totalActive = activeColFilters.length + (hasSearch ? 1 : 0);

  clearSrch.style.display = hasSearch ? "" : "none";
  clearAll.style.display = totalActive ? "" : "none";
  badge.style.display = totalActive ? "" : "none";
  badge.textContent = totalActive || "";

  let html = "";

  if (hasSearch) {
    html += `
      <span class="filter-chip" data-chip-type="search">
        <span class="filter-chip-label">Search</span>
        <span class="filter-chip-value">${esc(store.globalSearch)}</span>
        <button class="filter-chip-remove" data-chip-type="search" title="Remove">×</button>
      </span>`;
  }

  const cfg = views[store.view];
  for (const [key, val] of activeColFilters) {
    const col = cfg.columns.find((c) => c.key === key);
    const label = col ? col.label : key;
    html += `
      <span class="filter-chip" data-chip-col="${esc(key)}">
        <span class="filter-chip-label">${esc(label)}</span>
        <span class="filter-chip-value">${esc(val)}</span>
        <button class="filter-chip-remove" data-chip-col="${esc(key)}" title="Remove">×</button>
      </span>`;
  }

  chips.innerHTML = html;
}

function renderAdvancedFilterPanel() {
  const inner = el("advancedFiltersInner");
  const cfg = views[store.view];
  const fcols = cfg.columns.filter((c) => c.filterable !== false);

  if (!fcols.length) {
    inner.innerHTML = `
      <div class="adv-filter-cell" style="grid-column:1/-1">
        <span style="font-size:12px;color:var(--text-3)">
          No filterable columns available
        </span>
      </div>`;
    return;
  }

  inner.innerHTML = fcols
    .map((col) => {
      const cur = store.columnFilters[col.key] ?? "";

      if (col.type === "date") {
        return `
        <div class="adv-filter-cell">
          <label class="adv-filter-label" for="adv_${esc(col.key)}">${esc(col.label)}</label>
          <input class="adv-filter-input" type="date" id="adv_${esc(col.key)}"
                 value="${esc(cur)}" data-col="${esc(col.key)}" />
        </div>`;
      }

      if (col.type === "number") {
        return `
        <div class="adv-filter-cell">
          <label class="adv-filter-label">${esc(col.label)}</label>
          <input class="adv-filter-input" type="text"
                 placeholder="e.g. 10, >=5, 1-50"
                 value="${esc(cur)}" data-col="${esc(col.key)}" />
        </div>`;
      }

      return `
      <div class="adv-filter-cell">
        <label class="adv-filter-label" for="adv_${esc(col.key)}">${esc(col.label)}</label>
        <input class="adv-filter-input" type="text" id="adv_${esc(col.key)}"
               placeholder="Filter by ${esc(col.label).toLowerCase()}…"
               value="${esc(cur)}" data-col="${esc(col.key)}" />
      </div>`;
    })
    .join("");
}

function renderTable() {
  const cfg = views[store.view];
  const cols = cfg.columns || [];
  const term = store.globalSearch.trim().toLowerCase();
  const sheet = SHEET_MAP[store.view];
  const showActions = store.view !== "overview" && !!sheet;

  const isSchedule = store.view === "schedule";
  const thead =
    document.getElementById("theadSchedule") ||
    document.getElementById("theadDefault");

  const tbody =
    document.getElementById("tbodySchedule") ||
    document.getElementById("tbodyDefault");

  if (!thead || !tbody) {
    toast(
      "error",
      "Failed to load current view",
      "Table mount missing (thead/tbody)",
    );
    return;
  }

  if (!cols.length) {
    thead.innerHTML = "";
  } else {
    thead.innerHTML = `<tr>${cols
      .map((col) => {
        const sortable = col.sortable !== false;
        const isSorted = store.sortCol === col.key;
        const cls = [
          sortable ? "sortable" : "",
          isSorted && store.sortDir === "asc" ? "sort-asc" : "",
          isSorted && store.sortDir === "desc" ? "sort-desc" : "",
        ]
          .filter(Boolean)
          .join(" ");
        return `<th class="${cls}" data-sort="${sortable ? col.key : ""}">${esc(col.label)}</th>`;
      })
      .join("")}${showActions ? `<th>Actions</th>` : ""}</tr>`;
  }

  const filtered = deriveRows();
  const paged = pageRows(filtered);

  el("rowCount").textContent = filtered.length.toLocaleString();

  if (!paged.length) {
    const hasData = (store.data || []).length > 0;
    tbody.innerHTML = `
      <tr class="empty-row">
        <td colspan="${(cols.length || 1) + (showActions ? 1 : 0)}">
          <span class="empty-row-label">
            ${hasData ? "No records match the current filters" : "No data available yet"}
          </span>
        </td>
      </tr>`;
    return;
  }

  tbody.innerHTML = paged
    .map((row) => {
      const tds = cols
        .map((col) => {
          const raw = row[col.key] ?? "";
          const kLower = String(col.key).trim().toLowerCase();

          if (isStatusKey(col.key) && store.view !== "schedule") {
            return `<td>${statusBadgeHTML(raw)}</td>`;
          }

          if (
            store.view === "schedule" &&
            (kLower === "start" || kLower === "finish")
          ) {
            return `<td class="mono">${esc(formatDatePDF(raw))}</td>`;
          }

          if (
            store.view === "schedule" &&
            (kLower === "% complete" ||
              kLower === "percent complete" ||
              kLower === "progress")
          ) {
            return `<td class="mono sch-pct">${esc(formatSchedulePercent(raw))}</td>`;
          }

          if (store.view === "schedule" && kLower === "duration") {
            const startV = row[findKey(row, /^start$/i) || "Start"];
            const finV = row[findKey(row, /^finish$/i) || "Finish"];
            return `<td class="mono">${esc(formatScheduleDuration(raw, startV, finV))}</td>`;
          }

          if (store.view === "schedule" && kLower === "task name") {
            return `<td class="sch-task">${scheduleNameCellHTML(row)}</td>`;
          }

          if (col.type === "date") {
            return `<td class="mono">${esc(formatDateShort(raw))}</td>`;
          }

          return col.type === "number"
            ? `<td class="mono">${esc(raw)}</td>`
            : `<td>${highlight(raw, term)}</td>`;
        })
        .join("");

      const rowId =
        store.view === "sites"
          ? row.Location || row.location || ""
          : store.view === "schedule"
            ? (row.ID ?? row.Id ?? row.id ?? "")
            : row.id || "";

      const actions = showActions
        ? `<td class="mono">
            <button class="btn ghost btn-sm" data-act="edit" data-id="${esc(rowId)}">Edit</button>
            <button class="btn ghost btn-sm" data-act="del" data-id="${esc(rowId)}">Delete</button>
            ${
              store.view === "employees"
                ? `<button class="btn ghost btn-sm" data-act="export" data-id="${esc(rowId)}">Export</button>`
                : ""
            }
          </td>`
        : "";

      return `<tr>${tds}${actions}</tr>`;
    })
    .join("");
}

function renderPagination() {
  const total = deriveRows().length;
  const pages = Math.max(1, Math.ceil(total / store.pageSize));
  store.page = Math.min(store.page, pages);

  el("pageNo").textContent = store.page;
  el("pageTotal").textContent = pages;
  el("btnPrevPage").disabled = store.page <= 1;
  el("btnNextPage").disabled = store.page >= pages;
}

// ================================================================
// VIEW SWITCHING
// ================================================================

async function setView(view) {
  if (!view) view = "overview";

  store.view = view;
  applyViewLayout();
  mountScheduleGanttButton();
  const isGantt = store.view === "gantt";
  const scheduleHeader = document.getElementById("scheduleHeader");
  if (scheduleHeader) scheduleHeader.style.display = "none"; // topbar shows project info in gantt

  document.querySelectorAll(".nav-link").forEach((link) => {
    link.classList.remove("active");
  });

  const activeNav = document.querySelector(`.nav-link[data-view="${view}"]`);
  if (activeNav) activeNav.classList.add("active");

  const overviewView = document.getElementById("overviewView");
  const tableView = document.getElementById("tableView");

  if (overviewView)
    overviewView.style.display = view === "overview" ? "" : "none";
  if (tableView) tableView.style.display = view === "overview" ? "none" : "";

  localStorage.setItem("activeView", view);

  history.replaceState(
    null,
    "",
    `${location.pathname}${location.search}#${view}`,
  );

  await loadCurrentViewData();
  render();
}

function applyFiltersFromPanel() {
  qsa("#advancedFiltersInner [data-col]").forEach((inp) => {
    store.columnFilters[inp.dataset.col] = inp.value;
  });
  store.page = 1;
  render();
}

function resetFilters() {
  store.columnFilters = {};
  store.globalSearch = "";
  store.page = 1;
  el("globalSearch").value = "";
  qsa("#advancedFiltersInner [data-col]").forEach((inp) => (inp.value = ""));
  render();
}

// ================================================================
// MODAL FORM
// ================================================================

let currentFormMode = "add";
let currentFormId = null;

function openFormModal(mode, rowId = null) {
  const sheet = SHEET_MAP[store.view];
  if (!sheet) return;

  const cols = views[store.view].columns.filter((c) => {
  if (c.key === "createdAt" || c.key === "updatedAt") return false;

  if (store.view === "schedule" && /^id$/i.test(String(c.key))) return false;
  if (c.key === "id") return false;

  return true;
});

  if (!cols.length) {
    toast("error", "No columns found", "Check your data structure");
    return;
  }

  const modal = el("modalOverlay");
  const title = el("modalTitle");
  const formFields = el("formFields");
  const submitBtn = el("submitBtnText");
  const form = el("modalForm");

  currentFormMode = mode;
  currentFormId = rowId;

  const singular = views[store.view].title.endsWith("s")
    ? views[store.view].title.slice(0, -1)
    : views[store.view].title;

  if (mode === "add") {
    title.textContent = `Add ${singular}`;
    submitBtn.textContent = "Create";
  } else {
    title.textContent = `Edit ${singular}`;
    submitBtn.textContent = "Update";
  }

  const currentRow =
    mode === "edit" && rowId
      ? (store.view === "sites"
          ? store.data.find(
              (r) =>
                String(r.id || "").trim() === String(rowId).trim() ||
                String(r.Location || r.location || "")
                  .trim()
                  .toLowerCase() === String(rowId).trim().toLowerCase(),
            )
          : store.view === "schedule"
            ? store.data.find((r) => {
                const rid = String(rowId).trim();
                const a = r.id ?? r.ID ?? r.Id ?? "";
                return String(a).trim() === rid;
              })
            : store.data.find((r) => String(r.id) === String(rowId))) || {}
      : {};

  formFields.className =
    cols.length > 4 ? "form-fields two-col" : "form-fields";

  const isEng = store.view === "engagements";
  const isSch = store.view === "schedule";
  const sites = isEng ? getSiteNames() : [];
  const employees = isEng ? getEmployeesForPicker() : [];

  const keyNorm = (k) =>
    String(k || "")
      .toLowerCase()
      .replace(/[\s_]+/g, "");
  const isEmpIdKey = (k) =>
    /(^empid$)|(^employeeid$)|employee.*id/.test(keyNorm(k));
  const isStartKey = (k) =>
    /(^startdate$)|(^startingdate$)|start.*date/.test(keyNorm(k));
  const isEndKey = (k) => /(^enddate$)|end.*date/.test(keyNorm(k));
  const isDurationKey = (k) => /duration/.test(keyNorm(k));

  const toISODateInput = (val) => {
    if (!val) return "";
    const s = String(val).trim();

    // if already YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    // try mm/dd/yyyy
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const mm = String(m[1]).padStart(2, "0");
      const dd = String(m[2]).padStart(2, "0");
      return `${m[3]}-${mm}-${dd}`;
    }

    // fallback Date parse
    const d = new Date(s);
    if (isNaN(d.getTime())) return "";
    return d.toISOString().slice(0, 10);
  };

  const todayISO = () => {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  };

  const daysBetweenInclusive = (startISO, endISO) => {
    if (!startISO || !endISO) return "";
    const start = new Date(`${startISO}T00:00:00`);
    const end = new Date(`${endISO}T00:00:00`);
    const diff = Math.floor((end - start) / 86400000);
    if (isNaN(diff)) return "";
    return String(diff + 1);
  };

  const listId = `emp_list_${Math.random().toString(16).slice(2)}`;

  formFields.innerHTML = cols
    .map((col) => {
      const rawValue = currentRow[col.key] ?? "";
      const value = String(rawValue ?? "");
      const k = col.key;

      const isTextArea =
        String(k).toLowerCase().includes("description") ||
        String(k).toLowerCase().includes("notes");

      const fullWidth = isTextArea;

      if (isEng && isEmployeeNameKey(k)) {
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label required" for="field_${esc(k)}">${esc(col.label)}</label>
          <input
            type="text"
            id="field_${esc(k)}"
            name="${esc(k)}"
            class="form-input"
            list="${listId}"
            value="${esc(value)}"
            required
            placeholder="Search employee by name or EMP-00001"
            autocomplete="off"
            spellcheck="false"
          />
          <datalist id="${listId}">
            ${employees
              .map((e) => {
                const label = e.desig
                  ? `${e.name} — ${e.desig} (${e.id})`
                  : `${e.name} (${e.id})`;
                return `<option value="${esc(e.name)}" label="${esc(label)}"></option>`;
              })
              .join("")}
          </datalist>
          <span class="form-error">This field is required</span>
        </div>`;
      }

      if (isEng && isSiteKey(k)) {
        const current = String(value || "").trim();
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label required" for="field_${esc(k)}">${esc(col.label)}</label>
          <select id="field_${esc(k)}" name="${esc(k)}" class="form-select" required>
            <option value="">Select site...</option>
            ${sites
              .map(
                (s) =>
                  `<option value="${esc(s)}" ${s === current ? "selected" : ""}>${esc(s)}</option>`,
              )
              .join("")}
          </select>
          <span class="form-error">This field is required</span>
        </div>`;
      }

      if (isEng && isStatusKey(k)) {
        const v = String(value || "")
          .trim()
          .toLowerCase();
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label required" for="field_${esc(k)}">${esc(col.label)}</label>
          <select id="field_${esc(k)}" name="${esc(k)}" class="form-select" required>
            <option value="">Select status...</option>
            <option value="Active" ${v === "active" ? "selected" : ""}>Active</option>
            <option value="Completed" ${v === "completed" ? "selected" : ""}>Completed</option>
          </select>
          <span class="form-error">This field is required</span>
        </div>`;
      }

      if (store.view === "sites" && isStatusKey(k)) {
        const v = String(value || "")
          .trim()
          .toLowerCase();
        const isSelected = (x) => v === String(x).toLowerCase();
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label required" for="field_${esc(k)}">${esc(col.label)}</label>
          <select id="field_${esc(k)}" name="${esc(k)}" class="form-select" required>
            <option value="">Select status...</option>
            ${SITE_STATUS_OPTIONS.map(
              (opt) =>
                `<option value="${esc(opt)}" ${isSelected(opt) ? "selected" : ""}>${esc(opt)}</option>`,
            ).join("")}
          </select>
          <span class="form-error">This field is required</span>
        </div>`;
      }

      if (isSch && isStatusKey(k)) {
        const v =
          mode === "add"
            ? "tbd"
            : String(value || "")
                .trim()
                .toLowerCase();
        const isSelected = (x) => v === String(x).toLowerCase();
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label required" for="field_${esc(k)}">${esc(col.label)}</label>
          <select id="field_${esc(k)}" name="${esc(k)}" class="form-select" required>
            <option value="">Select status...</option>
            <option value="Active" ${isSelected("Active") ? "selected" : ""}>Active</option>
            <option value="Completed" ${isSelected("Completed") ? "selected" : ""}>Completed</option>
            <option value="Pending" ${isSelected("Pending") ? "selected" : ""}>Pending</option>
            <option value="On Hold" ${isSelected("On Hold") ? "selected" : ""}>On Hold</option>
            <option value="Cancelled" ${isSelected("Cancelled") ? "selected" : ""}>Cancelled</option>
          </select>
          <span class="form-error">This field is required</span>
        </div>`;
      }

      if (isEng && (isStartKey(k) || isEndKey(k))) {
        const iso = toISODateInput(value);
        const isEnd = isEndKey(k);
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label ${isEnd ? "" : "required"}" for="field_${esc(k)}">${esc(col.label)}</label>
          <input
            type="date"
            id="field_${esc(k)}"
            name="${esc(k)}"
            class="form-input"
            value="${esc(iso)}"
            ${isEnd ? "" : "required"}
          />
          ${isEnd ? "" : `<span class="form-error">This field is required</span>`}
        </div>`;
      }

      if (isSch && (isStartKey(k) || isEndKey(k))) {
        const iso = toISODateInput(value);
        const isEnd = isEndKey(k);
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label ${isEnd ? "" : "required"}" for="field_${esc(k)}">${esc(col.label)}</label>
          <input
            type="date"
            id="field_${esc(k)}"
            name="${esc(k)}"
            class="form-input"
            value="${esc(iso)}"
            ${isEnd ? "" : "required"}
          />
          ${isEnd ? "" : `<span class="form-error">This field is required</span>`}
        </div>`;
      }

      if (isSch && isDurationKey(k)) {
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label" for="field_${esc(k)}">${esc(col.label)}</label>
          <input
            type="number"
            id="field_${esc(k)}"
            name="${esc(k)}"
            class="form-input"
            value="${esc(value)}"
            readonly
            placeholder="Auto"
          />
        </div>`;
      }

      if (isEng && isDurationKey(k)) {
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label" for="field_${esc(k)}">${esc(col.label)}</label>
          <input
            type="number"
            id="field_${esc(k)}"
            name="${esc(k)}"
            class="form-input"
            value="${esc(value)}"
            readonly
            placeholder="Auto"
          />
        </div>`;
      }

      if (isEng && isEmpIdKey(k)) {
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label required" for="field_${esc(k)}">${esc(col.label)}</label>
          <input
            type="text"
            id="field_${esc(k)}"
            name="${esc(k)}"
            class="form-input"
            value="${esc(value)}"
            readonly
            required
            placeholder="Auto from employee"
          />
          <span class="form-error">This field is required</span>
        </div>`;
      }

      if (col.type === "date") {
        const iso = toISODateInput(value);
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label" for="field_${esc(k)}">${esc(col.label)}</label>
          <input type="date" id="field_${esc(k)}" name="${esc(k)}" class="form-input" value="${esc(iso)}" />
        </div>`;
      }

      if (col.type === "number") {
        return `
        <div class="form-group ${fullWidth ? "full-width" : ""}">
          <label class="form-label" for="field_${esc(k)}">${esc(col.label)}</label>
          <input type="number" id="field_${esc(k)}" name="${esc(k)}" class="form-input" value="${esc(value)}" step="any" />
        </div>`;
      }

      if (isTextArea) {
        return `
        <div class="form-group full-width">
          <label class="form-label" for="field_${esc(k)}">${esc(col.label)}</label>
          <textarea id="field_${esc(k)}" name="${esc(k)}" class="form-textarea" placeholder="Enter ${esc(col.label).toLowerCase()}...">${esc(value)}</textarea>
        </div>`;
      }

      const required =
        String(k).toLowerCase().includes("name") ||
        String(k).toLowerCase().includes("title");

      return `
      <div class="form-group ${fullWidth ? "full-width" : ""}">
        <label class="form-label ${required ? "required" : ""}" for="field_${esc(k)}">${esc(col.label)}</label>
        <input
          type="text"
          id="field_${esc(k)}"
          name="${esc(k)}"
          class="form-input"
          value="${esc(value)}"
          ${required ? "required" : ""}
          placeholder="Enter ${esc(col.label).toLowerCase()}..."
        />
        ${required ? `<span class="form-error">This field is required</span>` : ""}
      </div>`;
    })
    .join("");

  // Show modal
  modal.style.display = "flex";
  document.body.style.overflow = "hidden";

  if (isEng) {
    const inputs = Array.from(
      formFields.querySelectorAll("input[name], select[name], textarea[name]"),
    );

    const empNameEl = inputs.find(
      (x) => x.tagName === "INPUT" && isEmployeeNameKey(x.name),
    );
    const empIdEl = inputs.find(
      (x) => x.tagName === "INPUT" && isEmpIdKey(x.name),
    );
    const startEl = inputs.find(
      (x) => x.tagName === "INPUT" && isStartKey(x.name),
    );
    const endEl = inputs.find((x) => x.tagName === "INPUT" && isEndKey(x.name));
    const durEl = inputs.find(
      (x) => x.tagName === "INPUT" && isDurationKey(x.name),
    );
    const statusEl = inputs.find(
      (x) => x.tagName === "SELECT" && isStatusKey(x.name),
    );

    if (empIdEl) empIdEl.readOnly = true;
    if (durEl) durEl.readOnly = true;

    const syncEmpId = () => {
      if (!empNameEl || !empIdEl) return;
      const emp = resolveEmployeeFromInput(empNameEl.value);
      empIdEl.value = emp ? String(emp.id || "") : "";
    };

    const applyStatusRules = () => {
      if (!statusEl || !endEl) return;
      const status = String(statusEl.value || "")
        .trim()
        .toLowerCase();

      const lockEnd = status !== "completed";

      if (lockEnd) {
        endEl.value = "";
        endEl.required = false;
        endEl.classList.add("is-locked");
        endEl.setAttribute("aria-disabled", "true");
        endEl.tabIndex = -1;
      } else {
        endEl.required = true;
        endEl.classList.remove("is-locked");
        endEl.removeAttribute("aria-disabled");
        endEl.tabIndex = 0;
      }
    };

    const recomputeDuration = () => {
      if (!durEl || !startEl || !statusEl) return;

      const status = String(statusEl.value || "")
        .trim()
        .toLowerCase();
      const start = startEl.value;

      if (!start) {
        durEl.value = "";
        return;
      }

      if (status === "completed") {
        const end = endEl ? endEl.value : "";
        durEl.value = end ? daysBetweenInclusive(start, end) : "";
      } else {
        durEl.value = daysBetweenInclusive(start, todayISO());
      }
    };

    if (empNameEl) {
      empNameEl.addEventListener("input", syncEmpId);
      empNameEl.addEventListener("change", syncEmpId);
      syncEmpId();
    }

    if (statusEl) {
      statusEl.addEventListener("change", () => {
        applyStatusRules();
        recomputeDuration();
      });
    }

    if (startEl) startEl.addEventListener("change", recomputeDuration);
    if (endEl) endEl.addEventListener("change", recomputeDuration);

    // init
    applyStatusRules();
    recomputeDuration();

    const timer = setInterval(() => {
      if (!modal || modal.style.display !== "flex") {
        clearInterval(timer);
        return;
      }
      const status = String(statusEl?.value || "")
        .trim()
        .toLowerCase();
      if (status !== "completed") recomputeDuration();
    }, 60000);
  }

  if (isSch) {
    const inputs = Array.from(
      formFields.querySelectorAll("input[name], select[name], textarea[name]"),
    );

    const startEl = inputs.find(
      (x) => x.tagName === "INPUT" && isStartKey(x.name),
    );
    const endEl = inputs.find((x) => x.tagName === "INPUT" && isEndKey(x.name));
    const durEl = inputs.find(
      (x) => x.tagName === "INPUT" && isDurationKey(x.name),
    );
    const statusEl = inputs.find(
      (x) => x.tagName === "SELECT" && isStatusKey(x.name),
    );

    if (durEl) durEl.readOnly = true;

    const applyStatusRulesSch = () => {
      if (!statusEl || !endEl) return;
      const status = String(statusEl.value || "")
        .trim()
        .toLowerCase();

      const lockEnd = status !== "completed";

      if (lockEnd) {
        endEl.value = "";
        endEl.required = false;
        endEl.classList.add("is-locked");
        endEl.setAttribute("aria-disabled", "true");
        endEl.tabIndex = -1;
      } else {
        endEl.required = true;
        endEl.classList.remove("is-locked");
        endEl.removeAttribute("aria-disabled");
        endEl.tabIndex = 0;
      }
    };

    const recomputeDurationSch = () => {
      if (!durEl || !startEl) return;

      const start = startEl.value;
      if (!start) {
        durEl.value = "";
        return;
      }

      const status = String(statusEl?.value || "")
        .trim()
        .toLowerCase();

      if (status === "completed") {
        const end = endEl ? endEl.value : "";
        durEl.value = end ? daysBetweenInclusive(start, end) : "";
      } else {
        durEl.value = daysBetweenInclusive(start, todayISO());
      }
    };

    if (statusEl) {
      statusEl.addEventListener("change", () => {
        applyStatusRulesSch();
        recomputeDurationSch();
      });
    }

    if (startEl) startEl.addEventListener("change", recomputeDurationSch);
    if (endEl) endEl.addEventListener("change", recomputeDurationSch);

    applyStatusRulesSch();
    recomputeDurationSch();

    const timer = setInterval(() => {
      if (!modal || modal.style.display !== "flex") {
        clearInterval(timer);
        return;
      }
      const status = String(statusEl?.value || "")
        .trim()
        .toLowerCase();
      if (status !== "completed") recomputeDurationSch();
    }, 60000);
  }
  // focus first field
  setTimeout(() => {
    const first = formFields.querySelector("input, select, textarea");
    if (first) first.focus();
  }, 50);
}

function closeFormModal() {
  const modal = el("modalOverlay");
  modal.style.display = "none";
  document.body.style.overflow = "";
  el("modalForm").reset();

  qsa(".form-group").forEach((g) => g.classList.remove("error"));
}

async function handleFormSubmit(e) {
  e.preventDefault();

  const sheet = SHEET_MAP[store.view];
  if (!sheet) return;

  const form = el("modalForm");
  const formData = new FormData(form);
  const data = {};

  for (const [key, value] of formData.entries()) {
    data[key] = value;
  }

  let hasError = false;
  qsa(".form-group").forEach((group) => {
    group.classList.remove("error");
    const input = group.querySelector(
      ".form-input, .form-textarea, .form-select",
    );
    if (input && input.required && !String(input.value || "").trim()) {
      group.classList.add("error");
      hasError = true;
    }
  });

  if (hasError) {
    toast("error", "Please fill all required fields");
    return;
  }

  const submitBtn = el("btnSubmitForm");
  const originalText = submitBtn.innerHTML;
  submitBtn.disabled = true;
  submitBtn.innerHTML = "<span>Saving...</span>";

  try {
    if (store.view === "engagements") {
      const empNameKey = Object.keys(data).find((k) => isEmployeeNameKey(k));
      const typed = empNameKey ? data[empNameKey] : "";
      const emp = resolveEmployeeFromInput(typed);
      if (!emp)
        throw new Error("Employee not found. Please select valid employee.");

      if (empNameKey) data[empNameKey] = emp.name;

      const empIdKey =
        Object.keys(data).find((k) =>
          String(k).toLowerCase().replace(/\s+/g, "").includes("employeeid"),
        ) ||
        Object.keys(data).find((k) =>
          String(k).toLowerCase().includes("employee id"),
        ) ||
        "Employee ID";

      data[empIdKey] = emp.id;

      delete data.id;
    }

    if (currentFormMode === "add") {

  // ✅ schedule ONLY: never send any ID to backend
  if (store.view === "schedule") {
    delete data.ID;
    delete data.Id;
    delete data.id;
  }

  await apiCreateRow(sheet, store.view, data);
  toast("success", "Record created successfully");

} else {
  const updateId =
    store.view === "sites"
      ? data.Location || data.location || currentFormId
      : currentFormId;

  await apiUpdateRow(sheet, updateId, data);
  toast("success", "Record updated successfully");
}

    closeFormModal();
    await loadCurrentViewData();
    await loadAllData();
  } catch (e2) {
    toast("error", "Save failed", e2.message || "Unknown error");
  } finally {
    submitBtn.disabled = false;
    submitBtn.innerHTML = originalText;
  }
}

// ================================================================
// EVENT LISTENERS
// ================================================================

document.addEventListener("click", async (e) => {
  const navBtn = e.target.closest(".nav-link");
  if (navBtn && navBtn.dataset.view) {
    await setView(navBtn.dataset.view);
    return;
  }

  const chip = e.target.closest(".filter-chip-remove");
  if (chip) {
    if (chip.dataset.chipType === "search") {
      store.globalSearch = "";
      el("globalSearch").value = "";
    } else if (chip.dataset.chipCol) {
      store.columnFilters[chip.dataset.chipCol] = "";
    }
    store.page = 1;
    render();
    return;
  }

  const th = e.target.closest("th[data-sort]");
  if (th && th.dataset.sort) {
    const col = th.dataset.sort;
    if (store.sortCol === col)
      store.sortDir = store.sortDir === "asc" ? "desc" : "asc";
    else {
      store.sortCol = col;
      store.sortDir = "asc";
    }
    render();
    return;
  }

  const btn = e.target.closest("button[data-act]");
  if (!btn) return;

  const act = btn.dataset.act;
  const id = btn.dataset.id;
  const sheet = SHEET_MAP[store.view];
  if (!sheet) return;

  if (act === "export") {
    window.open(`/api/employees/${encodeURIComponent(id)}/export`, "_blank");
    return;
  }

  if (act === "del") {
    if (!confirm(`Delete record ${id}?`)) return;
    try {
      await apiDeleteRow(sheet, id);
      toast("success", "Record deleted");
      await loadCurrentViewData();
      await loadAllData();
    } catch (err) {
      toast("error", "Delete failed", err.message || "Unknown error");
    }
    return;
  }

  if (act === "edit") {
    openFormModal("edit", id);
    return;
  }
});

let searchTimer;
el("globalSearch").addEventListener("input", (e) => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    store.globalSearch = e.target.value;
    store.page = 1;
    render();
  }, 200);
});

el("clearSearch").addEventListener("click", () => {
  store.globalSearch = "";
  el("globalSearch").value = "";
  store.page = 1;
  render();
});

el("btnToggleFilters").addEventListener("click", () => {
  const panel = el("advancedFilters");
  const isOpen = panel.style.display !== "none";
  panel.style.display = isOpen ? "none" : "";
  if (!isOpen) renderAdvancedFilterPanel();
});

el("btnApplyFilters").addEventListener("click", applyFiltersFromPanel);
el("btnResetFilters").addEventListener("click", resetFilters);
el("btnClearAll").addEventListener("click", resetFilters);

el("advancedFilters").addEventListener("keydown", (e) => {
  if (e.key === "Enter") applyFiltersFromPanel();
});

el("pageSizeSelect").addEventListener("change", (e) => {
  store.pageSize = parseInt(e.target.value, 10);
  store.page = 1;
  render();
});

el("btnPrevPage").addEventListener("click", () => {
  store.page--;
  render();
});

el("btnNextPage").addEventListener("click", () => {
  store.page++;
  render();
});

el("modalForm").addEventListener("submit", handleFormSubmit);
el("modalClose").addEventListener("click", closeFormModal);
el("btnCancelForm").addEventListener("click", closeFormModal);

el("modalOverlay").addEventListener("click", (e) => {
  if (e.target === el("modalOverlay")) closeFormModal();
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && el("modalOverlay").style.display === "flex") {
    closeFormModal();
  }
});

// ================================================================
// INIT
// ================================================================

async function init() {
  mountToaster();
  await loadAllData();
  mountToolbarButtons();
  wireGanttControls();

  const initialView =
    getViewFromHash() || localStorage.getItem("activeView") || DEFAULT_VIEW;

  await setView(initialView);

  // profile menu wiring (after DOM is ready)
  const btn = document.getElementById("profileBtn");
  const menu = document.getElementById("profileMenu");
  const logoutBtn = document.getElementById("logoutBtn");

  if (btn && menu) {
    btn.addEventListener("click", () => menu.classList.toggle("show"));

    document.addEventListener("click", (e) => {
      if (!btn.contains(e.target) && !menu.contains(e.target)) {
        menu.classList.remove("show");
      }
    });
  }

  logoutBtn?.addEventListener("click", async () => {
    await fetch("/api/auth/logout", { method: "POST", credentials: "include" });
    window.location.href = "/login";
  });
}

document.addEventListener("DOMContentLoaded", () => {
  init().catch((err) => console.error("init failed:", err));
});
