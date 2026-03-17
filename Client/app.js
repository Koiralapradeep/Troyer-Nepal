"use strict";

const API = "/api";

const ACTIVE_SCHEDULE_STORAGE_KEY = "activeScheduleSheet";

function saveActiveScheduleSheet(sheet) {
  const value = String(sheet || "").trim();
  if (value) localStorage.setItem(ACTIVE_SCHEDULE_STORAGE_KEY, value);
  else localStorage.removeItem(ACTIVE_SCHEDULE_STORAGE_KEY);
}

const SHEET_MAP = {
  overview: "profiles",
  employees: "profiles",
  engagements: "Engagements",
  sites: "Sites_List",
  tools: "Tools",
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
  scheduleProjects: [],
  scheduleProjectRows: [],
  activeScheduleSheet: restoreActiveScheduleSheet(),
  scheduleProjectCache: {},
  scheduleProjectSearch: "",
  scheduleLoading: false,
  scheduleGrids: [],
  activeScheduleGrid: null,
  currentScheduleProjectIndex: 0,
};

const SITE_STATUS_OPTIONS = [
  "Installation",
  "Pre Commissioning",
  "Wet Commissioning",
  "Commissioned",
];

function restoreActiveScheduleSheet() {
  return String(localStorage.getItem(ACTIVE_SCHEDULE_STORAGE_KEY) || "").trim();
}

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
function getActiveScheduleSheet() {
  return store.activeScheduleSheet || store.scheduleProjects[0]?.sheet || "";
}

function getSheetForView(view) {
  if (view === "schedule" || view === "gantt") {
    return getActiveScheduleSheet();
  }
  return SHEET_MAP[view] || "";
}

function prettifyProjectName(sheet) {
  return String(sheet || "")
    .replace(/_timeline$/i, "")
    .replace(/_updated$/i, "")
    .replace(/_/g, " ")
    .trim();
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
  await apiFetch(`${API}/auth/logout`, { method: "POST" });
  localStorage.clear();
  sessionStorage.clear();
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
  tools: { title: "Tools", columns: [] },
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
  const isSchedule = store.view === "schedule" || store.view === "gantt";

  // body.is-schedule drives all CSS overrides for the flex chain
  document.body.classList.toggle("is-schedule", isSchedule);

  const overviewView = document.getElementById("overviewView");
  const tableView = document.getElementById("tableView");
  const scheduleLayout = document.getElementById("scheduleLayout");
  const scheduleScroller = document.getElementById("scheduleScroller");
  const scheduleSheet = document.getElementById("scheduleSheet");
  const tableCard = document.querySelector("#tableView .table-card");
  const tableToolbar = document.querySelector("#tableView .table-toolbar");
  const stats = document.getElementById("stats");
  const filterBar = document.getElementById("filterBar");
  const advancedFilters = document.getElementById("advancedFilters");
  const paginationBar = document.getElementById("paginationBar");
  const defaultWrap = document.getElementById("defaultTableWrap");
  const scheduleHeader = document.getElementById("scheduleHeader");
  const btnOpenGantt = document.getElementById("btnOpenGantt");

  // Step 1: Clear all inline cssText so CSS class rules take over
  [
    tableView,
    tableCard,
    scheduleLayout,
    scheduleScroller,
    scheduleSheet,
  ].forEach((el) => {
    if (el) el.style.cssText = "";
  });
  if (overviewView) overviewView.style.display = isOverview ? "" : "none";
  if (tableView) tableView.style.display = isOverview ? "none" : "";

  if (isSchedule) {
    if (scheduleLayout) scheduleLayout.style.display = "flex";
    if (scheduleScroller) scheduleScroller.style.display = "flex";
    if (scheduleSheet) {
      scheduleSheet.style.display = "flex";
      scheduleSheet.classList.add("schedule-screen");
    }
    if (scheduleHeader) scheduleHeader.style.display = "none";
    if (tableToolbar) tableToolbar.style.display = "";
    if (stats) stats.style.display = "none";
    if (filterBar) filterBar.style.display = "none";
    if (advancedFilters) advancedFilters.style.display = "none";
    if (paginationBar) paginationBar.style.display = "none";
    if (defaultWrap) defaultWrap.style.display = "none";
  } else {
    if (scheduleSheet) {
      scheduleSheet.style.display = "";
      scheduleSheet.classList.remove("schedule-screen");
    }
    if (scheduleLayout) scheduleLayout.style.display = "none";
    if (scheduleScroller) scheduleScroller.style.display = "";
    if (scheduleHeader) scheduleHeader.style.display = "none";
    if (tableCard) tableCard.style.display = isOverview ? "none" : "";
    if (tableToolbar) tableToolbar.style.display = "";
    if (defaultWrap) defaultWrap.style.display = !isOverview ? "" : "none";
    if (stats) stats.style.display = "";
    if (filterBar) filterBar.style.display = "";
    if (advancedFilters) advancedFilters.style.display = "none";
    if (paginationBar) paginationBar.style.display = "";
  }

  if (btnOpenGantt) btnOpenGantt.style.display = "none";
}

function mountScheduleGanttButton() {
  const btn = document.getElementById("btnOpenGantt");
  if (!btn) return;

  btn.style.display = store.view === "schedule" ? "" : "none";
  btn.onclick = async () => {
    store.view = "gantt";
    applyViewLayout();
    await loadCurrentViewData({ resetPage: true });
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
function getToolsParentRows() {
  return (store.data || []).filter((r) => {
    const sn = String(r["S.N"] ?? r["s.n"] ?? "").trim();
    const rowType = String(r.__rowType ?? "")
      .trim()
      .toLowerCase();
    return !!sn && rowType !== "child";
  });
}

function pickValue(row, names) {
  for (const name of names) {
    if (row[name] != null && String(row[name]).trim() !== "") return row[name];
  }

  const keys = Object.keys(row || {});
  for (const wanted of names) {
    const found = keys.find(
      (k) =>
        String(k).trim().toLowerCase() === String(wanted).trim().toLowerCase(),
    );
    if (found && row[found] != null && String(row[found]).trim() !== "") {
      return row[found];
    }
  }

  return "";
}

function toScheduleRow(row, index = 0) {
  const taskName = pickValue(row, [
    "Task Name",
    "TASK NAME",
    "Task Name / Milestone",
    "Milestone",
    "Task",
    "Activity",
    "Activities",
    "Task name",
    "Name",
    "Work Item",
    "Task Mode",
    "ID",
  ]);

  const start =
    pickValue(row, [
      "Start Date",
      "START DATE",
      "Start",
      "START",
      "Starting Date",
      "Planned Start",
      "Begin Date",
      "Begin",
    ]) || "";

  const end =
    pickValue(row, [
      "End Date",
      "END DATE",
      "Finish Date",
      "FINISH DATE",
      "Finish",
      "FINISH",
      "End",
      "END",
      "Planned Finish",
      "Completion Date",
    ]) || "";

  let completion = pickValue(row, [
    "% Completion",
    "% Complete",
    "% complete",
    "Progress",
    "Completion",
    "Progress %",
    "Percent Complete",
  ]);

  let duration =
    pickValue(row, [
      "Duration",
      "Duration (days)",
      "DURATION",
      "Planned Duration",
      "No. of Days",
      "Days",
    ]) || "";

  const remarks = pickValue(row, [
    "Remarks",
    "Remark",
    "Notes",
    "Comments",
    "Comment",
    "Description",
    "Notes / Section",
  ]);

  if (completion != null && String(completion).trim() !== "") {
    const raw = String(completion).trim();
    const n = Number(raw.replace(/[^\d.-]/g, ""));
    if (Number.isFinite(n)) {
      completion =
        n > 0 && n <= 1 ? String(Math.round(n * 100)) : String(Math.round(n));
    } else {
      completion = "0";
    }
  } else {
    completion = "0";
  }

  if (!String(duration).trim() && start && end) {
    duration = String(calcDurationDays(start, end));
  } else if (duration != null && typeof duration === "object") {
    duration = "";
  }

  const out = {
    id: pickValue(row, ["id", "ID", "Id"]) || `SCH-${index + 1}`,
    taskName: taskName || "",
    completion: completion || "0",
    duration: duration || "",
    startDate: start,
    endDate: end,
    remarks: remarks || "",
  };

  if (row.__highlight) out.__highlight = row.__highlight;
  if (row.__isSection) out.__isSection = true;
  if (row.__isTitle) out.__isTitle = true;

  const hasTask = String(out.taskName).trim() !== "";
  const hasDate =
    String(out.startDate).trim() !== "" || String(out.endDate).trim() !== "";
  const hasDuration = String(out.duration).trim() !== "";

  if (!hasTask && !hasDate && !hasDuration) return null;

  return out;
}

function getToolsRowKey(row) {
  return String(
    row.__rowKey ?? row.id ?? row["S.N"] ?? row["s.n"] ?? "",
  ).trim();
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

document.querySelectorAll(".nav-link[data-view]").forEach((btn) => {
  btn.addEventListener("click", async () => {
    const view = btn.dataset.view;
    await navigate(view, { pushHash: true });
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
    const res = await apiFetch(
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

async function fetchWorkbookSheets() {
  const res = await apiFetch(`${API}/excel/sheets`);
  const json = await res.json().catch(() => ({}));
  if (!json.ok) throw new Error(json.error || "Failed to load workbook sheets");
  return Array.isArray(json.sheets) ? json.sheets : [];
}

async function fetchScheduleGrid(sheet) {
  const res = await apiFetch(
    `${API}/excel/grid?sheet=${encodeURIComponent(sheet)}`,
  );
  const json = await res.json().catch(() => ({}));
  if (!json.ok)
    throw new Error(json.error || `Failed to load schedule grid for ${sheet}`);
  return json;
}

// Convert a grid (from /api/excel/grid) into keyed-object rows
// suitable for toScheduleRow(). The grid's first non-empty isHeader
// row supplies column names; subsequent data rows are mapped to those keys.
function gridToScheduleRows(grid) {
  const rawRows = Array.isArray(grid?.rows) ? grid.rows : [];
  if (!rawRows.length) return [];

  // Find the header row (first row where isHeader === true)
  const headerRow = rawRows.find((r) => r.isHeader);
  const headers = headerRow
    ? (headerRow.cells || []).map((c) => String(c.text ?? "").trim())
    : [];

  const results = [];
  rawRows.forEach((row, ri) => {
    if (row.isHeader) return; // skip header rows

    const cells = row.cells || [];
    const obj = {};

    // Map by position to header name
    headers.forEach((h, ci) => {
      if (h) obj[h] = cells[ci]?.text ?? "";
    });

    // Preserve highlight flag (yellow rows have bg #FFFF00 or similar)
    const firstCellBg = String(cells[0]?.bg || "").toLowerCase();
    if (firstCellBg.includes("ffff00") || firstCellBg.includes("fff200")) {
      obj.__highlight = "yellow";
    }

    // Preserve section/title metadata for row-type detection
    obj.__isSection = !!row.isSection;
    obj.__isTitle = !!row.isTitle;

    results.push(obj);
  });

  return results;
}

function isScheduleProjectSheet(name) {
  const s = String(name || "")
    .trim()
    .toLowerCase();

  if (!s) return false;

  if (
    s === "profiles" ||
    s === "engagements" ||
    s === "sites_list" ||
    s === "tools"
  ) {
    return false;
  }

  return (
    s === "seti_updated" || s === "khudi_timeline" || s.endsWith("_timeline")
  );
}

function getScheduleSheetCandidates(allSheets) {
  return (allSheets || []).filter(isScheduleProjectSheet);
}
async function loadScheduleProjects() {
  const workbookSheets = await fetchWorkbookSheets();
  const sheets = getScheduleSheetCandidates(workbookSheets);

  store.scheduleProjects = sheets.map((sheet) => ({
    sheet,
    title: sheet.replace(/_Timeline$/i, "").replace(/_/g, " "),
  }));

  const saved = restoreActiveScheduleSheet();
  if (!store.activeScheduleSheet && saved) {
    store.activeScheduleSheet = saved;
  }

  const exists = store.scheduleProjects.some(
    (p) => p.sheet === store.activeScheduleSheet,
  );

  if (!store.activeScheduleSheet && store.scheduleProjects.length) {
    store.activeScheduleSheet = store.scheduleProjects[0].sheet;
  } else if (store.activeScheduleSheet && !exists) {
    store.activeScheduleSheet = store.scheduleProjects[0]?.sheet || "";
  }

  saveActiveScheduleSheet(store.activeScheduleSheet);
  return store.scheduleProjects;
}

function getActiveScheduleProject() {
  return (
    (store.scheduleProjects || []).find(
      (p) => p.sheet === store.activeScheduleSheet,
    ) || null
  );
}

async function loadScheduleProjectRows(sheet, { force = false } = {}) {
  const targetSheet = String(sheet || "").trim();
  if (!targetSheet) return { sheet: "", title: "", rows: [] };

  if (!force && store.scheduleProjectCache[targetSheet]) {
    return store.scheduleProjectCache[targetSheet];
  }

  let rows = [];

  try {
    const grid = await fetchScheduleGrid(targetSheet);
    const keyed = gridToScheduleRows(grid);
    rows = keyed.map((row, index) => toScheduleRow(row, index)).filter(Boolean);
  } catch (gridErr) {
    console.warn(
      "Grid fetch failed, falling back to rows API:",
      targetSheet,
      gridErr?.message || gridErr,
    );
    try {
      const rawRows = await fetchAllRows(targetSheet);
      rows = rawRows
        .map((row, index) => toScheduleRow(row, index))
        .filter(Boolean);
    } catch (rowErr) {
      console.error("Both schedule fetch methods failed:", targetSheet, rowErr);
      rows = [];
    }
  }

  const project = (store.scheduleProjects || []).find(
    (p) => p.sheet === targetSheet,
  );
  const result = {
    sheet: targetSheet,
    title: project?.title || prettifyProjectName(targetSheet),
    rows,
  };

  store.scheduleProjectCache[targetSheet] = result;
  return result;
}

async function loadAllScheduleProjectRows() {
  await loadScheduleProjects();
  const results = [];
  for (const project of store.scheduleProjects || []) {
    results.push(await loadScheduleProjectRows(project.sheet));
  }
  store.scheduleProjectRows = results;
  return results;
}

async function refreshScheduleState(
  sheet = getActiveScheduleSheet(),
  { resetPage = false } = {},
) {
  await loadScheduleProjects();

  const targetSheet = String(sheet || getActiveScheduleSheet() || "").trim();
  if (targetSheet) {
    store.activeScheduleSheet = targetSheet;
    saveActiveScheduleSheet(targetSheet);
    delete store.scheduleProjectCache?.[targetSheet];
  }

  const activeProject = getActiveScheduleProject();
  const projectData = activeProject
    ? await loadScheduleProjectRows(activeProject.sheet, { force: true })
    : { sheet: "", title: "", rows: [] };

  store.scheduleProjectRows = projectData.sheet ? [projectData] : [];
  store.data = Array.isArray(projectData.rows) ? projectData.rows.slice() : [];
  store.allSchedule = store.data.slice();
  views.schedule.columns = inferColumnsFromData(store.data, "schedule");
  views.gantt.columns = inferColumnsFromData(store.data, "gantt");

  if (resetPage) store.page = 1;
  render();
  return projectData;
}

async function loadScheduleGrids() {
  const workbookSheets = await fetchWorkbookSheets().catch(() => []);
  const targets = getScheduleSheetCandidates(workbookSheets);

  console.log("All workbook sheets:", workbookSheets);
  console.log("Schedule target sheets:", targets);

  const grids = [];
  for (const sheet of targets) {
    try {
      const grid = await fetchScheduleGrid(sheet);
      console.log("Loaded schedule grid:", sheet, grid);
      if (grid?.rows?.length) grids.push(grid);
    } catch (err) {
      console.error("schedule grid load failed:", sheet, err);
    }
  }

  store.scheduleGrids = grids;
  return grids;
}

async function apiCreateRow(sheet, view, data) {
  const idPrefix = view === "schedule" ? "" : ID_PREFIX_MAP[view] || "ROW";
  const res = await apiFetch(
    `${API}/rows?sheet=${encodeURIComponent(sheet)}&idPrefix=${encodeURIComponent(idPrefix)}`,
    {
      method: "POST",
      body: JSON.stringify(data),
    },
  );
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Create failed");
  return json.item;
}

async function apiUpdateRow(sheet, id, patch) {
  const key = sheet === "Sites_List" ? "Location" : "id";

  const res = await apiFetch(
    `${API}/rows/${encodeURIComponent(id)}?sheet=${encodeURIComponent(sheet)}&key=${encodeURIComponent(key)}`,
    {
      method: "PUT",
      body: JSON.stringify(patch),
    },
  );

  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Update failed");
  return json.item;
}

async function apiDeleteRow(sheet, id) {
  const res = await apiFetch(
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
    const d = parseAnyDate(v);
    return d && !isNaN(d.getTime()) ? startOfDay(d) : null;
  };

  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));
  const toISO = (d) => d.toISOString().slice(0, 10);

  const tasks = (Array.isArray(rows) ? rows : [])
    .map((r, i) => {
      const s = parse(r.startDate);
      const f = parse(r.endDate);
      const durStr = String(r.duration ?? "").trim();
      const pct = parseFloat(String(r.completion ?? 0).replace("%", "")) || 0;

      const name = String(r.taskName ?? "").trim();

      let finish = f;
      if (!finish && s && durStr) {
        const durNum = parseFloat(durStr);
        if (Number.isFinite(durNum) && durNum > 0) {
          finish = new Date(s);
          finish.setDate(
            finish.getDate() + Math.max(0, Math.round(durNum) - 1),
          );
        }
      }

      const isMilestone = !!(s && finish && +s === +finish);

      return {
        id: r.id || i + 1,
        name,
        start: s,
        finish,
        pct: clamp(pct, 0, 100),
        duration: durStr,
        predecessors: "",
        isSummary: false,
        isMilestone,
      };
    })
    .filter((t) => t.start && t.finish);

  if (!tasks.length) {
    container.innerHTML =
      "<div style='padding:18px;font-weight:800'>No valid gantt data</div>";
    return;
  }

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

  const isoWeek = (date) => {
    const d = new Date(
      Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()),
    );
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  };

  const wrap = document.createElement("div");
  wrap.className = "gantt-container";

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
      <div>${esc(t.id)}</div>
      <div class="gantt-taskname">${esc(t.name)}</div>
      <div>${esc(toISO(t.start))}</div>
      <div>${esc(toISO(t.finish))}</div>
      <div>${Math.round(t.pct)}%</div>
    `;
    leftBody.appendChild(row);
  });

  left.appendChild(leftBody);

  const rightOuter = document.createElement("div");
  rightOuter.className = "gantt-right-outer";

  const head = document.createElement("div");
  head.className = "gantt-right-head";

  const yearsRow = document.createElement("div");
  yearsRow.className = "gantt-head-years";
  yearsRow.style.width = timelineWidth + "px";

  const weeksRow = document.createElement("div");
  weeksRow.className = "gantt-head-weeks";
  weeksRow.style.width = timelineWidth + "px";

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

  const right = document.createElement("div");
  right.className = "gantt-right";
  right.style.width = timelineWidth + "px";

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

  const today = startOfDay(new Date());
  if (today >= min && today <= max) {
    const offset = Math.floor((today - min) / 86400000);
    const tline = document.createElement("div");
    tline.className = "gantt-today-line";
    tline.style.left = offset * dayWidth + "px";
    right.appendChild(tline);
  }

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
    bar.className = "gantt-bar";
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

  const legend = document.createElement("div");
  legend.className = "gantt-legend";
  legend.innerHTML = `
    <div class="gl-item"><span class="gl-swatch bar"></span>Task</div>
    <div class="gl-item"><span class="gl-swatch milestone"></span>Milestone</div>
    <div class="gl-item"><span class="gl-swatch progress"></span>Progress</div>
    <div class="gl-item"><span class="gl-swatch today"></span>Today</div>
  `;

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
  rows = Array.isArray(rows) ? rows : [];
  if (viewName === "schedule" || viewName === "gantt") {
    return [
      {
        key: "taskName",
        label: "Task Name",
        type: "text",
        filterable: true,
        sortable: true,
      },
      {
        key: "completion",
        label: "% Completion",
        type: "number",
        filterable: true,
        sortable: true,
      },
      {
        key: "duration",
        label: "Duration",
        type: "number",
        filterable: true,
        sortable: true,
      },
      {
        key: "startDate",
        label: "Start Date",
        type: "date",
        filterable: true,
        sortable: true,
      },
      {
        key: "endDate",
        label: "End Date",
        type: "date",
        filterable: true,
        sortable: true,
      },
      {
        key: "remarks",
        label: "Remarks",
        type: "text",
        filterable: true,
        sortable: true,
      },
    ];
  }

  const first = rows[0] || {};

  return Object.keys(first)
    .filter((key) => {
      const lower = String(key).trim().toLowerCase();

      if (viewName === "sites" && lower === "id") return false;

      if (viewName === "tools" && lower === "id") return false;
      if (viewName === "tools" && lower === "__rowkey") return false;
      if (viewName === "tools" && lower === "__rowtype") return false;
      if (viewName === "tools" && lower === "__parentsn") return false;

      return true;
    })
    .map((key) => {
      const lk = String(key).trim().toLowerCase();

      let type = "text";

      if (
        lk.includes("date") ||
        lk.includes("createdat") ||
        lk.includes("updatedat")
      ) {
        type = "date";
      }

      if (
        lk === "s.n" ||
        lk === "sn" ||
        lk === "s n" ||
        lk === "qty" ||
        lk === "quantity" ||
        lk === "id"
      ) {
        type = "number";
      }

      return {
        key,
        label: key.replace(/_/g, " ").replace(/\b\w/g, (m) => m.toUpperCase()),
        type,
        filterable: key !== "id",
        sortable: true,
      };
    });
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
  return Math.max(0, Math.round((f - s) / 86400000) + 1);
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
    "End Date",
    "END DATE",
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

  if (name && !start && !finish && (dur === "" || dur === 0)) return "Summary";

  const s = toDateOnly(start);
  const f = toDateOnly(finish);

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
    const projectRows = await loadAllScheduleProjectRows();
    const firstScheduleSheet = getActiveScheduleSheet();

    const [employees, engagements, sites, schedule, tools] = await Promise.all([
      fetchAllRows(SHEET_MAP.employees).catch(() => []),
      fetchAllRows(SHEET_MAP.engagements).catch(() => []),
      fetchAllRows(SHEET_MAP.sites).catch(() => []),
      firstScheduleSheet
        ? fetchScheduleGrid(firstScheduleSheet)
            .then((grid) => {
              const keyed = gridToScheduleRows(grid);
              return keyed
                .map((row, index) => toScheduleRow(row, index))
                .filter(Boolean);
            })
            .catch(() =>
              fetchAllRows(firstScheduleSheet)
                .then((rows) =>
                  rows.map((r, i) => toScheduleRow(r, i)).filter(Boolean),
                )
                .catch(() => []),
            )
        : Promise.resolve([]),
      fetchAllRows(SHEET_MAP.tools).catch(() => []),
    ]);

    const safeEmployees = Array.isArray(employees) ? employees : [];
    const safeEngagements = Array.isArray(engagements) ? engagements : [];
    const safeSites = Array.isArray(sites) ? sites : [];
    const safeSchedule = Array.isArray(schedule) ? schedule : [];
    const safeTools = Array.isArray(tools) ? tools : [];

    store.allEmployees = safeEmployees;
    store.allEngagements = safeEngagements;
    store.allSites = safeSites;
    store.allSchedule = safeSchedule;
    store.allTools = safeTools;
    store.scheduleProjectRows = Array.isArray(projectRows) ? projectRows : [];

    views.employees.columns = inferColumnsFromData(safeEmployees, "employees");
    views.engagements.columns = inferColumnsFromData(
      safeEngagements,
      "engagements",
    );
    views.sites.columns = inferColumnsFromData(safeSites, "sites");
    views.schedule.columns = inferColumnsFromData(safeSchedule, "schedule");
    views.gantt.columns = inferColumnsFromData(safeSchedule, "gantt");
    views.tools.columns = inferColumnsFromData(safeTools, "tools");

    const hasData =
      safeEmployees.length ||
      safeEngagements.length ||
      safeSites.length ||
      safeSchedule.length ||
      safeTools.length ||
      store.scheduleProjectRows.length;

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

async function loadCurrentViewData({
  resetPage = false,
  goToLastPage = false,
} = {}) {
  const sheet = getSheetForView(store.view);

  try {
    if (store.view === "schedule" || store.view === "gantt") {
      await loadScheduleProjects();
      const activeProject = getActiveScheduleProject();
      const projectData = activeProject
        ? await loadScheduleProjectRows(activeProject.sheet)
        : { sheet: "", title: "", rows: [] };

      store.scheduleProjectRows = projectData.sheet ? [projectData] : [];
      store.data = Array.isArray(projectData.rows)
        ? projectData.rows.slice()
        : [];
      store.allSchedule = store.data.slice();

      views.schedule.columns = inferColumnsFromData(store.data, "schedule");
      views.gantt.columns = inferColumnsFromData(store.data, "gantt");

      const totalPages = Math.max(
        1,
        Math.ceil((store.data.length || 0) / store.pageSize),
      );

      if (resetPage) {
        store.page = 1;
      } else if (goToLastPage) {
        store.page = totalPages;
      } else if (store.page > totalPages) {
        store.page = totalPages;
      }

      render();
      return;
    }

    let rows = await fetchAllRows(sheet);

    store.data = rows;
    if (store.view === "employees") store.allEmployees = rows;
    if (store.view === "engagements") store.allEngagements = rows;
    if (store.view === "sites") store.allSites = rows;
    if (store.view === "tools") store.allTools = rows;

    views[store.view].columns = inferColumnsFromData(rows, store.view);
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

function toNum(v) {
  const n = Number(String(v ?? "").trim());
  return Number.isFinite(n) ? n : null;
}

async function downloadExcelExport() {
  const url = `${API}/excel/download?mode=full&ts=${Date.now()}`;

  try {
    const res = await apiFetch(url);

    if (!res.ok) {
      let msg = "Export failed";
      try {
        const json = await res.json();
        msg = json.error || msg;
      } catch {}
      throw new Error(msg);
    }

    const blob = await res.blob();
    const objectUrl = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = objectUrl;
    a.download = "troyer_full_dashboard_export.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);

    toast("success", "Excel downloaded successfully");
  } catch (err) {
    toast("error", "Excel export failed", err.message || "Download failed");
  }
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

function formatGridDateText(value) {
  const d = value ? new Date(value) : null;
  if (!d || isNaN(d.getTime())) return String(value ?? "");
  const month = d.toLocaleString("en-US", { month: "short" });
  const day = String(d.getDate());
  const yy = String(d.getFullYear()).slice(-2);
  return `${day}-${month}-${yy}`;
}

function escapeAttr(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ================================================================
// SCHEDULE GANTT — Definitive implementation
// ================================================================

(function _sgCSS() {
  if (document.getElementById("_sg_css")) return;
  const s = document.createElement("style");
  s.id = "_sg_css";
  s.textContent = `

/* ── Shell ─────────────────────────────────────────────────────── */
.sg {
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
  background: #fff;
  font-family: 'Inter', system-ui, sans-serif;
  font-size: 12px;
}

/* ── Project tab bar ────────────────────────────────────────────── */
.sg-topbar {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 16px;
  background: #f4fbf4;
  border-bottom: 2px solid #1d5c2e;
  flex-shrink: 0;
  flex-wrap: wrap;
}
.sg-tab {
  padding: 5px 18px;
  border: 1.5px solid #a8cca8;
  border-radius: 20px;
  background: #fff;
  color: #2e7d32;
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  white-space: nowrap;
  transition: all .15s;
  font-family: inherit;
}
.sg-tab:hover { background: #e8f5e9; border-color: #2e7d32; }
.sg-tab.active {
  background: #1d5c2e;
  color: #fff;
  border-color: #1d5c2e;
  box-shadow: 0 2px 8px rgba(29,92,46,.3);
}
.sg-tab-add {
  padding: 5px 14px;
  border: 1.5px dashed #66bb6a;
  border-radius: 20px;
  background: transparent;
  color: #2e7d32;
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  white-space: nowrap;
  transition: all .15s;
  font-family: inherit;
}
.sg-tab-add:hover { background: #e8f5e9; border-color: #2e7d32; }

/* ── Toolbar ────────────────────────────────────────────────────── */
.sg-toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 8px 16px;
  background: #fff;
  border-bottom: 1px solid #e0ece0;
  flex-shrink: 0;
  flex-wrap: wrap;
}
.sg-toolbar-left { display: flex; align-items: center; gap: 10px; flex: 1; min-width: 0; }
.sg-proj-name { font-weight: 700; font-size: 14px; color: #1d5c2e; }
.sg-badge {
  padding: 2px 10px;
  background: #e8f5e9;
  color: #2e7d32;
  border-radius: 12px;
  font-size: 11px;
  font-weight: 600;
  white-space: nowrap;
}
.sg-zoom-wrap { display: flex; align-items: center; gap: 8px; font-size: 11.5px; color: #555; }
.sg-zoom-wrap input[type=range] {
  -webkit-appearance: none;
  width: 100px;
  height: 4px;
  border-radius: 2px;
  background: linear-gradient(to right, #1d5c2e var(--pct,40%), #ddd var(--pct,40%));
  outline: none;
  cursor: pointer;
}
.sg-zoom-wrap input[type=range]::-webkit-slider-thumb {
  -webkit-appearance: none;
  width: 14px; height: 14px;
  border-radius: 50%;
  background: #1d5c2e;
  border: 2px solid #fff;
  box-shadow: 0 1px 4px rgba(0,0,0,.25);
  cursor: pointer;
}
.sg-zoom-label { font-weight: 600; color: #1d5c2e; min-width: 28px; font-size: 11px; }


/* ── THE scroll wrapper ──────────────────────────────────────────
   overflow-x:scroll = horizontal scrollbar ALWAYS visible.
   overflow-y:auto   = vertical only when needed.
────────────────────────────────────────────────────────────────── */
.sg-scroll {
  flex: 1;
  min-height: 0;
  overflow-x: scroll;
  overflow-y: auto;
  position: relative;
  -webkit-overflow-scrolling: touch;
}
.sg-scroll::-webkit-scrollbar { width: 8px; height: 11px; }
.sg-scroll::-webkit-scrollbar-track { background: #eef4ee; border-radius: 0; }
.sg-scroll::-webkit-scrollbar-thumb {
  background: #7db87d;
  border-radius: 6px;
  border: 2px solid #eef4ee;
  min-width: 50px;
}
.sg-scroll::-webkit-scrollbar-thumb:hover { background: #1d5c2e; }
.sg-scroll::-webkit-scrollbar-corner { background: #eef4ee; }
.sg-scroll { scrollbar-width: thin; scrollbar-color: #7db87d #eef4ee; }

/* ── Table ───────────────────────────────────────────────────────
   CRITICAL: border-collapse:collapse BREAKS position:sticky.
   Use border-collapse:separate + careful border management.
────────────────────────────────────────────────────────────────── */
.sg-tbl {
  border-collapse: separate;
  border-spacing: 0;
  table-layout: fixed;
  white-space: nowrap;
  min-width: max-content;
  font-size: 11.5px;
}

/* All cells: right + bottom border only */
.sg-tbl th, .sg-tbl td {
  border-right: 1px solid #c8ddc8;
  border-bottom: 1px solid #c8ddc8;
  height: 26px;
  line-height: 26px;
  vertical-align: middle;
  overflow: hidden;
  padding: 0 6px;
  box-sizing: border-box;
}

/* Left border only on the first frozen column */
.sg-tbl .fc-first {
  border-left: 1px solid #c8ddc8;
}

/* ── Sticky frozen left columns ──────────────────────────────────
   z-index ladder: body fc=3, thead fc=6, thead actions=7
────────────────────────────────────────────────────────────────── */
.sg-tbl .fc { position: sticky; z-index: 3; background: inherit; }
.sg-tbl thead .fc { z-index: 6; }

/* Drop shadow on the actions column to visually separate frozen from scroll */
.sg-tbl .fc-shadow {
  box-shadow: 3px 0 8px -2px rgba(0,0,0,.15);
  z-index: 4;
}
.sg-tbl thead .fc-shadow { z-index: 7; }

/* ── Header rows ─────────────────────────────────────────────────
   3 rows: [col labels + month groups] [month per-day] [day numbers]
   All 3 are sticky-top at different offsets.
────────────────────────────────────────────────────────────────── */

/* Row H1 — column labels (left) + month group (right) */
.sg-h1 th {
  background: #1d5c2e !important;
  color: #fff !important;
  font-weight: 700;
  font-size: 10.5px;
  text-transform: uppercase;
  letter-spacing: .05em;
  text-align: center;
  white-space: normal;
  line-height: 1.3;
  padding: 4px 6px;
  height: 40px;
  position: sticky;
  top: 0;
  z-index: 5;
  border-bottom: 1px solid #145220;
}
.sg-h1 th.fc { z-index: 7; }

/* Row H2 — month name per-column (date cols only) */
.sg-h2 th {
  background: #2e7d32 !important;
  color: rgba(255,255,255,.9) !important;
  font-size: 8.5px;
  font-weight: 700;
  text-align: center;
  text-transform: uppercase;
  letter-spacing: .04em;
  padding: 0 1px;
  height: 12px;
  line-height: 12px;
  position: sticky;
  top: 40px;
  z-index: 5;
  border-bottom: 1px solid #1a5c1e;
}
.sg-h2 th.fc {
  z-index: 7;
  background: #1d5c2e !important;
  height: 0; line-height: 0; padding: 0; border: none; overflow: hidden;
}

/* Row H3 — day numbers */
.sg-h3 th {
  background: #388e3c !important;
  color: rgba(255,255,255,.75) !important;
  font-size: 8px;
  font-weight: 400;
  text-align: center;
  padding: 0;
  height: 10px;
  line-height: 10px;
  position: sticky;
  top: 52px;
  z-index: 5;
  border-bottom: 2px solid #1a5228;
}
.sg-h3 th.fc {
  z-index: 7;
  background: #1d5c2e !important;
  height: 0; line-height: 0; padding: 0; border-bottom: 2px solid #1a5228; overflow: hidden;
}

/* Top border on very first row of thead */
.sg-h1 th { border-top: 1px solid #145220; }

/* ── Body row types ──────────────────────────────────────────── */
tr.sg-sec > td {
  background: #1d5c2e !important;
  color: #fff !important;
  font-weight: 700;
  font-size: 11.5px;
  letter-spacing: .02em;
}
tr.sg-sub > td {
  background: #388e3c !important;
  color: #fff !important;
  font-weight: 600;
  font-size: 11px;
}
tr.sg-task > td { background: #fff; color: #1a3c1a; }
tr.sg-task:nth-child(even) > td { background: #f5fbf5; }
tr.sg-task:hover > td { background: #e8f5e9 !important; cursor: default; transition: background .08s; }
tr.sg-yel > td { background: #fffde7 !important; color: #333 !important; }
tr.sg-yel:hover > td { background: #fff9c4 !important; }

/* Task name */
.sg-name { display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
tr.sg-task .sg-name { color: #1a5c20; font-weight: 500; }
tr.sg-yel .sg-name { color: #333; font-weight: 500; }

/* Center cols */
.sg-cc { text-align: center !important; font-size: 10.5px !important; }

/* ── Gantt bar cells ─────────────────────────────────────────── */
.sg-bar     { background: #2e7d32 !important; padding: 4px 0 !important; }
.sg-bar-y   { background: #f9a825 !important; padding: 4px 0 !important; }
.sg-bar-sub { background: #43a047 !important; padding: 4px 0 !important; }
.sg-bar-sec { background: #1b5e20 !important; padding: 4px 0 !important; }

/* Today column */
.sg-today   { border-left: 2px solid #e53935 !important; }
.sg-today-h { border-left: 2px solid rgba(229,57,53,.6) !important; }

/* ── Actions ─────────────────────────────────────────────────── */
.sg-acts { display: flex; gap: 3px; justify-content: center; align-items: center; padding: 0 2px; }
.sg-ab {
  padding: 2px 8px;
  font-size: 10px;
  font-weight: 600;
  border: 1px solid #ccc;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  font-family: inherit;
  line-height: 1.5;
  transition: all .1s;
}
.sg-ab:hover { background: #f0f0f0; }
.sg-ab.del { color: #c62828; border-color: #ffcdd2; }
.sg-ab.del:hover { background: #ffebee; border-color: #c62828; }

/* ── Legend ──────────────────────────────────────────────────── */
.sg-legend {
  display: flex;
  gap: 18px;
  align-items: center;
  padding: 6px 16px;
  border-top: 1px solid #e0ece0;
  font-size: 11px;
  color: #666;
  flex-shrink: 0;
  flex-wrap: wrap;
  background: #f9fcf9;
}
.sg-li { display: flex; align-items: center; gap: 5px; }
.sg-sw { width: 18px; height: 10px; border-radius: 2px; border: 1px solid rgba(0,0,0,.1); flex-shrink: 0; }
.sg-hint-scroll {
  flex: 1;
  text-align: right;
  font-size: 10.5px;
  color: #bbb;
  font-style: italic;
}

/* ── Modals ──────────────────────────────────────────────────── */
.sg-overlay {
  position: fixed; inset: 0;
  background: rgba(0,0,0,.5);
  z-index: 9999;
  display: flex; align-items: center; justify-content: center;
  animation: sgFi .15s ease;
}
@keyframes sgFi { from { opacity: 0 } to { opacity: 1 } }

.sg-btn {
  height: 40px;
  padding: 0 18px;
  border-radius: 14px;
  border: 1px solid rgba(160, 190, 255, 0.7);
  background: transparent;
  color: #cfe0ff;
  font-size: 14px;
  font-weight: 700;
  cursor: pointer;
  font-family: inherit;
  transition: all .18s ease;
  white-space: nowrap;
}

.sg-btn:hover {
  background: rgba(255, 255, 255, 0.04);
  border-color: #8fb2ff;
  color: #ffffff;
}

.sg-btn.primary {
  background: linear-gradient(180deg, #2f5fbf 0%, #2a4f9b 100%);
  color: #ffffff;
  border-color: #4b78d1;
  box-shadow: inset 0 1px 0 rgba(255,255,255,0.08);
}

.sg-btn.primary:hover {
  background: linear-gradient(180deg, #3b6bcb 0%, #3159aa 100%);
  border-color: #6d94e6;
}

.sg-modal {
  background: linear-gradient(180deg, #0d2238 0%, #0a1c2f 100%);
  border: 1px solid rgba(135, 168, 214, 0.18);
  border-radius: 20px;
  padding: 0;
  width: min(540px, 95vw);
  max-height: 90vh;
  overflow-y: auto;
  box-shadow: 0 24px 80px rgba(0,0,0,.45);
  animation: sgSu .2s ease;
}

.sg-modal h3 {
  margin: 0;
  padding: 20px 22px 0;
  font-size: 16px;
  font-weight: 700;
  color: #3f8cff;
}

.sg-modal p {
  margin: 6px 0 0;
  padding: 0 22px 18px;
  font-size: 12px;
  color: #92a7bf;
  border-bottom: 1px solid rgba(255,255,255,0.08);
}

.sg-fg {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 18px 18px;
  padding: 18px 22px 0;
}

.sg-f {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.sg-f.full {
  grid-column: 1 / -1;
}

.sg-f label {
  font-size: 11px;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: .06em;
  color: #aebed0;
}

.sg-f input,
.sg-f select {
  height: 44px;
  padding: 0 14px;
  border: 1px solid rgba(140, 165, 195, 0.28);
  border-radius: 14px;
  font-size: 14px;
  font-family: inherit;
  outline: none;
  background: rgba(255,255,255,0.06);
  color: #e8f0fb;
  transition: border-color .18s, box-shadow .18s, background .18s;
}

.sg-f input::placeholder,
.sg-f select::placeholder {
  color: #8ea1b8;
}

.sg-f input:focus,
.sg-f select:focus {
  border-color: #4d7dff;
  box-shadow: 0 0 0 3px rgba(77,125,255,.18);
  background: rgba(255,255,255,0.08);
}

.sg-f input[readonly] {
  background: rgba(255,255,255,0.04);
  color: #91a4ba;
}

.sg-f .hint {
  font-size: 11px;
  color: #7f93ab;
  margin-top: 2px;
}

.sg-err {
  display: none;
  margin: 14px 22px 0;
  padding: 10px 12px;
  background: rgba(220, 53, 69, 0.14);
  color: #ff808e;
  border: 1px solid rgba(255, 128, 142, 0.28);
  border-radius: 12px;
  font-size: 12px;
}

.sg-ma {
  display: flex;
  gap: 10px;
  justify-content: flex-end;
  margin-top: 20px;
  padding: 16px 22px 22px;
  border-top: 1px solid rgba(255,255,255,0.08);
  background: rgba(255,255,255,0.02);
}


/* ── Project blocks ───────────────────────────────────────────── */
.schedule-project-block {
  border: 1px solid #dce9dc;
  border-radius: 18px;
  background: #fff;
  overflow: hidden;
  box-shadow: 0 6px 18px rgba(29,92,46,.06);
}
.schedule-project-block__head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 14px 16px;
  background: linear-gradient(180deg, #f6fbf6 0%, #eef7ee 100%);
  border-bottom: 1px solid #dce9dc;
}
.schedule-project-block__title {
  font-size: 15px;
  font-weight: 800;
  color: #143d1c;
}
.schedule-project-block__sub {
  margin-top: 2px;
  font-size: 11px;
  color: #618061;
}
.schedule-project-block__tag {
  padding: 5px 10px;
  border-radius: 999px;
  background: #e8f5e9;
  color: #1d5c2e;
  font-size: 11px;
  font-weight: 700;
  white-space: nowrap;
}
.schedule-project-block__body {
  display: flex;
  flex-direction: column;
  min-height: 0;
  background: #fff;
}
.schedule-project-tablewrap {
  min-width: max-content;
}
.schedule-project-add-wrap {
  display: flex;
  justify-content: center;
  padding: 6px 0 18px;
}
.schedule-project-add {
  font-size: 12px;
}


/* ── Empty state ─────────────────────────────────────────────── */
.sg-empty {
  display: flex; flex-direction: column; align-items: center;
  justify-content: center; gap: 10px;
  padding: 60px 20px; color: #bbb; font-size: 13px;
}
.sg-empty-icon { font-size: 40px; }
.sg-empty-sub  { font-size: 11px; color: #ddd; }

.schedule-single-shell {
  gap: 0;
  overflow: hidden;
}
.schedule-project-selector {
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:12px;
  padding:12px 16px;
  border-bottom:1px solid #e3ece3;
  background:linear-gradient(180deg,#f8fbf8,#f3f8f3);
  flex-wrap:wrap;
}
.schedule-project-selector__left,
.schedule-project-selector__right {
  display:flex;
  align-items:center;
  gap:10px;
  flex-wrap:wrap;
}
.schedule-project-selector__label {
  font-size:12px;
  font-weight:700;
  color:#1d5c2e;
}
.schedule-project-selector__search,
.schedule-project-selector__select {
  height:36px;
  border:1px solid #c9dcc9;
  border-radius:10px;
  background:#fff;
  padding:0 12px;
  font:600 12px/1.2 Inter, system-ui, sans-serif;
  color:#1b3e1f;
}
.schedule-project-selector__search {
  min-width:200px;
}
.schedule-project-selector__select {
  min-width:220px;
}
.schedule-project-block__body {
  display:flex;
  flex-direction:column;
  min-height:0;
  background:#fff;
  overflow:hidden;
}
.schedule-project-tablewrap {
  min-width:max-content;
}
  `;
  document.head.appendChild(s);
})();

// ─── Date helpers ─────────────────────────────────────────────────
const _p2 = (n) => String(n).padStart(2, "0");
const _sod = (d) => {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
};
const _iso = (d) =>
  `${d.getFullYear()}-${_p2(d.getMonth() + 1)}-${_p2(d.getDate())}`;
const _add = (d, n) => {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
};

function _sgPD(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return isNaN(v) ? null : _sod(v);
  const s = String(v).trim();
  if (!s) return null;
  const n = Number(s);
  if (!isNaN(n) && n > 20000 && n < 90000)
    return _sod(new Date(Date.UTC(1899, 11, 30) + n * 86400000));
  const m1 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m1) return _sod(new Date(+m1[1], +m1[2] - 1, +m1[3]));
  const MO = {
    jan: 0,
    feb: 1,
    mar: 2,
    apr: 3,
    may: 4,
    jun: 5,
    jul: 6,
    aug: 7,
    sep: 8,
    oct: 9,
    nov: 10,
    dec: 11,
  };
  const m2 = s.match(/^(\d{1,2})[-\/\s]([A-Za-z]{3,})[-\/\s](\d{2,4})$/);
  if (m2) {
    const mo = MO[m2[2].toLowerCase().slice(0, 3)];
    let y = +m2[3];
    if (y < 100) y += y < 50 ? 2000 : 1900;
    if (mo != null) return _sod(new Date(y, mo, +m2[1]));
  }
  const d = new Date(s);
  return isNaN(d) ? null : _sod(d);
}
function _sgFD(v) {
  const d = _sgPD(v);
  if (!d) return "";
  const M = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  return `${d.getDate()}-${M[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
}
function _sgDays(rows) {
  let mn = null,
    mx = null;
  rows.forEach((r) => {
    const s = _sgPD(r.startDate),
      e = _sgPD(r.endDate);
    if (s && (!mn || s < mn)) mn = new Date(s);
    if (e && (!mx || e > mx)) mx = new Date(e);
  });
  if (!mn || !mx) return [];
  const days = [];
  let c = new Date(mn);
  while (c <= mx) {
    days.push(new Date(c));
    c = _add(c, 1);
  }
  return days;
}
function _sgMG(days) {
  const M = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  const g = [];
  let cur = null;
  days.forEach((d) => {
    const k = `${M[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
    if (!cur || cur.k !== k) {
      cur = { k, label: k, count: 0 };
      g.push(cur);
    }
    cur.count++;
  });
  return g;
}
function _sgRC(row) {
  if (row.__isSection) return "sg-sec";
  if (row.__isTitle) return "sg-sub";
  const nm = String(row.taskName || "").trim();
  const hd = !!(_sgPD(row.startDate) && _sgPD(row.endDate));
  if (!hd)
    return nm === nm.toUpperCase() && nm.replace(/\s/g, "").length > 1
      ? "sg-sec"
      : "sg-sub";
  if (row.__highlight === "yellow") return "sg-task sg-yel";
  return "sg-task";
}

// ─── Overlay helper ───────────────────────────────────────────────
function _sgOv() {
  const bg = document.createElement("div");
  bg.className = "sg-overlay";
  bg.onclick = (e) => {
    if (e.target === bg) bg.remove();
  };
  return bg;
}

// ─── Add Project modal ────────────────────────────────────────────

function _sgAddProject() {
  const bg = _sgOv();
  bg.innerHTML = `<div class="sg-modal" style="width:min(420px,95vw)">
    <h3>➕ New Project</h3>
    <p>Creates a new project sheet with the same schedule structure, ready for tasks.</p>
    <div class="sg-fg">
      <div class="sg-f full">
        <label>Project Name *</label>
        <input id="_pn" type="text" placeholder="e.g. Ramhari HPP, Upper Seti…" autocomplete="off"/>
        <span class="hint">Sheet name: &lt;name&gt;_timeline (auto-generated)</span>
      </div>
    </div>
    <div class="sg-err" id="_pe"></div>
    <div class="sg-ma">
      <button class="sg-btn" id="_pc">Cancel</button>
      <button class="sg-btn primary" id="_ps">Create Project</button>
    </div>
  </div>`;
  document.body.appendChild(bg);
  const ni = bg.querySelector("#_pn");
  bg.querySelector("#_pc").onclick = () => bg.remove();
  ni.focus();
  bg.querySelector("#_ps").onclick = async () => {
    const name = ni.value.trim();
    const err = bg.querySelector("#_pe");
    if (!name) {
      err.textContent = "Project name is required.";
      err.style.display = "block";
      ni.focus();
      return;
    }

    const sheet =
      name
        .replace(/[^\w\s-]/g, "")
        .trim()
        .replace(/\s+/g, "_") + "_timeline";
    const dup = (store.scheduleProjects || []).some(
      (p) => p.sheet.toLowerCase() === sheet.toLowerCase(),
    );
    if (dup) {
      err.textContent = `"${name}" already exists.`;
      err.style.display = "block";
      return;
    }

    const btn = bg.querySelector("#_ps");
    btn.disabled = true;
    btn.textContent = "Creating…";

    try {
      await apiCreateRow(sheet, "schedule", {
        "Task Name / Milestone": "",
        "Duration (days)": "",
        "Start Date": "",
        "Finish Date": "",
        "Notes / Section": "",
      });

      store.scheduleProjectCache = {};
      await loadScheduleProjects();
      store.activeScheduleSheet = sheet;
      store.scheduleProjectSearch = name;
      await loadCurrentViewData({ resetPage: true });
      toast(
        "success",
        `Project "${name}" created`,
        "New schedule sheet is ready ✓",
      );
      bg.remove();
    } catch (e) {
      err.textContent = e.message || "Failed to create project.";
      err.style.display = "block";
      btn.disabled = false;
      btn.textContent = "Create Project";
    }
  };
}

// ─── Task CRUD modal ──────────────────────────────────────────────
function _sgTask(mode, row, sheet) {
  const bg = _sgOv();
  const v = row || {};
  const fi = (x) => {
    const d = _sgPD(x);
    return d ? _iso(d) : "";
  };
  bg.innerHTML = `<div class="sg-modal">
    <h3>${mode === "add" ? "➕ Add Task" : "✏️ Edit Task"}</h3>
    <p>Saved directly to the Excel workbook.</p>
    <div class="sg-fg">
      <div class="sg-f full">
        <label>Task Name / Milestone *</label>
        <input id="_tn" type="text" value="${esc(v.taskName || "")}" placeholder="Task name or milestone…"/>
      </div>
      <div class="sg-f">
        <label>Start Date</label>
        <input id="_ts" type="date" value="${fi(v.startDate)}"/>
      </div>
      <div class="sg-f">
        <label>Finish Date</label>
        <input id="_te" type="date" value="${fi(v.endDate)}"/>
      </div>
      <div class="sg-f">
        <label>Duration (days)</label>
       <input id="_td" type="number" value="${esc(String(v.duration || ""))}" placeholder="Enter days or use dates"/>
       <span class="hint">Auto-calculated from dates, or enter manually</span>
      </div>
      <div class="sg-f">
        <label>Notes / Section</label>
        <input id="_tnotes" type="text" value="${esc(v.remarks || "")}" placeholder="Optional…"/>
      </div>
    </div>
    <div class="sg-err" id="_te2"></div>
    <div class="sg-ma">
      <button class="sg-btn" id="_tc">Cancel</button>
      <button class="sg-btn primary" id="_tv">${mode === "add" ? "Create Task" : "Save Changes"}</button>
    </div>
  </div>`;
  document.body.appendChild(bg);
  const ne = bg.querySelector("#_tn"),
    se = bg.querySelector("#_ts"),
    ee = bg.querySelector("#_te"),
    de = bg.querySelector("#_td"),
    err = bg.querySelector("#_te2");
  const rc = (source = "") => {
    const s = _sgPD(se.value);
    const e = _sgPD(ee.value);
    const d = Number(String(de.value || "").trim());

    if (source !== "duration" && s && e && e >= s) {
      de.value = Math.round((e - s) / 86400000) + 1;
      return;
    }

    if (source === "duration" && s && Number.isFinite(d) && d > 0) {
      const finish = new Date(s);
      finish.setDate(finish.getDate() + Math.max(0, Math.round(d) - 1));
      ee.value = _iso(finish);
    }
  };

  se.onchange = () => rc("start");
  ee.onchange = () => rc("end");
  de.oninput = () => rc("duration");
  rc();
  bg.querySelector("#_tc").onclick = () => bg.remove();
  ne.focus();
  bg.querySelector("#_tv").onclick = async () => {
    const name = ne.value.trim();
    if (!name) {
      err.textContent = "Task name is required.";
      err.style.display = "block";
      ne.focus();
      return;
    }
    err.style.display = "none";
    const payload = {
      "Task Name": name,
      "Start Date": se.value || "",
      "End Date": ee.value || "",
      Duration: de.value || "",
      Remarks: bg.querySelector("#_tnotes").value || "",
    };
    const btn = bg.querySelector("#_tv");
    btn.disabled = true;
    btn.textContent = "Saving…";
    try {
      if (mode === "add") {
        await apiCreateRow(sheet, "schedule", payload);
        toast("success", "Task created", "Saved to Excel");
      } else {
        await apiUpdateRow(sheet, v.id, payload);
        toast("success", "Task updated", "Saved to Excel");
      }

      await refreshScheduleState(sheet, { resetPage: false });

      bg.remove();
      renderScheduleSheet();
    } catch (e) {
      err.textContent = e.message || "Save failed.";
      err.style.display = "block";
      btn.disabled = false;
      btn.textContent = mode === "add" ? "Create Task" : "Save Changes";
    }
  };
}

async function deleteScheduleTaskRow(sheet, row) {
  const ok = window.confirm(
    `Are you sure you want to delete "${row?.["Task Name"] || row?.taskName || "this row"}"?`,
  );
  if (!ok) return;

  try {
    await apiDeleteRow(sheet, row.id);

    delete store.scheduleProjectCache?.[sheet];
    delete store.scheduleGridCache?.[sheet];

    const refreshed = await loadScheduleProjectRows(sheet, { force: true });

    store.scheduleProjectRows = [
      ...(store.scheduleProjectRows || []).filter((p) => p.sheet !== sheet),
      refreshed,
    ];

    if (store.activeScheduleSheet === sheet) {
      store.data = Array.isArray(refreshed.rows) ? refreshed.rows.slice() : [];
      store.allSchedule = store.data.slice();
      views.schedule.columns = inferColumnsFromData(store.data, "schedule");
      views.gantt.columns = inferColumnsFromData(store.data, "gantt");
    }

    toast("success", "Task deleted", "Removed from Excel");
    renderScheduleSheet();
  } catch (err) {
    toast("error", "Delete failed", err.message || "");
  }
}

async function apiRenameScheduleProject(sheet, name) {
  const res = await apiFetch(`${API}/excel/schedule-project`, {
    method: "PUT",
    body: JSON.stringify({ sheet, name }),
  });
  const json = await res.json().catch(() => ({}));
  if (!json.ok) throw new Error(json.error || "Failed to rename project");
  return json;
}

async function apiDeleteScheduleProject(sheet) {
  const res = await apiFetch(
    `${API}/excel/schedule-project?sheet=${encodeURIComponent(sheet)}`,
    {
      method: "DELETE",
    },
  );
  const json = await res.json().catch(() => ({}));
  if (!json.ok) throw new Error(json.error || "Failed to delete project");
  return json;
}

// ─── Core table renderer ──────────────────────────────────────────
function _sgDraw(wrap, rows, dayW, activeSheet) {
  dayW = Math.max(14, Math.min(52, dayW || 22));
  wrap.innerHTML = "";

  // Fixed column definitions
  const FC = [
    {
      label: "Task Name /\nMilestone",
      w: 230,
      cls: "",
      get: (r) => r.taskName || "",
    },
    {
      label: "Duration\n(days)",
      w: 64,
      cls: "sg-cc",
      get: (r) => r.duration || "",
    },
    {
      label: "Start\nDate",
      w: 82,
      cls: "sg-cc",
      get: (r) => _sgFD(r.startDate),
    },
    {
      label: "Finish\nDate",
      w: 82,
      cls: "sg-cc",
      get: (r) => _sgFD(r.endDate),
    },
    { label: "Notes /\nSection", w: 90, cls: "", get: (r) => r.remarks || "" },
  ];
  const AW = 82; // actions col width

  // Compute cumulative left for sticky positioning
  const lefts = [];
  let cum = 0;
  FC.forEach((f) => {
    lefts.push(cum);
    cum += f.w;
  });
  const actsLeft = cum; // left of actions col
  const frozenW = cum + AW; // total frozen width (helps with gantt scroll offset)

  const days = _sgDays(rows);
  const today = _iso(_sod(new Date()));
  const MNAMES = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];

  // ── Build table DOM ───────────────────────────────────────────
  const tbl = document.createElement("table");
  tbl.className = "sg-tbl";
  const TH = document.createElement("thead");
  const TB = document.createElement("tbody");
  tbl.append(TH, TB);

  // Helper: make a frozen-col th/td
  function mkFth(i, extra) {
    const th = document.createElement("th");
    th.className = `fc${i === 0 ? " fc-first" : ""}${i === FC.length - 1 ? " fc-last" : ""}`;
    th.classList.add("fc");
    th.style.cssText = `width:${FC[i].w}px;min-width:${FC[i].w}px;left:${lefts[i]}px;${extra || ""}`;
    return th;
  }
  function mkFtd(i) {
    const td = document.createElement("td");
    td.className = `fc${i === 0 ? " fc-first" : ""}`;
    td.classList.add("fc");
    td.style.cssText = `width:${FC[i].w}px;min-width:${FC[i].w}px;left:${lefts[i]}px;`;
    return td;
  }
  function mkActTh() {
    const th = document.createElement("th");
    th.className = "fc fc-shadow";
    th.style.cssText = `width:${AW}px;min-width:${AW}px;left:${actsLeft}px;`;
    return th;
  }
  function mkActTd() {
    const td = document.createElement("td");
    td.className = "fc fc-shadow";
    td.style.cssText = `width:${AW}px;min-width:${AW}px;left:${actsLeft}px;background:inherit;`;
    return td;
  }
  function mkDayTh(day, cls) {
    const th = document.createElement("th");
    th.style.cssText = `width:${dayW}px;min-width:${dayW}px;max-width:${dayW}px;`;
    if (_iso(day) === today) th.classList.add(cls || "sg-today-h");
    return th;
  }
  function mkDayTd(day) {
    const td = document.createElement("td");
    td.style.cssText = `width:${dayW}px;min-width:${dayW}px;max-width:${dayW}px;padding:3px 0;`;
    if (_iso(day) === today) td.classList.add("sg-today");
    return td;
  }

  // ── HEADER ROW 1: labels + month groups ──────────────────────
  const h1 = document.createElement("tr");
  h1.className = "sg-h1";
  FC.forEach((_, i) => {
    const th = mkFth(i);
    th.style.zIndex = "8";
    th.innerHTML = FC[i].label.replace(/\n/, "<br>");
    h1.appendChild(th);
  });
  const ah1 = mkActTh();
  ah1.style.zIndex = "8";
  ah1.textContent = "Actions";
  h1.appendChild(ah1);
  _sgMG(days).forEach(({ label, count }) => {
    const th = document.createElement("th");
    th.colSpan = count;
    th.textContent = label;
    h1.appendChild(th);
  });
  TH.appendChild(h1);

  // ── HEADER ROW 2: month name per day column ───────────────────
  const h2 = document.createElement("tr");
  h2.className = "sg-h2";
  FC.forEach((_, i) => {
    const th = mkFth(i);
    h2.appendChild(th);
  });
  const ah2 = mkActTh();
  h2.appendChild(ah2);
  days.forEach((day) => {
    const th = mkDayTh(day, "sg-today-h");
    // Show month abbreviation only on 1st of month
    if (day.getDate() === 1) th.textContent = MNAMES[day.getMonth()];
    h2.appendChild(th);
  });
  TH.appendChild(h2);

  // ── HEADER ROW 3: day numbers ─────────────────────────────────
  const h3 = document.createElement("tr");
  h3.className = "sg-h3";
  FC.forEach((_, i) => {
    const th = mkFth(i);
    h3.appendChild(th);
  });
  const ah3 = mkActTh();
  h3.appendChild(ah3);
  let todayIdx = -1;
  days.forEach((day, di) => {
    const th = mkDayTh(day, "sg-today-h");
    if (dayW >= 18) th.textContent = day.getDate();
    else if (day.getDate() % 5 === 0) th.textContent = day.getDate();
    if (_iso(day) === today) todayIdx = di;
    h3.appendChild(th);
  });
  TH.appendChild(h3);

  // ── BODY ROWS ─────────────────────────────────────────────────
  rows.forEach((row) => {
    const cls = _sgRC(row);
    const isHdr = cls === "sg-sec" || cls === "sg-sub";
    const isYel = cls.includes("sg-yel");
    const rS = _sgPD(row.startDate);
    const rE = _sgPD(row.endDate);
    const tr = document.createElement("tr");
    tr.className = cls;

    if (isHdr) {
      // Section/sub: name spans all fixed cols + actions, no individual bars
      const td = document.createElement("td");
      td.className = "fc fc-first fc-shadow";
      td.colSpan = FC.length + 1; // covers all FC cols + actions
      td.style.cssText = `left:0;width:${frozenW}px;min-width:${frozenW}px;`;
      td.textContent = row.taskName || "";
      tr.appendChild(td);
      days.forEach((day) => {
        const td = mkDayTd(day);
        if (rS && rE && day >= rS && day <= rE)
          td.classList.add(cls === "sg-sec" ? "sg-bar-sec" : "sg-bar-sub");
        tr.appendChild(td);
      });
    } else {
      // Normal task: all cols individually
      FC.forEach((_, i) => {
        const td = mkFtd(i);
        if (FC[i].cls) td.classList.add(...FC[i].cls.split(" "));
        const val = FC[i].get(row);
        if (i === 0) {
          const sp = document.createElement("span");
          sp.className = "sg-name";
          sp.title = val;
          sp.textContent = val;
          td.appendChild(sp);
        } else {
          td.textContent = val;
        }
        tr.appendChild(td);
      });
      // Actions
      const atd = mkActTd();
      const aw = document.createElement("div");
      aw.className = "sg-acts";
      const eb = document.createElement("button");
      eb.className = "sg-ab";
      eb.textContent = "Edit";
      eb.onclick = (e) => {
        e.stopPropagation();
        _sgTask("edit", row, activeSheet);
      };
      const db = document.createElement("button");
      db.className = "sg-ab del";
      db.textContent = "Del";
      db.onclick = (e) => {
        e.stopPropagation();
        deleteScheduleTaskRow(activeSheet, row);
      };
      aw.append(eb, db);
      atd.appendChild(aw);
      tr.appendChild(atd);
      // Date/gantt cells
      days.forEach((day) => {
        const td = mkDayTd(day);
        if (rS && rE && day >= rS && day <= rE)
          td.classList.add(isYel ? "sg-bar-y" : "sg-bar");
        tr.appendChild(td);
      });
    }
    TB.appendChild(tr);
  });

  wrap.appendChild(tbl);

  // Auto-scroll so today is visible in the gantt area
  if (todayIdx >= 0) {
    requestAnimationFrame(() => {
      const scroller = wrap.closest(".sg-scroll");
      if (!scroller) return;
      const ganttStart = frozenW + todayIdx * dayW;
      const viewW = scroller.clientWidth;
      // Position today at ~35% from left edge of visible gantt
      scroller.scrollLeft = Math.max(0, ganttStart - frozenW - viewW * 0.35);
    });
  }
}

// ─── Main entry point ─────────────────────────────────────────────

function renderScheduleSheet() {
  const mount = el("scheduleSheet");
  const header = el("scheduleHeader");
  if (!mount) return;
  if (header) header.style.display = "none";

  const projects = Array.isArray(store.scheduleProjects)
    ? store.scheduleProjects
    : [];
  const activeProject = getActiveScheduleProject();
  const projectData =
    Array.isArray(store.scheduleProjectRows) && store.scheduleProjectRows[0]
      ? store.scheduleProjectRows[0]
      : activeProject
        ? { sheet: activeProject.sheet, title: activeProject.title, rows: [] }
        : null;

  mount.innerHTML = "";
  mount.style.cssText =
    ""; /* Let body.is-schedule #scheduleSheet CSS class control layout */

  const shell = document.createElement("section");
  shell.className = "schedule-project-block schedule-single-shell";

  const topbar = document.createElement("div");
  topbar.className = "schedule-project-selector";

  const left = document.createElement("div");
  left.className = "schedule-project-selector__left";

  const label = document.createElement("div");
  label.className = "schedule-project-selector__label";
  label.textContent = "Project";

  const search = document.createElement("input");
  search.className = "schedule-project-selector__search";
  search.type = "text";
  search.placeholder = "Filter projects...";
  search.value = store.scheduleProjectSearch || "";

  const select = document.createElement("select");
  select.className = "schedule-project-selector__select";

  const filteredProjects = projects.filter((project) => {
    const q = String(store.scheduleProjectSearch || "")
      .trim()
      .toLowerCase();
    if (!q) return true;
    return [project.title, project.sheet].some((v) =>
      String(v || "")
        .toLowerCase()
        .includes(q),
    );
  });

  const optionList = filteredProjects.length ? filteredProjects : projects;
  select.innerHTML = optionList
    .map((project) => {
      const selected =
        project.sheet === store.activeScheduleSheet ? "selected" : "";
      return `<option value="${escapeAttr(project.sheet)}" ${selected}>${esc(project.title || project.sheet)}</option>`;
    })
    .join("");

  if (!optionList.length) {
    select.innerHTML = `<option value="">No projects yet</option>`;
    select.disabled = true;
  }

  search.oninput = () => {
    store.scheduleProjectSearch = search.value || "";
    renderScheduleSheet();
  };

  select.onchange = async () => {
    store.activeScheduleSheet = select.value;
    saveActiveScheduleSheet(store.activeScheduleSheet);
    delete store.scheduleProjectCache?.[store.activeScheduleSheet];
    await loadCurrentViewData({ resetPage: true });
  };

  left.append(label, search, select);

  const right = document.createElement("div");
  right.className = "schedule-project-selector__right";

  let dayW = window.__sgDayW || 22;

  const zoomWrap = document.createElement("div");
  zoomWrap.className = "sg-zoom-wrap";
  const zl = document.createElement("span");
  zl.textContent = "Zoom:";
  const zs = document.createElement("input");
  zs.type = "range";
  zs.min = 14;
  zs.max = 52;
  zs.value = dayW;
  const zv = document.createElement("span");
  zv.className = "sg-zoom-label";
  zv.textContent = `${dayW}px`;
  const syncZoomTrack = () => {
    const pct = (((dayW - 14) / (52 - 14)) * 100).toFixed(1);
    zs.style.setProperty("--pct", pct + "%");
  };
  syncZoomTrack();
  zoomWrap.append(zl, zs, zv);

  const addProjectBtn = document.createElement("button");
  addProjectBtn.className = "sg-btn";
  addProjectBtn.textContent = "+ New Project";
  addProjectBtn.onclick = () => _sgAddProject();

  const addTaskBtn = document.createElement("button");
  addTaskBtn.className = "sg-btn primary";
  addTaskBtn.textContent = "+ Add Task";
  addTaskBtn.disabled = !projectData?.sheet;
  addTaskBtn.onclick = () => {
    if (projectData?.sheet) _sgTask("add", null, projectData.sheet);
  };

  right.append(zoomWrap, addProjectBtn, addTaskBtn);
  topbar.append(left, right);

  shell.appendChild(topbar);

  if (!projectData?.sheet) {
    const empty = document.createElement("div");
    empty.className = "sg-empty";
    empty.innerHTML = `
      <div class="sg-empty-icon">📋</div>
      <div>No schedule projects found</div>
      <div class="sg-empty-sub">Create your first project to start the schedule.</div>
    `;
    shell.appendChild(empty);
    mount.appendChild(shell);
    const rc = el("rowCount");
    if (rc) rc.textContent = "0 rows";
    return;
  }

  const head = document.createElement("div");
  head.className = "schedule-project-block__head";
  head.innerHTML = `
    <div>
      <div class="schedule-project-block__title">${esc(projectData.title || projectData.sheet)}</div>
      <div class="schedule-project-block__sub">Sheet: ${esc(projectData.sheet)}</div>
    </div>
    <div class="schedule-project-block__tag">${(projectData.rows || []).length} tasks</div>
  `;

  const body = document.createElement("div");
  body.className = "schedule-project-block__body";

  const scrollDiv = document.createElement("div");
  scrollDiv.className = "sg-scroll";

  const tableWrap = document.createElement("div");
  tableWrap.className = "schedule-project-tablewrap";
  scrollDiv.appendChild(tableWrap);

  if (!projectData.rows || !projectData.rows.length) {
    tableWrap.innerHTML = `
      <div class="sg-empty">
        <div class="sg-empty-icon">📋</div>
        <div>No tasks found for <strong>${esc(projectData.title || projectData.sheet)}</strong></div>
        <div class="sg-empty-sub">Use "+ Add Task" to populate this project. The Gantt updates from duration, start, and finish dates.</div>
      </div>
    `;
  } else {
    _sgDraw(tableWrap, projectData.rows, dayW, projectData.sheet);
  }

  zs.oninput = () => {
    dayW = Number(zs.value);
    window.__sgDayW = dayW;
    zv.textContent = `${dayW}px`;
    syncZoomTrack();
    if (projectData.rows && projectData.rows.length) {
      _sgDraw(tableWrap, projectData.rows, dayW, projectData.sheet);
    }
  };

  const legend = document.createElement("div");
  legend.className = "sg-legend";
  legend.innerHTML = `
    <strong style="color:#1d5c2e">Legend:</strong>
    <div class="sg-li"><div class="sg-sw" style="background:#2e7d32"></div>Task</div>
    <div class="sg-li"><div class="sg-sw" style="background:#f9a825;border-color:#e09020"></div>Highlighted</div>
    <div class="sg-li"><div class="sg-sw" style="background:#1b5e20"></div>Section header</div>
    <div class="sg-li"><div class="sg-sw" style="background:#43a047"></div>Sub-section</div>
    <div class="sg-li"><div style="width:3px;height:14px;background:#e53935;border-radius:1px;flex-shrink:0"></div>Today</div>
    <div class="sg-hint-scroll">← Scroll horizontally only when the timeline is wider than the frame →</div>
  `;

  body.append(scrollDiv, legend);
  shell.append(head, body);
  mount.appendChild(shell);

  const rc = el("rowCount");
  if (rc) {
    rc.textContent = `${(projectData.rows || []).length} rows`;
  }
}

// OVERVIEW DASHBOARD
// ================================================================

function render() {
  setActiveNav(store.view);
  applyViewLayout();

  if (store.view === "overview") {
    renderOverviewDashboard();
    return;
  }

  renderTableView();
}

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

  if (chartInstances.engagements) {
    chartInstances.engagements.destroy();
    chartInstances.engagements = null;
  }

  const projects = Array.isArray(store.scheduleProjectRows)
    ? store.scheduleProjectRows
    : [];

  if (!projects.length) {
    canvas.parentElement.innerHTML =
      '<div class="chart-empty">No schedule project data available</div>';
    return;
  }

  const items = projects.map((project) => {
    const rows = Array.isArray(project.rows) ? project.rows : [];
    const taskCount = rows.filter((row) => {
      if (!row) return false;
      if (row.__isSection || row.__isTitle) return false;
      return String(row.taskName || "").trim() !== "";
    }).length;

    return {
      label: project.title || prettifyProjectName(project.sheet),
      value: taskCount,
    };
  });

  items.sort((a, b) => b.value - a.value);

  chartInstances.engagements = new Chart(canvas, {
    type: "bar",
    data: {
      labels: items.map((x) => x.label),
      datasets: [
        {
          label: "Schedule Tasks",
          data: items.map((x) => x.value),
          backgroundColor: "#4f8ef7",
          borderRadius: 8,
          barPercentage: 0.72,
          categoryPercentage: 0.82,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { precision: 0 },
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

  const addBtn = document.createElement("button");
  addBtn.id = "btnAddRow";
  addBtn.className = "btn primary btn-sm";
  addBtn.textContent = "Add Row";
  addBtn.onclick = () => openFormModal("add", null);

  wrap.appendChild(backBtn);
  wrap.appendChild(ganttBtn);
  wrap.appendChild(zoomWrap);
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
  const sheet = getSheetForView(store.view);
  const showActions = store.view !== "overview" && !!sheet;

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
      const isToolChild =
        store.view === "tools" &&
        String(row.__rowType ?? "")
          .trim()
          .toLowerCase() === "child";

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

          if (
            store.view === "schedule" &&
            (kLower === "duration" || kLower === "duration (days)")
          ) {
            const startV = row[findKey(row, /^start$/i) || "Start"];
            const finV = row[findKey(row, /^finish$/i) || "Finish"];
            return `<td class="mono">${esc(formatScheduleDuration(raw, startV, finV))}</td>`;
          }

          if (store.view === "schedule" && kLower === "task name") {
            return `<td class="sch-task">${scheduleNameCellHTML(row)}</td>`;
          }

          if (store.view === "tools") {
            if (kLower === "s.n" || kLower === "sn" || kLower === "s n") {
              return `<td class="mono">${isToolChild ? "" : esc(raw)}</td>`;
            }

            if (kLower === "item") {
              return `<td class="${isToolChild ? "tool-child-item" : "tool-parent-item"}">
                ${
                  isToolChild
                    ? `<span class="tool-child-label">↳ ${highlight(raw, term)}</span>`
                    : highlight(raw, term)
                }
              </td>`;
            }
          }

          if (col.type === "date") {
            return `<td class="mono">${esc(formatDateShort(raw))}</td>`;
          }

          if (col.type === "number") {
            return `<td class="mono">${esc(raw)}</td>`;
          }

          return `<td>${highlight(raw, term)}</td>`;
        })
        .join("");

      const rowId =
        store.view === "sites"
          ? row.Location || row.location || ""
          : store.view === "schedule" || store.view === "gantt"
            ? (row.id ?? "")
            : store.view === "tools"
              ? (row.__rowKey ?? row.id ?? row["S.N"] ?? "")
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

      return `<tr class="${isToolChild ? "tool-child-row" : ""}">${tds}${actions}</tr>`;
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
  // tableView display is managed by applyViewLayout — no override needed here

  localStorage.setItem("activeView", view);

  history.replaceState(
    null,
    "",
    `${location.pathname}${location.search}#${view}`,
  );

  await loadCurrentViewData();
  render();
}
const navExportBtn = document.getElementById("btnExportExcel");
if (navExportBtn) {
  navExportBtn.addEventListener("click", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    await downloadExcelExport();
  });
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
  const sheet =
    store.view === "schedule" || store.view === "gantt"
      ? store.activeScheduleSheet
      : getSheetForView(store.view);
  if (!sheet) return;

  let cols;

  if (store.view === "schedule" || store.view === "gantt") {
    cols = [
      { key: "taskName", label: "Task Name", type: "text" },
      { key: "completion", label: "% Completion", type: "number" },
      { key: "duration", label: "Duration", type: "number" },
      { key: "startDate", label: "Start Date", type: "date" },
      { key: "endDate", label: "End Date", type: "date" },
      { key: "remarks", label: "Remarks", type: "text" },
    ];
  } else {
    cols = views[store.view].columns.filter((c) => {
      const key = String(c.key || "")
        .trim()
        .toLowerCase();

      if (key === "createdat" || key === "updatedat") return false;
      if (key === "id") return false;

      if (
        store.view === "tools" &&
        (key === "s.n" || key === "sn" || key === "s n")
      ) {
        return false;
      }

      return true;
    });
  }

  if (!cols.length) {
    toast("error", "No columns found", "Check your data structure");
    return;
  }

  const modal = el("modalOverlay");
  const title = el("modalTitle");
  const formFields = el("formFields");
  const submitBtn = el("submitBtnText");

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
          : store.view === "schedule" || store.view === "gantt"
            ? store.data.find(
                (r) => String(r.id || "").trim() === String(rowId).trim(),
              )
            : store.data.find((r) => String(r.id || "") === String(rowId))) ||
        {}
      : {};

  formFields.className =
    cols.length > 4 ? "form-fields two-col" : "form-fields";

  const isEng = store.view === "engagements";
  const isSch = store.view === "schedule";
  const isTools = store.view === "tools";
  const toolParents = isTools ? getToolsParentRows() : [];

  const sites = isEng ? getSiteNames() : [];
  const employees = isEng ? getEmployeesForPicker() : [];

  const keyNorm = (k) =>
    String(k || "")
      .toLowerCase()
      .replace(/[\s_]+/g, "");

  const isEmpIdKey = (k) =>
    /(^empid$)|(^employeeid$)|employee.*id/.test(keyNorm(k));
  const isStartKey = (k) => keyNorm(k) === "startdate";
  const isEndKey = (k) => keyNorm(k) === "enddate";
  const isDurationKey = (k) => keyNorm(k) === "duration";

  const toISODateInput = (val) => {
    if (!val) return "";
    const s = String(val).trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const mm = String(m[1]).padStart(2, "0");
      const dd = String(m[2]).padStart(2, "0");
      return `${m[3]}-${mm}-${dd}`;
    }

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

  formFields.innerHTML =
    cols
      .map((col) => {
        const rawValue = currentRow[col.key] ?? "";
        const value = String(rawValue ?? "");
        const k = col.key;

        const isTextArea =
          String(k).toLowerCase().includes("description") ||
          String(k).toLowerCase().includes("notes") ||
          String(k).toLowerCase().includes("remarks");

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
          String(k).toLowerCase().includes("title") ||
          (isTools &&
            ["item", "uom", "qty", "location"].includes(
              String(k).trim().toLowerCase(),
            ));

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
      .join("") +
    (isTools
      ? `
      <div class="form-group full-width">
        <label class="form-label" style="display:flex;align-items:center;gap:10px;cursor:pointer;">
          <input
            type="checkbox"
            id="field___isChild"
            name="__isChild"
            value="true"
            ${
              String(currentRow.__rowType ?? "")
                .trim()
                .toLowerCase() === "child"
                ? "checked"
                : ""
            }
          />
          <span>This is a child row</span>
        </label>
      </div>

      <div
        class="form-group full-width"
        id="toolsParentWrap"
        style="display:none;"
      >
        <label class="form-label required" for="field___parentSn">Parent Item</label>
        <select id="field___parentSn" name="__parentSn" class="form-select">
          <option value="">Select parent item...</option>
          ${getToolsParentRows()
            .filter((r) => {
              const currentSn = String(currentRow["S.N"] ?? "").trim();
              const parentSn = String(r["S.N"] ?? "").trim();

              if (!parentSn) return false;

              if (mode === "edit" && currentSn && parentSn === currentSn) {
                return false;
              }

              return true;
            })
            .map((r) => {
              const sn = String(r["S.N"] ?? "").trim();
              const item = String(r["Item"] ?? "").trim();
              const selected =
                String(currentRow.__parentSn ?? "").trim() === sn
                  ? "selected"
                  : "";
              return `<option value="${esc(sn)}" ${selected}>${esc(sn)} - ${esc(item)}</option>`;
            })
            .join("")}
        </select>
        <span class="form-error">Please select parent item</span>
      </div>`
      : "");
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
      (x) => x.tagName === "INPUT" && x.name === "startDate",
    );
    const endEl = inputs.find(
      (x) => x.tagName === "INPUT" && x.name === "endDate",
    );
    const durEl = inputs.find(
      (x) => x.tagName === "INPUT" && x.name === "duration",
    );

    if (durEl) durEl.readOnly = true;

    const recomputeDurationSch = () => {
      if (!durEl || !startEl || !endEl) return;

      const start = startEl.value;
      const end = endEl.value;

      if (!start || !end) {
        durEl.value = "";
        return;
      }

      durEl.value = daysBetweenInclusive(start, end);
    };

    if (startEl) startEl.addEventListener("change", recomputeDurationSch);
    if (endEl) endEl.addEventListener("change", recomputeDurationSch);

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

  if (isTools) {
    const childEl = formFields.querySelector("#field___isChild");
    const parentWrapEl = formFields.querySelector("#toolsParentWrap");
    const parentEl = formFields.querySelector("#field___parentSn");

    const itemEl = formFields.querySelector('[name="Item"]');
    const uomEl = formFields.querySelector('[name="UoM"]');
    const qtyEl = formFields.querySelector('[name="Qty"]');
    const locEl = formFields.querySelector('[name="Location"]');
    const remarksEl = formFields.querySelector('[name="Remarks"]');

    const syncToolsChildMode = () => {
      const isChild = !!childEl?.checked;

      if (parentWrapEl) parentWrapEl.style.display = isChild ? "" : "none";
      if (parentEl) {
        parentEl.required = isChild;
        if (!isChild) parentEl.value = "";
      }

      if (itemEl)
        itemEl.placeholder = isChild ? "Enter child item..." : "Enter item...";
      if (uomEl) uomEl.placeholder = "Enter uom...";
      if (qtyEl) qtyEl.placeholder = "Enter qty...";
      if (locEl) locEl.placeholder = "Enter location...";
      if (remarksEl)
        remarksEl.placeholder = isChild
          ? "Enter child remarks..."
          : "Enter remarks...";
    };

    childEl?.addEventListener("change", syncToolsChildMode);
    syncToolsChildMode();
  }
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

  const sheet = getSheetForView(store.view);
  if (!sheet) return;

  const form = el("modalForm");
  const formData = new FormData(form);
  const data = {};

  for (const [key, value] of formData.entries()) {
    data[key] = value;
  }
  if (store.view === "schedule" || store.view === "gantt") {
    const mapped = {
      "Task Name": data.taskName ?? "",
      "% Completion": data.completion ?? "",
      Duration: data.duration ?? "",
      "Start Date": data.startDate ?? "",
      "End Date": data.endDate ?? "",
      Remarks: data.remarks ?? "",
    };

    delete data.taskName;
    delete data.completion;
    delete data.duration;
    delete data.startDate;
    delete data.endDate;
    delete data.remarks;

    Object.assign(data, mapped);
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
    if (store.view === "tools") {
      delete data.id;
      delete data.ID;
      delete data.Id;
      delete data["S.N"];
      delete data["s.n"];

      const isChild =
        String(data.__isChild ?? "")
          .trim()
          .toLowerCase() === "true";

      if (isChild && !String(data.__parentSn || "").trim()) {
        throw new Error("Please select parent item for child row.");
      }

      if (!isChild) {
        delete data.__parentSn;
      }
    }

    if (currentFormMode === "add") {
      if (store.view === "schedule" || store.view === "gantt") {
        delete data.ID;
        delete data.Id;
        delete data.id;
      }

      await apiCreateRow(
        sheet,
        store.view === "gantt" ? "schedule" : store.view,
        data,
      );
      toast("success", "Record created successfully");

      closeFormModal();
      await loadCurrentViewData({ goToLastPage: true });
      await loadAllData();
    } else {
      const updateId =
        store.view === "sites"
          ? data.Location || data.location || currentFormId
          : currentFormId;

      await apiUpdateRow(sheet, updateId, data);
      toast("success", "Record updated successfully");

      closeFormModal();
      await loadCurrentViewData({ resetPage: false });
      await loadAllData();
    }
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
  const sheet = getSheetForView(store.view);
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
      await loadCurrentViewData({ resetPage: false });
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

/* =========================
   SMOOTH SCHEDULE OVERRIDES
   ========================= */

async function apiCreateScheduleProject(name) {
  const res = await apiFetch(`${API}/excel/schedule-project`, {
    method: "POST",
    body: JSON.stringify({ name }),
  });
  const json = await res.json().catch(() => ({}));
  if (!json.ok) throw new Error(json.error || "Failed to create project");
  return json;
}

function _smoothClassifyRow(row) {
  if (row.__isSection) return "section";
  if (row.__isTitle) return "sub";
  const hasDates = !!(parseAnyDate(row.startDate) && parseAnyDate(row.endDate));
  if (!hasDates) {
    const t = String(row.taskName || "").trim();
    if (t && t === t.toUpperCase() && t.replace(/\s/g, "").length > 1)
      return "section";
    if (t) return "sub";
  }
  return row.__highlight === "yellow" ? "highlight" : "task";
}

function _smoothBuildTimeline(rows) {
  const dated = rows
    .map((r) => {
      const s = parseAnyDate(r.startDate);
      let e = parseAnyDate(r.endDate);
      if (!e && s && String(r.duration || "").trim()) {
        const n = Number(String(r.duration).replace(/[^\d.-]/g, ""));
        if (Number.isFinite(n) && n > 0) {
          e = new Date(s);
          e.setDate(e.getDate() + Math.max(0, Math.round(n) - 1));
        }
      }
      return { row: r, start: s, end: e };
    })
    .filter((x) => x.start && x.end);

  if (!dated.length)
    return { days: [], min: null, max: null, todayIdx: -1, datedRows: [] };

  let min = dated[0].start,
    max = dated[0].end;
  dated.forEach((x) => {
    if (x.start < min) min = x.start;
    if (x.end > max) max = x.end;
  });
  min = new Date(min);
  min.setDate(min.getDate() - 2);
  max = new Date(max);
  max.setDate(max.getDate() + 2);

  const days = [];
  const cursor = new Date(min);
  while (cursor <= max) {
    days.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayIdx = days.findIndex((d) => d.getTime() === today.getTime());
  return { days, min, max, todayIdx, datedRows: dated };
}

function _smoothMonthGroups(days) {
  const out = [];
  const M = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  let cur = null;
  days.forEach((d) => {
    const k = `${M[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
    if (!cur || cur.k !== k) {
      cur = { k, label: k, count: 0 };
      out.push(cur);
    }
    cur.count++;
  });
  return out;
}

/* ─── New-Project Modal (proper form, replaces prompt) ─────── */
(function _mountNewProjectModal() {
  if (document.getElementById("npOverlay")) return;

  const html = `
<div id="npOverlay" class="np-overlay" role="dialog" aria-modal="true" aria-labelledby="npTitle">
  <div class="np-card">
    <h2 class="np-title" id="npTitle">Create New Project</h2>
    <p class="np-desc">Enter a name — an empty Gantt chart will be created automatically.</p>
    <div class="np-field-group">
      <label for="npNameInput">Project Name <span style="color:#c62828">*</span></label>
      <input id="npNameInput" class="np-input" type="text" maxlength="80"
             placeholder="e.g. Upper Myagdi-2 (HPP)" autocomplete="off"/>
      <span class="np-err" id="npNameErr">Please enter a project name.</span>
    </div>
    <div class="np-footer">
      <button type="button" class="np-btn cancel" id="npCancel">Cancel</button>
      <button type="button" class="np-btn create" id="npCreate">Create Project</button>
    </div>
  </div>
</div>`;

  document.body.insertAdjacentHTML("beforeend", html);

  const overlay = document.getElementById("npOverlay");
  const nameInput = document.getElementById("npNameInput");
  const nameErr = document.getElementById("npNameErr");
  const createBtn = document.getElementById("npCreate");
  const cancelBtn = document.getElementById("npCancel");

  function openModal() {
    nameInput.value = "";
    nameInput.classList.remove("error");
    nameErr.classList.remove("show");
    createBtn.disabled = false;
    createBtn.textContent = "Create Project";
    overlay.classList.add("open");
    document.body.style.overflow = "hidden";
    setTimeout(() => nameInput.focus(), 60);
  }
  function closeModal() {
    overlay.classList.remove("open");
    document.body.style.overflow = "";
  }

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeModal();
  });
  cancelBtn.addEventListener("click", closeModal);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && overlay.classList.contains("open")) closeModal();
  });

  async function doCreate() {
    const name = nameInput.value.trim();
    if (!name) {
      nameInput.classList.add("error");
      nameErr.classList.add("show");
      nameInput.focus();
      return;
    }

    nameInput.classList.remove("error");
    nameErr.classList.remove("show");
    createBtn.disabled = true;
    createBtn.textContent = "Creating…";

    try {
      const json = await apiCreateScheduleProject(name);
      store.scheduleProjectSearch = "";
      await refreshScheduleState(json.sheet, { resetPage: true });
      closeModal();
      toast("success", "Project created", `${name} is ready`);
    } catch (err) {
      toast("error", "Create project failed", err.message || "");
      createBtn.disabled = false;
      createBtn.textContent = "Create Project";
    }
  }

  createBtn.addEventListener("click", doCreate);
  nameInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") doCreate();
  });

  window._openNewProjectModal = openModal;
})();
const editProjectBtn = document.createElement("button");
editProjectBtn.className = "btn ghost btn-sm";
editProjectBtn.textContent = "Edit Project";

const deleteProjectBtn = document.createElement("button");
deleteProjectBtn.className = "btn ghost btn-sm";
deleteProjectBtn.textContent = "Delete Project";

editProjectBtn.addEventListener("click", async () => {
  const active = store.activeScheduleSheet;
  if (!active) return;

  const current = getActiveScheduleProject();
  const currentTitle =
    current?.title || active.replace(/_Timeline$/i, "").replace(/_/g, " ");
  const nextName = window.prompt("Rename project", currentTitle);

  if (!nextName || !nextName.trim()) return;

  try {
    const json = await apiRenameScheduleProject(active, nextName.trim());

    delete store.scheduleProjectCache?.[active];
    await refreshScheduleState(json.sheet, { resetPage: true });
    toast(
      "success",
      "Project renamed",
      `${nextName.trim()} updated successfully`,
    );
  } catch (err) {
    toast("error", "Rename failed", err.message || "");
  }
});

deleteProjectBtn.addEventListener("click", async () => {
  const active = store.activeScheduleSheet;
  if (!active) return;

  const current = getActiveScheduleProject();
  const title = current?.title || active;

  const ok = window.confirm(
    `Delete project "${title}"?\n\nThis will remove the whole worksheet.`,
  );
  if (!ok) return;

  try {
    await apiDeleteScheduleProject(active);

    delete store.scheduleProjectCache?.[active];
    await loadScheduleProjects();
    store.activeScheduleSheet = getActiveScheduleSheet();
    saveActiveScheduleSheet(store.activeScheduleSheet);
    await refreshScheduleState(store.activeScheduleSheet, { resetPage: true });

    toast("success", "Project deleted", `${title} removed successfully`);
  } catch (err) {
    toast("error", "Delete failed", err.message || "");
  }
});

function renderScheduleSheet() {
  const mount = el("scheduleSheet");
  const header = el("scheduleHeader");
  if (!mount) return;
  if (header) header.style.display = "none";

  const projects = Array.isArray(store.scheduleProjects)
    ? store.scheduleProjects
    : [];
  const activeSheet = getActiveScheduleSheet();
  const activeProject =
    projects.find((p) => p.sheet === activeSheet) || projects[0] || null;
  const rows = Array.isArray(store.data) ? store.data : [];
  const dayW = Number(store.ganttCellW || window.__smoothDayW || 22);

  mount.innerHTML = "";
  mount.style.cssText =
    ""; /* Let body.is-schedule #scheduleSheet CSS class control layout */
  const shell = document.createElement("div");
  shell.className = "smooth-shell";
  mount.appendChild(shell);

  const top = document.createElement("div");
  top.className = "smooth-topbar";
  top.innerHTML = `<span class="smooth-label">Project</span>`;
  const search = document.createElement("input");
  search.className = "smooth-search";
  search.placeholder = "Filter projects...";
  search.value = store.scheduleProjectSearch || "";
  const select = document.createElement("select");
  select.className = "smooth-select";

  const filtered = projects.filter(
    (p) =>
      !search.value.trim() ||
      (p.title || p.sheet)
        .toLowerCase()
        .includes(search.value.trim().toLowerCase()),
  );
  const fillSelect = () => {
    const items = projects.filter(
      (p) =>
        !search.value.trim() ||
        (p.title || p.sheet)
          .toLowerCase()
          .includes(search.value.trim().toLowerCase()),
    );
    select.innerHTML = items
      .map(
        (p) =>
          `<option value="${esc(p.sheet)}" ${p.sheet === activeSheet ? "selected" : ""}>${esc(p.title || p.sheet)}</option>`,
      )
      .join("");
  };
  fillSelect();
  search.oninput = () => {
    store.scheduleProjectSearch = search.value;
    fillSelect();
  };
  select.onchange = async () => {
    store.activeScheduleSheet = select.value;
    delete store.scheduleProjectCache[select.value];
    await loadCurrentViewData({ resetPage: true });
  };

  const addTask = document.createElement("button");
  addTask.className = "btn primary btn-sm";
  addTask.textContent = "Add Row";
  addTask.onclick = () => _sgTask("add", null, getActiveScheduleSheet());

  const addProject = document.createElement("button");
  addProject.className = "btn ghost btn-sm";
  addProject.textContent = "+ New Project";
  addProject.onclick = () => _openNewProjectModal();

  top.append(search, select, addTask, addProject);
  shell.appendChild(top);

  const h = document.createElement("div");
  h.className = "smooth-head";
  h.innerHTML = `<div><div class="smooth-title">${esc(activeProject ? activeProject.title : "Schedule")}</div><div class="smooth-sub">Sheet: ${esc(activeProject ? activeProject.sheet : activeSheet || "")}</div></div><div class="smooth-tag">${rows.length} tasks</div>`;
  shell.appendChild(h);

  const viewport = document.createElement("div");
  viewport.className = "smooth-viewport";
  shell.appendChild(viewport);

  /* ── Bulletproof height: measure actual available space after paint ── */
  function _setViewportHeight() {
    const shellEl = viewport.parentElement; // smooth-shell
    if (!shellEl) return;
    const shellRect = shellEl.getBoundingClientRect();
    const topbarEl = shellEl.querySelector(".smooth-topbar");
    const headEl = shellEl.querySelector(".smooth-head");
    const topbarH = topbarEl ? topbarEl.getBoundingClientRect().height : 0;
    const headH = headEl ? headEl.getBoundingClientRect().height : 0;
    const available = shellRect.height - topbarH - headH;
    if (available > 80) {
      viewport.style.height = available + "px";
      viewport.style.flex = "none"; // let explicit height win
    } else {
      // fallback: measure from viewport bottom
      const vTop = viewport.getBoundingClientRect().top;
      const h = window.innerHeight - vTop - 4;
      if (h > 80) {
        viewport.style.height = h + "px";
        viewport.style.flex = "none";
      }
    }
  }
  /* Run after layout + after any async paint */
  requestAnimationFrame(() => {
    _setViewportHeight();
  });
  setTimeout(_setViewportHeight, 80);
  /* Re-run on window resize */
  if (!window.__schedViewportResizeHandler) {
    window.__schedViewportResizeHandler = () => {
      document.querySelectorAll(".smooth-viewport").forEach((vp) => {
        const shellEl = vp.parentElement;
        if (!shellEl) return;
        const shellRect = shellEl.getBoundingClientRect();
        const topbarH = (
          shellEl.querySelector(".smooth-topbar") || {
            getBoundingClientRect: () => ({ height: 0 }),
          }
        ).getBoundingClientRect().height;
        const headH = (
          shellEl.querySelector(".smooth-head") || {
            getBoundingClientRect: () => ({ height: 0 }),
          }
        ).getBoundingClientRect().height;
        const available = shellRect.height - topbarH - headH;
        if (available > 80) {
          vp.style.height = available + "px";
          vp.style.flex = "none";
        } else {
          const h = window.innerHeight - vp.getBoundingClientRect().top - 4;
          if (h > 80) {
            vp.style.height = h + "px";
            vp.style.flex = "none";
          }
        }
      });
    };
    window.addEventListener("resize", window.__schedViewportResizeHandler);
  }

  const leftCols = [
    {
      key: "taskName",
      label: "TASK NAME /<br>MILESTONE",
      width: 350,
      align: "left",
      fmt: (r) => r.taskName || "",
    },
    {
      key: "duration",
      label: "DURATION<br>(DAYS)",
      width: 82,
      align: "center",
      fmt: (r) => r.duration || "",
    },
    {
      key: "startDate",
      label: "START<br>DATE",
      width: 110,
      align: "center",
      fmt: (r) => _sgFD(r.startDate),
    },
    {
      key: "endDate",
      label: "FINISH<br>DATE",
      width: 110,
      align: "center",
      fmt: (r) => _sgFD(r.endDate),
    },
    {
      key: "remarks",
      label: "NOTES /<br>SECTION",
      width: 150,
      align: "left",
      fmt: (r) => r.remarks || "",
    },
    {
      key: "__actions",
      label: "ACTIONS",
      width: 110,
      align: "center",
      fmt: null,
    },
  ];
  const leftWidth = leftCols.reduce((a, c) => a + c.width, 0);

  if (!rows.length) {
    const emptyGrid = document.createElement("div");
    emptyGrid.className = "smooth-grid";
    emptyGrid.style.width = `${leftWidth + 600}px`;
    emptyGrid.style.setProperty("--dayw", `${dayW}px`);
    viewport.appendChild(emptyGrid);

    const emptyHeaderWrap = document.createElement("div");
    emptyHeaderWrap.className = "smooth-header-row";
    emptyGrid.appendChild(emptyHeaderWrap);

    const emptyLeftHead = document.createElement("div");
    emptyLeftHead.className = "smooth-left";
    emptyLeftHead.style.width = `${leftWidth}px`;
    emptyHeaderWrap.appendChild(emptyLeftHead);

    let rl = 0;
    leftCols.forEach((col) => {
      const cell = document.createElement("div");
      cell.className = "smooth-hcell";
      cell.style.width = `${col.width}px`;
      cell.style.minWidth = `${col.width}px`;
      cell.style.position = "sticky";
      cell.style.left = `${rl}px`;
      cell.style.zIndex = "16";
      cell.innerHTML = col.label;
      emptyLeftHead.appendChild(cell);
      rl += col.width;
    });

    // Placeholder timeline header (6 months from today)
    const emptyTlHead = document.createElement("div");
    emptyTlHead.className = "smooth-timeline-head";
    emptyTlHead.style.width = "600px";
    emptyHeaderWrap.appendChild(emptyTlHead);

    const today = new Date();
    const M = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];
    const phMonthRow = document.createElement("div");
    phMonthRow.style.display = "flex";
    for (let m = 0; m < 6; m++) {
      const d = new Date(today.getFullYear(), today.getMonth() + m, 1);
      const cell = document.createElement("div");
      cell.className = "smooth-hcell smooth-month";
      const daysInMonth = new Date(
        d.getFullYear(),
        d.getMonth() + 1,
        0,
      ).getDate();
      const w = Math.min(daysInMonth * dayW, 100);
      cell.style.width = `${w}px`;
      cell.style.minWidth = `${w}px`;
      cell.textContent = `${M[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
      phMonthRow.appendChild(cell);
    }
    emptyTlHead.appendChild(phMonthRow);

    const emptyMsg = document.createElement("div");
    emptyMsg.className = "smooth-empty";
    emptyMsg.innerHTML = `<svg style="margin-bottom:12px;opacity:.4" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg><br>No tasks yet. Click <strong>+ Add Row</strong> to add the first task.`;
    viewport.appendChild(emptyMsg);

    const rc = el("rowCount");
    if (rc) rc.textContent = "0 rows";
    return;
  }

  const timeline = _smoothBuildTimeline(rows);
  const timelineWidth = Math.max(600, timeline.days.length * dayW);

  const grid = document.createElement("div");
  grid.className = "smooth-grid";
  grid.style.width = `${leftWidth + timelineWidth}px`;
  grid.style.setProperty("--dayw", `${dayW}px`);
  viewport.appendChild(grid);

  const headerWrap = document.createElement("div");
  headerWrap.className = "smooth-header-row";
  grid.appendChild(headerWrap);

  const leftHead = document.createElement("div");
  leftHead.className = "smooth-left";
  leftHead.style.width = `${leftWidth}px`;
  headerWrap.appendChild(leftHead);

  let runningLeft = 0;
  leftCols.forEach((col) => {
    const cell = document.createElement("div");
    cell.className = "smooth-hcell";
    cell.style.width = `${col.width}px`;
    cell.style.minWidth = `${col.width}px`;
    cell.style.position = "sticky";
    cell.style.left = `${runningLeft}px`;
    cell.style.zIndex = "16";
    cell.innerHTML = col.label;
    leftHead.appendChild(cell);
    runningLeft += col.width;
  });

  const timelineHead = document.createElement("div");
  timelineHead.className = "smooth-timeline-head";
  timelineHead.style.width = `${timelineWidth}px`;
  headerWrap.appendChild(timelineHead);

  const monthRow = document.createElement("div");
  monthRow.style.display = "flex";
  _smoothMonthGroups(timeline.days).forEach((g) => {
    const d = document.createElement("div");
    d.className = "smooth-hcell smooth-month";
    d.style.width = `${g.count * dayW}px`;
    d.style.minWidth = `${g.count * dayW}px`;
    d.textContent = g.label.toUpperCase();
    monthRow.appendChild(d);
  });
  timelineHead.appendChild(monthRow);

  const dayRow = document.createElement("div");
  dayRow.style.display = "flex";
  timeline.days.forEach((d) => {
    const c = document.createElement("div");
    c.className = "smooth-hcell smooth-day";
    c.style.width = `${dayW}px`;
    c.style.minWidth = `${dayW}px`;
    c.textContent = d.getDate();
    dayRow.appendChild(c);
  });
  timelineHead.appendChild(dayRow);

  rows.forEach((row, idx) => {
    const cls = _smoothClassifyRow(row);
    const line = document.createElement("div");
    line.className = `smooth-row ${cls === "section" ? "is-section" : cls === "sub" ? "is-sub" : cls === "highlight" ? "is-highlight" : ""}`;
    grid.appendChild(line);

    const left = document.createElement("div");
    left.className = "smooth-left";
    left.style.width = `${leftWidth}px`;
    line.appendChild(left);

    let x = 0;
    leftCols.forEach((col) => {
      const cell = document.createElement("div");
      cell.className = "smooth-cell";
      cell.style.width = `${col.width}px`;
      cell.style.minWidth = `${col.width}px`;
      cell.style.position = "sticky";
      cell.style.left = `${x}px`;
      cell.style.zIndex = "6";
      cell.style.justifyContent =
        col.align === "center" ? "center" : "flex-start";

      if (col.key === "__actions" && cls !== "section" && cls !== "sub") {
        const acts = document.createElement("div");
        acts.className = "smooth-actions";
        acts.innerHTML = `<button class="smooth-act">Edit</button><button class="smooth-act del">Del</button>`;
        const [eb, db] = acts.querySelectorAll("button");
        eb.onclick = () => _sgTask("edit", row, activeSheet);
        db.onclick = () => _sgDel(row, activeSheet);
        cell.appendChild(acts);
      } else {
        const span = document.createElement("span");
        span.className = "smooth-name";
        span.textContent = col.fmt ? col.fmt(row) : "";
        cell.appendChild(span);
      }
      left.appendChild(cell);
      x += col.width;
    });

    const track = document.createElement("div");
    track.className = "smooth-track is-week";
    track.style.width = `${timelineWidth}px`;
    line.appendChild(track);

    if (timeline.todayIdx >= 0) {
      const t = document.createElement("div");
      t.className = "smooth-today";
      t.style.left = `${timeline.todayIdx * dayW}px`;
      track.appendChild(t);
    }

    const s = parseAnyDate(row.startDate);
    let e = parseAnyDate(row.endDate);
    if (!e && s && String(row.duration || "").trim()) {
      const n = Number(String(row.duration).replace(/[^\d.-]/g, ""));
      if (Number.isFinite(n) && n > 0) {
        e = new Date(s);
        e.setDate(e.getDate() + Math.max(0, Math.round(n) - 1));
      }
    }
    if (s && e && timeline.min) {
      const leftPx = Math.round((s - timeline.min) / 86400000) * dayW;
      const widthPx = Math.max(1, Math.round((e - s) / 86400000) + 1) * dayW;
      const bar = document.createElement("div");
      bar.className = `smooth-bar ${cls === "highlight" ? "highlight" : cls === "section" ? "section" : cls === "sub" ? "sub" : ""}`;
      bar.style.left = `${leftPx}px`;
      bar.style.width = `${widthPx}px`;
      track.appendChild(bar);
    }
  });

  // Auto-scroll so today is roughly centred in the Gantt timeline
  requestAnimationFrame(() => {
    if (timeline.todayIdx >= 0) {
      viewport.scrollLeft = Math.max(
        0,
        leftWidth + timeline.todayIdx * dayW - viewport.clientWidth * 0.45,
      );
    }
  });

  const rc = el("rowCount");
  if (rc) rc.textContent = `${rows.length} rows`;
}

function openProjectModal() {
  const overlay = document.getElementById("projectModalOverlay");
  const input = document.getElementById("projectNameInput");
  if (!overlay) return;
  overlay.style.display = "flex";
  document.body.style.overflow = "hidden";
  if (input) {
    input.value = "";
    setTimeout(() => input.focus(), 0);
  }
}

function closeProjectModal() {
  const overlay = document.getElementById("projectModalOverlay");
  const form = document.getElementById("projectModalForm");
  if (!overlay) return;
  overlay.style.display = "none";
  document.body.style.overflow = "";
  form?.reset();
}

document
  .getElementById("projectModalClose")
  ?.addEventListener("click", closeProjectModal);
document
  .getElementById("projectModalCancel")
  ?.addEventListener("click", closeProjectModal);

document
  .getElementById("projectModalOverlay")
  ?.addEventListener("click", (e) => {
    if (e.target.id === "projectModalOverlay") closeProjectModal();
  });
document
  .getElementById("projectModalForm")
  ?.addEventListener("submit", async (e) => {
    e.preventDefault();

    const input = document.getElementById("projectNameInput");
    const name = String(input?.value || "").trim();
    if (!name) return;

    try {
      const res = await apiFetch("/api/excel/schedule-project", {
        method: "POST",
        body: JSON.stringify({ name }),
      });

      const json = await res.json().catch(() => ({}));
      if (!json.ok) throw new Error(json.error || "Failed to create project");

      closeProjectModal();
      await loadScheduleProjects();

      store.activeScheduleSheet = json.sheet;
      await loadCurrentViewData({ resetPage: true });

      toast("success", "Project created", `${name} is ready`);
    } catch (err) {
      toast("error", "Project create failed", err.message || "");
    }
  });

function applyViewLayout() {
  const isOverview = store.view === "overview";
  const isSchedule = store.view === "schedule" || store.view === "gantt";

  document.body.classList.toggle("is-schedule", isSchedule);

  const overviewView = document.getElementById("overviewView");
  const tableView = document.getElementById("tableView");
  const scheduleLayout = document.getElementById("scheduleLayout");
  const scheduleScroller = document.getElementById("scheduleScroller");
  const scheduleSheet = document.getElementById("scheduleSheet");
  const tableCard = document.querySelector("#tableView .table-card");
  const tableToolbar = document.querySelector("#tableView .table-toolbar");
  const stats = document.getElementById("stats");
  const filterBar = document.getElementById("filterBar");
  const advancedFilters = document.getElementById("advancedFilters");
  const paginationBar = document.getElementById("paginationBar");
  const defaultWrap = document.getElementById("defaultTableWrap");
  const scheduleHeader = document.getElementById("scheduleHeader");
  const btnOpenGantt = document.getElementById("btnOpenGantt");

  [
    tableView,
    tableCard,
    scheduleLayout,
    scheduleScroller,
    scheduleSheet,
  ].forEach((el) => {
    if (el) el.style.cssText = "";
  });

  if (overviewView) overviewView.style.display = isOverview ? "" : "none";
  if (tableView) tableView.style.display = isOverview ? "none" : "";

  if (isSchedule) {
    // Explicitly set flex — #scheduleLayout starts as display:none in HTML
    if (scheduleLayout) scheduleLayout.style.display = "flex";
    if (scheduleScroller) scheduleScroller.style.display = "flex";
    if (scheduleSheet) {
      scheduleSheet.style.display = "flex";
      scheduleSheet.classList.add("schedule-screen");
    }
    if (scheduleHeader) scheduleHeader.style.display = "none";
    if (tableToolbar) tableToolbar.style.display = "";
    if (stats) stats.style.display = "none";
    if (filterBar) filterBar.style.display = "none";
    if (advancedFilters) advancedFilters.style.display = "none";
    if (paginationBar) paginationBar.style.display = "none";
    if (defaultWrap) defaultWrap.style.display = "none";
  } else {
    if (scheduleSheet) {
      scheduleSheet.style.display = "";
      scheduleSheet.classList.remove("schedule-screen");
    }
    if (scheduleLayout) scheduleLayout.style.display = "none";
    if (scheduleScroller) scheduleScroller.style.display = "";
    if (scheduleHeader) scheduleHeader.style.display = "none";
    if (tableCard) tableCard.style.display = isOverview ? "none" : "";
    if (tableToolbar) tableToolbar.style.display = "";
    if (defaultWrap) defaultWrap.style.display = !isOverview ? "" : "none";
    if (stats) stats.style.display = "";
    if (filterBar) filterBar.style.display = "";
    if (advancedFilters) advancedFilters.style.display = "none";
    if (paginationBar) paginationBar.style.display = "";
  }
  if (btnOpenGantt) btnOpenGantt.style.display = "none";
}

/* =========================
   REVISED SCHEDULE UI OVERRIDES
   ========================= */
function _scheduleTimelineFromRows(rows) {
  const dated = (Array.isArray(rows) ? rows : [])
    .map((r) => {
      const s = parseAnyDate(r.startDate);
      let e = parseAnyDate(r.endDate);
      if (!e && s && String(r.duration || "").trim()) {
        const n = Number(String(r.duration).replace(/[^\d.-]/g, ""));
        if (Number.isFinite(n) && n > 0) {
          e = new Date(s);
          e.setDate(e.getDate() + Math.max(0, Math.round(n) - 1));
        }
      }
      return { row: r, start: s, end: e };
    })
    .filter((x) => x.start && x.end);

  if (!dated.length) {
    const base = new Date();
    base.setHours(0, 0, 0, 0);
    const min = new Date(base.getFullYear(), base.getMonth(), 1);
    const max = new Date(base.getFullYear(), base.getMonth() + 2, 0);
    return { min, max, days: _daysBetween(min, max), datedRows: [] };
  }

  let min = dated[0].start;
  let max = dated[0].end;
  for (const item of dated) {
    if (item.start < min) min = item.start;
    if (item.end > max) max = item.end;
  }

  min = new Date(min.getFullYear(), min.getMonth(), 1);
  max = new Date(max.getFullYear(), max.getMonth() + 1, 0);
  return { min, max, days: _daysBetween(min, max), datedRows: dated };
}

function _daysBetween(min, max) {
  const out = [];
  const cursor = new Date(min);
  cursor.setHours(0, 0, 0, 0);
  while (cursor <= max) {
    out.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return out;
}

function _scheduleMonthGroups(days) {
  const M = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  const out = [];
  let current = null;
  for (const d of days) {
    const key = `${d.getFullYear()}-${d.getMonth()}`;
    if (!current || current.key !== key) {
      current = {
        key,
        label: `${M[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`,
        count: 0,
      };
      out.push(current);
    }
    current.count += 1;
  }
  return out;
}

function _syncScheduleScroll(a, b) {
  if (!a || !b) return;
  let busy = false;
  const sync = (src, dst) => {
    src.addEventListener(
      "scroll",
      () => {
        if (busy) return;
        busy = true;
        dst.scrollLeft = src.scrollLeft;
        requestAnimationFrame(() => {
          busy = false;
        });
      },
      { passive: true },
    );
  };
  sync(a, b);
  sync(b, a);
}

function renderScheduleSheet() {
  const mount = el("scheduleSheet");
  const header = el("scheduleHeader");
  if (!mount) return;
  if (header) header.style.display = "none";

  const projects = Array.isArray(store.scheduleProjects)
    ? store.scheduleProjects
    : [];
  const activeSheet = getActiveScheduleSheet();
  const activeProject =
    projects.find((p) => p.sheet === activeSheet) || projects[0] || null;
  const rows = Array.isArray(store.data) ? store.data : [];
  const dayW = Math.max(
    22,
    Number(store.ganttCellW || window.__smoothDayW || 30),
  );

  mount.innerHTML = "";
  mount.style.cssText = "";

  const shell = document.createElement("div");
  shell.className = "rev-schedule-shell";
  mount.appendChild(shell);

  const toolbar = document.createElement("div");
  toolbar.className = "rev-schedule-toolbar";
  shell.appendChild(toolbar);

  const toolbarLeft = document.createElement("div");
  toolbarLeft.className = "rev-schedule-toolbar__left";
  toolbar.appendChild(toolbarLeft);

  const label = document.createElement("span");
  label.className = "rev-schedule-label";
  label.textContent = "Project";

  const search = document.createElement("input");
  search.className = "rev-schedule-search";
  search.placeholder = "Filter projects...";
  search.value = store.scheduleProjectSearch || "";

  const select = document.createElement("select");
  select.className = "rev-schedule-select";

  const fillSelect = () => {
    const q = String(search.value || "")
      .trim()
      .toLowerCase();
    const items = projects.filter(
      (p) =>
        !q ||
        String(p.title || p.sheet || "")
          .toLowerCase()
          .includes(q),
    );
    select.innerHTML = items
      .map(
        (p) =>
          `<option value="${esc(p.sheet)}" ${p.sheet === activeSheet ? "selected" : ""}>${esc(p.title || p.sheet)}</option>`,
      )
      .join("");
  };
  fillSelect();
  search.oninput = () => {
    store.scheduleProjectSearch = search.value;
    fillSelect();
  };
  select.onchange = async () => {
    store.activeScheduleSheet = select.value;
    await loadCurrentViewData({ resetPage: true });
  };

  toolbarLeft.append(label, search, select);

  const toolbarRight = document.createElement("div");
  toolbarRight.className = "rev-schedule-toolbar__right";
  toolbar.appendChild(toolbarRight);

  const addTask = document.createElement("button");
  addTask.className = "btn primary btn-sm";
  addTask.textContent = " Add Row";
  addTask.onclick = () => _sgTask("add", null, getActiveScheduleSheet());

  const addProject = document.createElement("button");
  addProject.className = "rev-schedule-btn ghost";
  addProject.textContent = "+ New Project";
  addProject.onclick = () => _openNewProjectModal();

  toolbarRight.append(editProjectBtn, deleteProjectBtn, addTask, addProject);

  const meta = document.createElement("div");
  meta.className = "rev-schedule-meta";
  meta.innerHTML = `
    <div>
      <div class="rev-schedule-title">${esc(activeProject ? activeProject.title : "Schedule")}</div>
      <div class="rev-schedule-sub">Sheet: ${esc(activeProject ? activeProject.sheet : activeSheet || "")}</div>
    </div>
    <div class="rev-schedule-chip">${rows.length} rows</div>
  `;
  shell.appendChild(meta);

  const topScroll = document.createElement("div");
  topScroll.className = "rev-top-scroll";
  topScroll.innerHTML = '<div class="rev-top-scroll__inner"></div>';
  shell.appendChild(topScroll);

  const viewport = document.createElement("div");
  viewport.className = "rev-schedule-viewport";
  shell.appendChild(viewport);

  const leftCols = [
    {
      key: "taskName",
      label: "TASK NAME /<br>MILESTONE",
      width: 350,
      align: "left",
      fmt: (r) => r.taskName || "",
    },
    {
      key: "duration",
      label: "DURATION<br>(DAYS)",
      width: 72,
      align: "center",
      fmt: (r) => r.duration || "",
    },
    {
      key: "startDate",
      label: "START<br>DATE",
      width: 84,
      align: "center",
      fmt: (r) => _sgFD(r.startDate),
    },
    {
      key: "endDate",
      label: "FINISH<br>DATE",
      width: 84,
      align: "center",
      fmt: (r) => _sgFD(r.endDate),
    },
    {
      key: "remarks",
      label: "NOTES /<br>SECTION",
      width: 130,
      align: "left",
      fmt: (r) => r.remarks || "",
    },
    {
      key: "__actions",
      label: "ACTIONS",
      width: 106,
      align: "center",
      fmt: null,
    },
  ];
  const leftWidth = leftCols.reduce((sum, col) => sum + col.width, 0);
  const timeline = _scheduleTimelineFromRows(rows);
  const timelineWidth = Math.max(1100, timeline.days.length * dayW);
  const totalWidth = leftWidth + timelineWidth;
  topScroll.firstElementChild.style.width = `${totalWidth}px`;

  const grid = document.createElement("div");
  grid.className = "rev-schedule-grid";
  grid.style.width = `${totalWidth}px`;
  viewport.appendChild(grid);

  const headerRow = document.createElement("div");
  headerRow.className = "rev-schedule-header-row";
  grid.appendChild(headerRow);

  const leftHeader = document.createElement("div");
  leftHeader.className = "rev-schedule-left rev-schedule-left--head";
  leftHeader.style.width = `${leftWidth}px`;
  headerRow.appendChild(leftHeader);

  let runningLeft = 0;
  leftCols.forEach((col) => {
    const cell = document.createElement("div");
    cell.className = "rev-schedule-hcell";
    cell.style.width = `${col.width}px`;
    cell.style.minWidth = `${col.width}px`;
    cell.style.left = `${runningLeft}px`;
    cell.innerHTML = col.label;
    leftHeader.appendChild(cell);
    runningLeft += col.width;
  });

  const timelineHead = document.createElement("div");
  timelineHead.className = "rev-schedule-timeline-head";
  timelineHead.style.width = `${timelineWidth}px`;
  headerRow.appendChild(timelineHead);

  const monthsRow = document.createElement("div");
  monthsRow.className = "rev-months-row";
  _scheduleMonthGroups(timeline.days).forEach((group) => {
    const cell = document.createElement("div");
    cell.className = "rev-schedule-hcell rev-schedule-hcell--month";
    cell.style.width = `${group.count * dayW}px`;
    cell.style.minWidth = `${group.count * dayW}px`;
    cell.textContent = group.label.toUpperCase();
    monthsRow.appendChild(cell);
  });
  timelineHead.appendChild(monthsRow);

  const daysRow = document.createElement("div");
  daysRow.className = "rev-days-row";
  timeline.days.forEach((d) => {
    const cell = document.createElement("div");
    cell.className = "rev-schedule-hcell rev-schedule-hcell--day";
    cell.style.width = `${dayW}px`;
    cell.style.minWidth = `${dayW}px`;
    cell.textContent = d.getDate();
    daysRow.appendChild(cell);
  });
  timelineHead.appendChild(daysRow);

  if (!rows.length) {
    const empty = document.createElement("div");
    empty.className = "rev-schedule-empty";
    empty.innerHTML =
      '<div class="rev-schedule-empty__title">No tasks yet</div><div class="rev-schedule-empty__sub">Use <strong>+ Add Row</strong> to add the first activity for this project.</div>';
    shell.appendChild(empty);
    _syncScheduleScroll(topScroll, viewport);
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayIdx = timeline.days.findIndex(
    (d) => d.getTime() === today.getTime(),
  );

  rows.forEach((row) => {
    const kind = _smoothClassifyRow(row);
    const rowEl = document.createElement("div");
    rowEl.className = `rev-schedule-row ${kind === "section" ? "is-section" : kind === "sub" ? "is-sub" : kind === "highlight" ? "is-highlight" : ""}`;
    grid.appendChild(rowEl);

    const left = document.createElement("div");
    left.className = "rev-schedule-left";
    left.style.width = `${leftWidth}px`;
    rowEl.appendChild(left);

    let stickyLeft = 0;
    leftCols.forEach((col) => {
      const cell = document.createElement("div");
      cell.className = "rev-schedule-cell";
      cell.dataset.col = col.key;
      cell.style.width = `${col.width}px`;
      cell.style.minWidth = `${col.width}px`;
      cell.style.left = `${stickyLeft}px`;
      cell.style.justifyContent =
        col.align === "center" ? "center" : "flex-start";

      if (col.key === "__actions" && kind !== "section" && kind !== "sub") {
        const acts = document.createElement("div");
        acts.className = "rev-schedule-actions";
        acts.innerHTML =
          '<button class="rev-act-btn">Edit</button><button class="rev-act-btn del">Del</button>';
        const [editBtn, delBtn] = acts.querySelectorAll("button");
        editBtn.onclick = () => _sgTask("edit", row, activeSheet);
        delBtn.onclick = () => deleteScheduleTaskRow(activeSheet, row);
        cell.appendChild(acts);
      } else {
        const span = document.createElement("span");
        span.className = "rev-schedule-cell__text";

        if (
          col.key === "startDate" ||
          col.key === "endDate" ||
          col.key === "duration"
        ) {
          span.style.whiteSpace = "nowrap";
          span.style.overflow = "visible";
          span.style.textOverflow = "clip";
          span.style.fontSize = "12px";
        }

        span.textContent = col.fmt ? col.fmt(row) : "";
        cell.appendChild(span);
      }

      left.appendChild(cell);
      stickyLeft += col.width;
    });

    const timelineRow = document.createElement("div");
    timelineRow.className = "rev-schedule-track";
    timelineRow.style.width = `${timelineWidth}px`;
    rowEl.appendChild(timelineRow);

    if (todayIdx >= 0) {
      const todayLine = document.createElement("div");
      todayLine.className = "rev-schedule-today";
      todayLine.style.left = `${todayIdx * dayW}px`;
      timelineRow.appendChild(todayLine);
    }

    const start = parseAnyDate(row.startDate);
    let end = parseAnyDate(row.endDate);
    if (!end && start && String(row.duration || "").trim()) {
      const n = Number(String(row.duration).replace(/[^\d.-]/g, ""));
      if (Number.isFinite(n) && n > 0) {
        end = new Date(start);
        end.setDate(end.getDate() + Math.max(0, Math.round(n) - 1));
      }
    }

    if (start && end && timeline.min) {
      const leftPx = Math.round((start - timeline.min) / 86400000) * dayW;
      const widthPx = Math.max(
        dayW,
        (Math.round((end - start) / 86400000) + 1) * dayW,
      );
      const bar = document.createElement("div");
      bar.className = `rev-schedule-bar ${kind === "highlight" ? "highlight" : kind === "section" ? "section" : kind === "sub" ? "sub" : ""}`;
      bar.style.left = `${leftPx}px`;
      bar.style.width = `${widthPx}px`;
      timelineRow.appendChild(bar);
    }
  });

  _syncScheduleScroll(topScroll, viewport);

  const firstDated = timeline.datedRows[0];
  requestAnimationFrame(() => {
    viewport.scrollLeft = 0;
    topScroll.scrollLeft = 0;
  });

  const rc = el("rowCount");
  if (rc) rc.textContent = `${rows.length} rows`;
}
