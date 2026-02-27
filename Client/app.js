"use strict";

const API = "/api";

const SHEET_MAP = {
  overview: "profiles",
  employees: "profiles",
  engagements: "Engagements",
  sites: "Sites_List",
  schedule: "Seti_Updated",
};

const ID_PREFIX_MAP = {
  employees: "EMP",
  engagements: "ENG",
  sites: "SITE",
  schedule: "SETI",
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
  if (options.body && !(options.body instanceof FormData) && !headers.has("Content-Type")) {
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
  const idPrefix = ID_PREFIX_MAP[view] || "ROW";
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

// ================================================================
// DATA LOADING
// ================================================================

function inferColumnsFromData(rows, viewName) {
  const first = rows[0] || {};

  return Object.keys(first)
    .filter((key) => {
      const lower = key.toLowerCase();

      // Hide ID for Sites (already there)
      if (viewName === "sites" && lower === "id") return false;

      // 🚀 Hide ID for Schedule (Seti_Updated)
      if (viewName === "schedule" && lower === "id") return false;

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
  await loadCurrentViewData();
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
    if (store.view === "schedule") rows = rows.map(computeScheduleFields);
    store.data = rows;

    if (store.view === "employees") store.allEmployees = rows;
    if (store.view === "engagements") store.allEngagements = rows;
    if (store.view === "sites") store.allSites = rows;
    if (store.view === "schedule") store.allSchedule = rows;
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

  // Detect fields
  const startKey =
    keys.find((k) => /starting\s*date/i.test(k)) ||
    keys.find((k) => /start\s*date/i.test(k)) ||
    keys.find((k) => /start/i.test(k) && /date/i.test(k)) ||
    keys.find((k) => /^date$/i.test(k)) ||
    null;

  const endKey =
    keys.find((k) => /end\s*date/i.test(k)) ||
    keys.find((k) => /ending\s*date/i.test(k)) ||
    keys.find((k) => /end/i.test(k) && /date/i.test(k)) ||
    null;

  const durationKey =
    keys.find((k) => /duration/i.test(k)) || "Duration (Days)";

  const statusKey = keys.find((k) => /status|phase/i.test(k)) || "status";

  const groupKey = keys.find((k) => /group/i.test(k)) || null;

  // Parse dates
  const start = startKey ? parseAnyDate(obj[startKey]) : null;
  const end = endKey ? parseAnyDate(obj[endKey]) : null;

  // Normalize date display (short yyyy-mm-dd)
  if (startKey)
    obj[startKey] = start
      ? toISODate(start)
      : obj[startKey]
        ? String(obj[startKey]).slice(0, 10)
        : "";

  if (endKey) obj[endKey] = end ? toISODate(end) : "";

  const today = startOfDay(new Date());

  // Calculate duration
  const duration = start
    ? end
      ? daysInclusive(start, end)
      : daysInclusive(start, today)
    : "";

  obj[durationKey] = duration === "" ? "" : String(duration);

  const hasGroup = groupKey && String(obj[groupKey] ?? "").trim() !== "";

  if (!hasGroup) {
    obj[statusKey] = "";
  } else {
    obj[statusKey] = end ? "Completed" : start ? "Active" : "";
  }

  return obj;
}
// ================================================================
// RENDER MAIN
// ================================================================

function render() {
  if (store.view === "overview") {
    renderOverviewDashboard();
  } else {
    renderTableView();
  }
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

  // Final fallback: use ID sequence (EMP-00012, ENG-00007, SETI-00003)
  return idSeq(row.id);
}

function formatActivityTime(row) {
  const t = activityTimeValue(row);
  // if it's an epoch (ms), show date; if it's just a small id seq, show "Recently"
  if (t > 1000000000) return new Date(t).toLocaleDateString();
  return "Recently";
}

// ================================================================
// TABLE VIEW RENDERING
// ================================================================

function renderTableView() {
  el("overviewView").style.display = "none";
  el("tableView").style.display = "";

  renderStats();
  renderFilterChips();
  renderAdvancedFilterPanel();
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

  const downloadBtn = document.createElement("button");
  downloadBtn.className = "btn ghost btn-sm";
  downloadBtn.textContent = "Download Excel";
  downloadBtn.onclick = () => {
    window.location.href = `${API}/excel/download`;
  };

  const addBtn = document.createElement("button");
  addBtn.className = "btn primary btn-sm";
  addBtn.textContent = "Add Row";
  addBtn.onclick = () => openFormModal("add", null);

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
  const cols = cfg.columns;
  const term = store.globalSearch.trim().toLowerCase();
  const sheet = SHEET_MAP[store.view];
  const showActions = store.view !== "overview" && !!sheet;

  if (!cols.length) {
    el("thead").innerHTML = "";
  } else {
    el("thead").innerHTML = `<tr>${cols
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
    const hasData = store.data.length > 0;
    el("tbody").innerHTML = `
      <tr class="empty-row">
        <td colspan="${(cols.length || 1) + (showActions ? 1 : 0)}">
          <span class="empty-row-label">
            ${hasData ? "No records match the current filters" : "No data available yet"}
          </span>
        </td>
      </tr>`;
    return;
  }

  el("tbody").innerHTML = paged
    .map((row) => {
      const tds = cols
        .map((col) => {
          const raw = row[col.key] ?? "";

          if (isStatusKey(col.key)) {
            return `<td>${statusBadgeHTML(raw)}</td>`;
          }

          return col.type === "number"
            ? `<td class="mono">${esc(raw)}</td>`
            : col.type === "date"
              ? `<td class="mono">${esc(formatDateShort(raw))}</td>`
              : `<td>${highlight(raw, term)}</td>`;
        })
        .join("");

      const rowId = store.view === "sites"
       ? row.Location || row.location || ""
       : row.id || "";
      const actions = showActions
        ? `
<td class="mono">
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

  // Update store
  store.view = view;

  // Remove active class from all nav links
  document.querySelectorAll(".nav-link").forEach((link) => {
    link.classList.remove("active");
  });

  // Add active class to selected
  const activeNav = document.querySelector(`.nav-link[data-view="${view}"]`);
  if (activeNav) activeNav.classList.add("active");

  // Hide all view sections
  document.querySelectorAll(".view-section").forEach((section) => {
    section.style.display = "none";
  });

  // Show current view section
  const activeSection = document.getElementById(view);
  if (activeSection) activeSection.style.display = "block";
  localStorage.setItem("activeView", view);

  history.replaceState(
    null,
    "",
    `${location.pathname}${location.search}#${view}`,
  );

  // Load data for that page if needed
  await loadCurrentViewData();
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

  const cols = views[store.view].columns.filter(
    (c) => c.key !== "id" && c.key !== "createdAt" && c.key !== "updatedAt",
  );

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
                String(r.Location || r.location || "").trim().toLowerCase() ===
                String(rowId).trim().toLowerCase()
            )
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
      await apiCreateRow(sheet, store.view, data);
      toast("success", "Record created successfully");
    } else {
      const updateId =
      store.view === "sites"
    ? (data.Location || data.location || currentFormId)
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
  const initialView =
    getViewFromHash() || localStorage.getItem("activeView") || DEFAULT_VIEW;
  await setView(initialView);
}

// Run once when DOM is ready
document.addEventListener("DOMContentLoaded", () => {
  init().catch((err) => console.error("init failed:", err));
});

init();

document.addEventListener("DOMContentLoaded", () => {

  const btn = document.getElementById("profileBtn");
  const menu = document.getElementById("profileMenu");
  const logoutBtn = document.getElementById("logoutBtn");

  if (!btn || !menu) return;

  btn.addEventListener("click", () => {
    menu.classList.toggle("show");
  });

  document.addEventListener("click", (e) => {
    if (!btn.contains(e.target) && !menu.contains(e.target)) {
      menu.classList.remove("show");
    }
  });

  logoutBtn?.addEventListener("click", async () => {
    await fetch("/api/auth/logout", {
      method: "POST",
      credentials: "include"
    });
    window.location.href = "/login";
  });

});