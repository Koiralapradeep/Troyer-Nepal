import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";
import { withWriteLock } from "./excel.js";

console.log("Using excel.store.js from:", import.meta.url);

const STORAGE_DIR = path.resolve("storage");
const BACKUP_DIR = path.join(STORAGE_DIR, "backups");
const UPLOADS_DIR = path.join(STORAGE_DIR, "uploads");
const MASTER_PATH = path.join(STORAGE_DIR, "data.xlsx");

const text = (v) => String(v ?? "").trim();

function findHeaderKey(sample, patterns) {
  const keys = Object.keys(sample || {});
  return (
    keys.find((k) => patterns.some((re) => re.test(String(k)))) || null
  );
}

function getAllObjects(ws) {
  const last = ws.actualRowCount || ws.rowCount || 1;
  const items = [];
  for (let r = 2; r <= last; r++) items.push(rowToObject(ws, r));
  return items;
}

function normKey(v) {
  return text(v)
    .replace(/\u00A0/g, " ")
    .replace(/[_\-]+/g, " ")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function ensureDirs() {
  fs.mkdirSync(STORAGE_DIR, { recursive: true });
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

function todayStamp() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function newestUploadPath() {
  if (!fs.existsSync(UPLOADS_DIR)) return null;
  const files = fs
    .readdirSync(UPLOADS_DIR)
    .map((f) => path.join(UPLOADS_DIR, f))
    .filter((p) => fs.existsSync(p) && fs.statSync(p).isFile())
    .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);

  return files[0] || null;
}

function syncMasterFromUploads() {
  ensureDirs();
  const newest = newestUploadPath();
  if (!newest) return { synced: false };

  try {
    const masterStat = fs.existsSync(MASTER_PATH) ? fs.statSync(MASTER_PATH) : null;
    const upStat = fs.statSync(newest);

    if (!masterStat || upStat.mtimeMs > masterStat.mtimeMs) {
      fs.copyFileSync(newest, MASTER_PATH);
      return { synced: true, from: newest };
    }

    return { synced: false };
  } catch (e) {
    return { synced: false, error: String(e.message || e) };
  }
}

function startOfDay(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function daysInclusive(a, b) {
  const ms = startOfDay(b).getTime() - startOfDay(a).getTime();
  return Math.floor(ms / 86400000) + 1;
}

function toISODate(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}


function parseAnyDate(v) {
  if (!v) return null;

  if (v instanceof Date && !isNaN(v.getTime())) {
    return startOfDay(v);
  }

  const s = String(v).trim();
  if (!s) return null;

  // Detect Excel serial number FIRST
  const num = Number(s);
  if (!isNaN(num) && num > 20000 && num < 60000) {
    const base = new Date(Date.UTC(1899, 11, 30));
    const dt = new Date(base.getTime() + num * 86400000);
    return startOfDay(dt);
  }

  // ISO format
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return isNaN(d.getTime()) ? null : startOfDay(d);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return startOfDay(d);

  return null;
}

function computeEngagementFields(obj) {
  const start = parseAnyDate(obj["Starting Date"]);
  const end = parseAnyDate(obj["End Date"]);

  const status = end ? "Completed" : start ? "Active" : "";

  const today = startOfDay(new Date());
  const duration =
    start ? (end ? daysInclusive(start, end) : daysInclusive(start, today)) : "";

  return {
    ...obj,
    "Starting Date": start
      ? toISODate(start)
      : obj["Starting Date"]
      ? String(obj["Starting Date"]).slice(0, 10)
      : "",
    "End Date": end ? toISODate(end) : "",
    "Duration (Days)": duration === "" ? "" : String(duration),
    status,
  };
}
// ===============================
//  Employee bundle workbook
// ===============================
export async function exportEmployeeWorkbookBuffer(employeeId) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();

    const wsProfiles = getSheetOrThrow(wb, "profiles");
    const wsEng = getSheetOrThrow(wb, "Engagements");
    const wsSites = getSheetOrThrow(wb, "Sites_List");
    const wsSeti = wb.getWorksheet("Seti_Updated");

    // Ensure columns are correct
    prepareSheet(wsProfiles, "profiles");
    prepareSheet(wsEng, "Engagements");
    prepareSheet(wsSites, "Sites_List");
    if (wsSeti) prepareSheet(wsSeti, "Seti_Updated");

    // ---- Find employee in profiles by id
    const pLast = wsProfiles.actualRowCount || wsProfiles.rowCount || 1;
    let employeeRow = null;

    for (let r = 2; r <= pLast; r++) {
      const obj = rowToObject(wsProfiles, r);
      if (String(obj.id || "").trim().toLowerCase() === String(employeeId).trim().toLowerCase()) {
        employeeRow = obj;
        break;
      }
    }

    if (!employeeRow) throw new Error(`Employee not found: ${employeeId}`);

    const employeeName = String(employeeRow["Employee Name"] || "").trim().toLowerCase();

    // ---- Filter engagements by Employee ID
    const eLast = wsEng.actualRowCount || wsEng.rowCount || 1;
    const myEngagements = [];

    for (let r = 2; r <= eLast; r++) {
      const obj = computeEngagementFields(rowToObject(wsEng, r));
      const eid = String(obj["Employee ID"] || "").trim().toLowerCase();
      if (eid === String(employeeId).trim().toLowerCase()) {
        myEngagements.push(obj);
      }
    }

    // ---- Collect sites from engagements (Site Engaged)
    const siteSet = new Set(
      myEngagements
        .map((e) => String(e["Site Engaged"] || "").trim().toLowerCase())
        .filter(Boolean),
    );

    // ---- Filter Sites_List by Location
    const sLast = wsSites.actualRowCount || wsSites.rowCount || 1;
    const mySites = [];
    for (let r = 2; r <= sLast; r++) {
      const obj = rowToObject(wsSites, r);
      const loc = String(obj["Location"] || "").trim().toLowerCase();
      if (loc && siteSet.has(loc)) mySites.push(obj);
    }

    // ---- Filter schedule rows (Seti_Updated) by Responsibility/Owner/Assigned (best-effort)
    const mySchedule = [];
    if (wsSeti) {
      const last = wsSeti.actualRowCount || wsSeti.rowCount || 1;
      const header = rowToObject(wsSeti, 1); 
      const headerRow = wsSeti.getRow(1);
      const maxCol = wsSeti.columnCount || headerRow.cellCount || 0;
      let respColName = null;

      for (let c = 1; c <= maxCol; c++) {
        const name = String(cellToPlain(headerRow.getCell(c)) || "").trim();
        if (/responsibility|owner|assigned/i.test(name)) {
          respColName = name;
          break;
        }
      }

      if (respColName) {
        for (let r = 2; r <= last; r++) {
          const obj = rowToObject(wsSeti, r);
          const v = String(obj[respColName] || "").trim().toLowerCase();
          const idLower = String(employeeId).trim().toLowerCase();

          if (v === idLower || (employeeName && v.includes(employeeName))) {
            mySchedule.push(obj);
          }
        }
      }
    }

    const out = new ExcelJS.Workbook();
    out.created = new Date();

    const wsP = out.addWorksheet("Employee_Profile");
    wsP.addRow(Object.keys(employeeRow));
    wsP.addRow(Object.values(employeeRow));

    const wsE = out.addWorksheet("Engagements");
    wsE.addRow(myEngagements.length ? Object.keys(myEngagements[0]) : []);
    myEngagements.forEach((x) => wsE.addRow(Object.values(x)));

    const wsS = out.addWorksheet("Sites");
    wsS.addRow(mySites.length ? Object.keys(mySites[0]) : []);
    mySites.forEach((x) => wsS.addRow(Object.values(x)));

    const wsSch = out.addWorksheet("Schedule");
    wsSch.addRow(mySchedule.length ? Object.keys(mySchedule[0]) : []);
    mySchedule.forEach((x) => wsSch.addRow(Object.values(x)));

    return out.xlsx.writeBuffer();
  });
}

function backfillEngagementEmployeeIds(wsEng, wsProfiles) {
  const pIdx = headerIndexMapNormalized(wsProfiles);
  const pIdCol = pIdx.get("id");
  const pNameCol = pIdx.get("employee name");
  if (!pIdCol || !pNameCol) return false;

  const nameToId = new Map();
  const pLast = wsProfiles.actualRowCount || wsProfiles.rowCount || 1;
  for (let r = 2; r <= pLast; r++) {
    const row = wsProfiles.getRow(r);
    const id = text(cellToPlain(row.getCell(pIdCol)));
    const name = text(cellToPlain(row.getCell(pNameCol)));
    if (!id || !name) continue;
    const k = normKey(name);
    if (k && !nameToId.has(k)) nameToId.set(k, id);
  }

  if (!nameToId.size) return false;

  const eIdx = headerIndexMapNormalized(wsEng);
  const eEmpIdCol = eIdx.get("employee id");
  const eEmpNameCol = eIdx.get("employee name");
  if (!eEmpIdCol || !eEmpNameCol) return false;

  const last = wsEng.actualRowCount || wsEng.rowCount || 1;
  let changed = false;
  for (let r = 2; r <= last; r++) {
    const row = wsEng.getRow(r);
    const curId = text(cellToPlain(row.getCell(eEmpIdCol)));
    if (curId) continue;
    const name = text(cellToPlain(row.getCell(eEmpNameCol)));
    if (!name) continue;
    const id = nameToId.get(normKey(name));
    if (!id) continue;
    row.getCell(eEmpIdCol).value = id;
    row.commit?.();
    changed = true;
  }

  return changed;
}

function cellToPlain(cell) {
  const v = cell?.value;
  if (v == null) return "";
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    if (v.text != null) return v.text;
    if (v.result != null) return v.result;
    if (v.richText && Array.isArray(v.richText)) return v.richText.map((x) => x.text).join("");
    if (v.formula != null && v.value != null) return v.value;
  }
  return String(v);
}

function headerIndexMapNormalized(ws) {
  const map = new Map();
  const header = ws.getRow(1);
  const maxCol = ws.columnCount || header.cellCount || 0;

  for (let col = 1; col <= maxCol; col++) {
    const k = normKey(cellToPlain(header.getCell(col)));
    if (k && !map.has(k)) map.set(k, col);
  }

  return map;
}

function ensureColumnCanonical(ws, canonicalName, pos, aliases = []) {
  const header = ws.getRow(1);
  const maxCol = Math.max(ws.columnCount || 0, header.cellCount || 0);

  const want = normKey(canonicalName);
  const aliasSet = new Set([want, ...aliases.map(normKey)]);

  let foundCol = null;
  for (let c = 1; c <= maxCol; c++) {
    const k = normKey(cellToPlain(header.getCell(c)));
    if (aliasSet.has(k)) {
      foundCol = c;
      break;
    }
  }

  let changed = false;

  if (!foundCol) {
    ws.spliceColumns(pos, 0, [canonicalName]);
    changed = true;
    return changed;
  }

  const existingName = text(cellToPlain(header.getCell(foundCol)));
  if (normKey(existingName) !== want) {
    header.getCell(foundCol).value = canonicalName;
    changed = true;
  }

  if (foundCol !== pos) {
    const vals = [];
    for (let r = 1; r <= (ws.actualRowCount || ws.rowCount || 1); r++) {
      vals.push(ws.getRow(r).getCell(foundCol).value);
    }
    ws.spliceColumns(foundCol, 1);
    ws.spliceColumns(pos, 0, vals);
    changed = true;
  }

  return changed;
}

function ensureColumnAppend(ws, canonicalName, aliases = []) {
  const header = ws.getRow(1);
  const maxCol = Math.max(ws.columnCount || 0, header.cellCount || 0);

  const want = normKey(canonicalName);
  const aliasSet = new Set([want, ...aliases.map(normKey)]);

  for (let c = 1; c <= maxCol; c++) {
    const k = normKey(cellToPlain(header.getCell(c)));
    if (aliasSet.has(k)) {
      const existingName = text(cellToPlain(header.getCell(c)));
      if (normKey(existingName) !== want) header.getCell(c).value = canonicalName;
      return normKey(existingName) !== want;
    }
  }

  header.getCell(maxCol + 1).value = canonicalName;
  return true;
}

function getSheetOrThrow(wb, sheet) {
  const ws = wb.getWorksheet(sheet);
  if (!ws) throw new Error(`Sheet not found: ${sheet}`);
  return ws;
}

async function atomicSave(wb) {
  ensureDirs();
  const tmp = path.join(STORAGE_DIR, `data.${Date.now()}.tmp.xlsx`);
  await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, MASTER_PATH);
}

function createDefaultWorkbook() {
  const wb = new ExcelJS.Workbook();

  const profiles = wb.addWorksheet("profiles");
  profiles.addRow(["id", "Employee Name", "Designation"]);

  const engagements = wb.addWorksheet("Engagements");
  engagements.addRow([
    "Employee Name",
    "Site Engaged",
    "Starting Date",
    "End Date",
    "Duration (Days)",
    "status",
    "id",
    "Employee ID",
  ]);

  const sites = wb.addWorksheet("Sites_List");
  sites.addRow(["Location"]);

  const seti = wb.addWorksheet("Seti_Updated");
  seti.addRow(["id"]); 

  return wb;
}

function prepareSheet(ws, sheet) {
  let changed = false;

  if (sheet === "profiles") {
    changed = ensureColumnCanonical(ws, "id", 1, ["ID", "Id"]) || changed;
    changed =
      ensureColumnCanonical(ws, "Employee Name", 2, ["EmployeeName", "Name", "Employee  Name"]) ||
      changed;
    changed =
      ensureColumnCanonical(ws, "Designation", 3, ["Designations", "Position", "Role"]) || changed;
    return changed;
  }

  if (sheet === "Engagements") {
    changed = ensureColumnAppend(ws, "Employee Name", ["Emp Name", "EmployeeName", "Name"]) || changed;
    changed = ensureColumnAppend(ws, "Site Engaged", ["Site", "SiteEngaged", "Location"]) || changed;
    changed = ensureColumnAppend(ws, "Starting Date", ["Start Date", "StartDate"]) || changed;
    changed = ensureColumnAppend(ws, "End Date", ["Ending Date", "EndDate"]) || changed;
    changed = ensureColumnAppend(ws, "Duration (Days)", ["Duration", "Days", "Duration Days"]) || changed;
    changed = ensureColumnAppend(ws, "status", ["Status", "Phase"]) || changed;
    changed = ensureColumnAppend(ws, "id", ["ID"]) || changed;
    changed = ensureColumnAppend(ws, "Employee ID", ["Emp ID", "EmployeeID"]) || changed;
    return changed;
  }

  if (sheet === "Sites_List") {
    changed = ensureColumnCanonical(ws, "Location", 1, ["location", "Site", "Site Name"]) || changed;
    return changed;
  }

  if (sheet === "Seti_Updated") {
    changed = ensureColumnAppend(ws, "id", ["ID", "Id"]) || changed;
    return changed;
  }

  return changed;
}

let __bootPrepared = false;

async function prepareWorkbookOnBoot() {
  if (__bootPrepared) return;
  __bootPrepared = true;

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(MASTER_PATH);

  let changed = false;
  for (const name of ["profiles", "Engagements", "Sites_List", "Seti_Updated"]) {
    const ws = wb.getWorksheet(name);
    if (!ws) continue;
    changed = prepareSheet(ws, name) || changed;
  }

  if (changed) await atomicSave(wb);
}

export async function ensureMasterExists() {
  ensureDirs();
  const syncInfo = syncMasterFromUploads();

  if (!fs.existsSync(MASTER_PATH)) {
    const wb = await createDefaultWorkbook();
    await wb.xlsx.writeFile(MASTER_PATH);
  }

  await prepareWorkbookOnBoot();

  const stat = fs.statSync(MASTER_PATH);
  return { exists: true, bytes: stat.size, path: MASTER_PATH, ...syncInfo };
}

export async function getMasterPath() {
  await ensureMasterExists();
  return MASTER_PATH;
}

async function loadWorkbook() {
  await ensureMasterExists();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(MASTER_PATH);
  return wb;
}

function backfillIds(ws, prefix) {
  const idx = headerIndexMapNormalized(ws);
  const idCol = idx.get("id") || idx.get("ID".toLowerCase());
  if (!idCol) return false;

  const last = ws.actualRowCount || ws.rowCount || 1;
  let max = 0;

  for (let r = 2; r <= last; r++) {
    const val = text(cellToPlain(ws.getRow(r).getCell(idCol)));
    const m = val.match(new RegExp(`^${prefix}-(\\d+)$`, "i"));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }

  let changed = false;
  for (let r = 2; r <= last; r++) {
    const row = ws.getRow(r);
    const val = text(cellToPlain(row.getCell(idCol)));
    if (val) continue;
    max += 1;
    row.getCell(idCol).value = `${prefix}-${String(max).padStart(5, "0")}`;
    row.commit?.();
    changed = true;
  }

  return changed;
}

function findRowByLocation(ws, location) {
  const idx = headerIndexMapNormalized(ws);
  const locCol = idx.get("location");
  if (!locCol) return -1;

  const last = ws.actualRowCount || ws.rowCount || 1;
  const target = normKey(location);

  for (let r = 2; r <= last; r++) {
    const v = text(cellToPlain(ws.getRow(r).getCell(locCol)));
    if (normKey(v) === target) return r;
  }

  return -1;
}

function rowToObject(ws, rowNumber) {
  const headerRow = ws.getRow(1);
  const row = ws.getRow(rowNumber);
  const obj = {};

  const maxCol = ws.columnCount || headerRow.cellCount || 0;

  for (let col = 1; col <= maxCol; col++) {
    const key = text(cellToPlain(headerRow.getCell(col)));
    if (!key) continue;

    const value = text(cellToPlain(row.getCell(col)));

    if (obj[key] && text(obj[key]) && !text(value)) continue;
    if (!obj[key] || !text(obj[key])) obj[key] = value;
  }

  return obj;
}

function setRowFromObject(ws, rowNumber, patch) {
  const idx = headerIndexMapNormalized(ws);
  const row = ws.getRow(rowNumber);

  const incoming = new Map();
  for (const [k, v] of Object.entries(patch || {})) incoming.set(normKey(k), v);

  for (const [kLower, col] of idx.entries()) {
    if (!incoming.has(kLower)) continue;
    row.getCell(col).value = incoming.get(kLower) ?? "";
  }

  row.commit?.();
}

function mapDataToHeaderRow(ws, dataObj) {
  const incoming = new Map();
  for (const [k, v] of Object.entries(dataObj || {})) incoming.set(normKey(k), v);

  const header = ws.getRow(1);
  const maxCol = ws.columnCount || header.cellCount || 0;

  const vals = [];
  for (let col = 1; col <= maxCol; col++) {
    const k = normKey(cellToPlain(header.getCell(col)));
    vals.push(incoming.has(k) ? incoming.get(k) ?? "" : "");
  }
  return vals;
}

async function listRows({ sheet = "profiles", page = 1, pageSize = 50 } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();
    const ws = getSheetOrThrow(wb, sheet);

    let changed = false;
    changed = prepareSheet(ws, sheet) || changed;

    if (sheet === "profiles") changed = backfillIds(ws, "EMP") || changed;
    if (sheet === "Engagements") changed = backfillIds(ws, "ENG") || changed;
    if (sheet === "Seti_Updated") changed = backfillIds(ws, "SETI") || changed;

    if (sheet === "Engagements") {
      const wsProfiles = wb.getWorksheet("profiles");
      if (wsProfiles) changed = backfillEngagementEmployeeIds(ws, wsProfiles) || changed;
    }

    if (changed) await atomicSave(wb);

    const last = ws.actualRowCount || ws.rowCount || 1;
    const total = Math.max(0, last - 1);

    const p = Math.max(1, parseInt(page, 10) || 1);
    const ps = Math.min(500, Math.max(1, parseInt(pageSize, 10) || 50));

    const startRow = 2 + (p - 1) * ps;
    const endRow = Math.min(last, startRow + ps - 1);

    const items = [];
    for (let r = startRow; r <= endRow; r++) {
      const obj = rowToObject(ws, r);
      if (sheet === "Engagements") items.push(computeEngagementFields(obj));
      else items.push(obj);
    }

    return { ok: true, total, page: p, pageSize: ps, items };
  });
}

export async function createRow({ sheet = "profiles", idPrefix, data = {} } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();
    const ws = getSheetOrThrow(wb, sheet);

    const changed = prepareSheet(ws, sheet);
    if (changed) await atomicSave(wb);

    if (sheet === "Sites_List") {
      const location = text(data.Location ?? data.location);
      if (!location) throw new Error("Location is required");
      if (findRowByLocation(ws, location) !== -1) {
        throw new Error(`Duplicate Location: ${location}`);
      }
      ws.addRow(mapDataToHeaderRow(ws, { Location: location }));
      await atomicSave(wb);
      return { Location: location, id: location };
    }

    if (sheet === "profiles") idPrefix = "EMP";
    if (sheet === "Engagements") idPrefix = "ENG";
    if (sheet === "Seti_Updated") idPrefix = "SETI";

    const patch = { ...data };

    if (idPrefix) {
      const idx = headerIndexMapNormalized(ws);
      const idCol = idx.get("id");
      if (idCol) {
        const next = (() => {
          const last = ws.actualRowCount || ws.rowCount || 1;
          let max = 0;
          for (let r = 2; r <= last; r++) {
            const val = text(cellToPlain(ws.getRow(r).getCell(idCol)));
            const m = val.match(new RegExp(`^${idPrefix}-(\\d+)$`, "i"));
            if (m) max = Math.max(max, parseInt(m[1], 10));
          }
          return `${idPrefix}-${String(max + 1).padStart(5, "0")}`;
        })();
        patch.id = patch.id || next;
      }
    }

    ws.addRow(mapDataToHeaderRow(ws, patch));
    await atomicSave(wb);
    return patch;
  });
}

export async function updateRow({ sheet = "profiles", id, patch = {}, data } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();
    const ws = getSheetOrThrow(wb, sheet);

    prepareSheet(ws, sheet);

    const idx = headerIndexMapNormalized(ws);
    const idCol = idx.get("id");
    if (!idCol) throw new Error(`Sheet ${sheet} has no id column`);

    const last = ws.actualRowCount || ws.rowCount || 1;
    let rowNumber = -1;

    for (let r = 2; r <= last; r++) {
      const v = text(cellToPlain(ws.getRow(r).getCell(idCol)));
      if (v === id) {
        rowNumber = r;
        break;
      }
    }

    if (rowNumber === -1) throw new Error(`Row not found for id: ${id}`);

    const payload = data ?? patch ?? {};
    setRowFromObject(ws, rowNumber, payload);
    await atomicSave(wb);

    return rowToObject(ws, rowNumber);
  });
}

export async function deleteRow({ sheet = "profiles", id } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();
    const ws = getSheetOrThrow(wb, sheet);

    prepareSheet(ws, sheet);

    if (sheet === "Sites_List") {
      const location = text(id);
      const r = findRowByLocation(ws, location);
      if (r === -1) throw new Error(`Location not found: ${location}`);
      ws.spliceRows(r, 1);
      await atomicSave(wb);
      return { ok: true };
    }

    const idx = headerIndexMapNormalized(ws);
    const idCol = idx.get("id");
    if (!idCol) throw new Error(`Sheet ${sheet} has no id column`);

    const last = ws.actualRowCount || ws.rowCount || 1;
    let rowNumber = -1;

    for (let r = 2; r <= last; r++) {
      const v = text(cellToPlain(ws.getRow(r).getCell(idCol)));
      if (v === id) {
        rowNumber = r;
        break;
      }
    }

    if (rowNumber === -1) throw new Error(`Row not found for id: ${id}`);

    ws.spliceRows(rowNumber, 1);
    await atomicSave(wb);

    return { ok: true };
  });
}

export { listRows };