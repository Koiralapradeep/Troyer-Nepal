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
  return keys.find((k) => patterns.some((re) => re.test(String(k)))) || null;
}

function getAllObjects(ws, headerRowIndex = 1) {
  const last = ws.actualRowCount || ws.rowCount || 1;
  const items = [];

  for (let r = headerRowIndex + 1; r <= last; r++) {
    const obj = rowToObject(ws, r, headerRowIndex);
    const hasAny = Object.values(obj).some(
      (v) => String(v ?? "").trim() !== "",
    );
    if (!hasAny) continue;
    items.push(obj);
  }

  return items;
}

function normKey(v) {
  return String(v ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[_\-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
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

function getHeaderRowIndexBySheet(sheet, ws = null) {
  if (sheet === "Tools") return 2;

  const s = String(sheet || "")
    .trim()
    .toLowerCase();

  if (ws && (s.endsWith("_timeline") || s === "seti_updated")) {
    for (
      let r = 1;
      r <= Math.min(10, ws.actualRowCount || ws.rowCount || 10);
      r++
    ) {
      const row = ws.getRow(r);
      const vals = [];

      for (let c = 1; c <= Math.min(20, ws.columnCount || 20); c++) {
        vals.push(
          String(cellToPlain(row.getCell(c)) ?? "")
            .trim()
            .toLowerCase(),
        );
      }

      const joined = vals.join(" | ");

      const looksLikeScheduleHeader =
        joined.includes("task name") ||
        joined.includes("task name / milestone") ||
        joined.includes("start date") ||
        joined.includes("start") ||
        joined.includes("finish date") ||
        joined.includes("finish") ||
        joined.includes("duration") ||
        joined.includes("notes / section");

      if (looksLikeScheduleHeader) return r;
    }
  }

  return 1;
}

function dataStartRowBySheet(sheet) {
  return getHeaderRowIndexBySheet(sheet) + 1;
}

function toDateOnly(v) {
  if (!v) return null;
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function diffDays(start, finish) {
  const s = toDateOnly(start);
  const f = toDateOnly(finish);
  if (!s || !f) return "";
  return Math.max(0, Math.round((f - s) / 86400000));
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
    const masterStat = fs.existsSync(MASTER_PATH)
      ? fs.statSync(MASTER_PATH)
      : null;
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

  const num = Number(s);
  if (!isNaN(num) && num > 20000 && num < 60000) {
    const base = new Date(Date.UTC(1899, 11, 30));
    const dt = new Date(base.getTime() + num * 86400000);
    return startOfDay(dt);
  }

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
  const duration = start
    ? end
      ? daysInclusive(start, end)
      : daysInclusive(start, today)
    : "";

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

export async function createScheduleProjectSheet({ sheetName, title } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();

    if (wb.getWorksheet(sheetName)) {
      throw new Error(`Sheet already exists: ${sheetName}`);
    }

    const ws = wb.addWorksheet(sheetName);

    ws.addRow([
      title || sheetName.replace(/_Timeline$/i, "").replace(/_/g, " "),
    ]);
    ws.addRow([`Sheet: ${sheetName}`]);
    ws.addRow([
      "id",
      "Task Name",
      "% Completion",
      "Duration",
      "Start Date",
      "End Date",
      "Remarks",
    ]);

    ws.getRow(1).font = { bold: true, size: 16 };
    ws.getRow(2).font = { italic: true, color: { argb: "666666" } };
    ws.getRow(3).font = { bold: true };

    ws.columns = [
      { width: 12 },
      { width: 42 },
      { width: 14 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 28 },
    ];

    prepareScheduleProjectSheet(ws);

    await atomicSave(wb);

    return {
      ok: true,
      sheet: sheetName,
      title: title || sheetName,
    };
  });
}
export async function renameScheduleProjectSheet({
  oldSheetName,
  newTitle,
} = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();

    const oldName = String(oldSheetName || "").trim();
    const title = String(newTitle || "").trim();

    if (!oldName) throw new Error("Old sheet name is required");
    if (!title) throw new Error("Project name is required");

    const ws = wb.getWorksheet(oldName);
    if (!ws) throw new Error(`Sheet not found: ${oldName}`);

    const safeBase = title
      .replace(/[^\w\s-]/g, "")
      .trim()
      .replace(/\s+/g, "_");
    const newSheetName = `${safeBase}_Timeline`;

    if (oldName !== newSheetName && wb.getWorksheet(newSheetName)) {
      throw new Error(`A project already exists with name: ${title}`);
    }

    ws.name = newSheetName;

    await atomicSave(wb);

    return {
      ok: true,
      oldSheet: oldName,
      sheet: newSheetName,
      title,
    };
  });
}

export async function deleteScheduleProjectSheet({ sheetName } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();

    const target = String(sheetName || "").trim();
    if (!target) throw new Error("Sheet name is required");

    const ws = wb.getWorksheet(target);
    if (!ws) throw new Error(`Sheet not found: ${target}`);

    wb.removeWorksheet(ws.id);

    await atomicSave(wb);

    return {
      ok: true,
      sheet: target,
    };
  });
}

export async function exportEmployeeWorkbookBuffer(employeeId) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();

    const wsProfiles = getSheetOrThrow(wb, "profiles");
    const wsEng = getSheetOrThrow(wb, "Engagements");
    const wsSites = getSheetOrThrow(wb, "Sites_List");
    const wsSeti = wb.getWorksheet("Seti_Updated");

    prepareSheet(wsProfiles, "profiles");
    prepareSheet(wsEng, "Engagements");
    prepareSheet(wsSites, "Sites_List");
    if (wsSeti) prepareSheet(wsSeti, "Seti_Updated");

    const pLast = wsProfiles.actualRowCount || wsProfiles.rowCount || 1;
    let employeeRow = null;

    for (let r = 2; r <= pLast; r++) {
      const obj = rowToObject(wsProfiles, r, 1);
      if (
        String(obj.id || "")
          .trim()
          .toLowerCase() === String(employeeId).trim().toLowerCase()
      ) {
        employeeRow = obj;
        if (String(obj["Duration (days)"] ?? "").trim() === "") {
          obj["Duration (days)"] = diffDays(obj["Start"], obj["Finish"]);
        }
        break;
      }
    }

    if (!employeeRow) throw new Error(`Employee not found: ${employeeId}`);

    const employeeName = String(employeeRow["Employee Name"] || "")
      .trim()
      .toLowerCase();

    const eLast = wsEng.actualRowCount || wsEng.rowCount || 1;
    const myEngagements = [];

    for (let r = 2; r <= eLast; r++) {
      const obj = computeEngagementFields(rowToObject(wsEng, r, 1));
      const eid = String(obj["Employee ID"] || "")
        .trim()
        .toLowerCase();
      if (eid === String(employeeId).trim().toLowerCase()) {
        myEngagements.push(obj);
      }
    }

    const siteSet = new Set(
      myEngagements
        .map((e) =>
          String(e["Site Engaged"] || "")
            .trim()
            .toLowerCase(),
        )
        .filter(Boolean),
    );

    const sLast = wsSites.actualRowCount || wsSites.rowCount || 1;
    const mySites = [];
    for (let r = 2; r <= sLast; r++) {
      const obj = rowToObject(wsSites, r, 1);
      const loc = String(obj["Location"] || "")
        .trim()
        .toLowerCase();
      if (loc && siteSet.has(loc)) mySites.push(obj);
    }

    const mySchedule = [];
    if (wsSeti) {
      const last = wsSeti.actualRowCount || wsSeti.rowCount || 1;
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
          const obj = rowToObject(wsSeti, r, 1);
          const v = String(obj[respColName] || "")
            .trim()
            .toLowerCase();
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
  const pIdx = headerIndexMapNormalized(wsProfiles, 1);
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

  const eIdx = headerIndexMapNormalized(wsEng, 1);
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
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean")
    return v;
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    if (v.text != null) return v.text;
    if (v.result != null) return v.result;
    if (v.richText && Array.isArray(v.richText))
      return v.richText.map((x) => x.text).join("");
    if (v.formula != null && v.value != null) return v.value;
  }
  return String(v);
}

function headerIndexMapNormalized(ws, headerRowIndex = 1) {
  const map = new Map();
  const header = ws.getRow(headerRowIndex);
  const maxCol = ws.columnCount || header.cellCount || 0;

  for (let col = 1; col <= maxCol; col++) {
    const k = normKey(cellToPlain(header.getCell(col)));
    if (k && !map.has(k)) map.set(k, col);
  }

  return map;
}

function ensureColumnCanonical(
  ws,
  canonicalName,
  pos,
  aliases = [],
  headerRowIndex = 1,
) {
  const header = ws.getRow(headerRowIndex);
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
    const values = [];
    const totalRows = ws.actualRowCount || ws.rowCount || 1;
    for (let r = 1; r <= totalRows; r++) {
      values.push(r === headerRowIndex ? canonicalName : "");
    }
    ws.spliceColumns(pos, 0, values);
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

function ensureColumnAppend(
  ws,
  canonicalName,
  aliases = [],
  headerRowIndex = 1,
) {
  const header = ws.getRow(headerRowIndex);
  const maxCol = Math.max(ws.columnCount || 0, header.cellCount || 0);

  const want = normKey(canonicalName);
  const aliasSet = new Set([want, ...aliases.map(normKey)]);

  for (let c = 1; c <= maxCol; c++) {
    const k = normKey(cellToPlain(header.getCell(c)));
    if (aliasSet.has(k)) {
      const existingName = text(cellToPlain(header.getCell(c)));
      if (normKey(existingName) !== want) {
        header.getCell(c).value = canonicalName;
        return true;
      }
      return false;
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
  sites.addRow(["Location", "Status"]);

  const seti = wb.addWorksheet("Seti_Updated");
  seti.addRow(["id"]);

  const tools = wb.addWorksheet("Tools");
  tools.addRow(["TOOLS"]);
  tools.addRow(["S.N", "Item", "UoM", "Qty", "Location", "Remarks"]);

  return wb;
}

function prepareSheet(ws, sheet) {
  let changed = false;

  if (sheet === "profiles") {
    changed = ensureColumnCanonical(ws, "id", 1, ["ID", "Id"], 1) || changed;
    changed =
      ensureColumnCanonical(
        ws,
        "Employee Name",
        2,
        ["EmployeeName", "Name", "Employee  Name"],
        1,
      ) || changed;
    changed =
      ensureColumnCanonical(
        ws,
        "Designation",
        3,
        ["Designations", "Position", "Role"],
        1,
      ) || changed;
    return changed;
  }

  if (sheet === "Engagements") {
    changed =
      ensureColumnAppend(
        ws,
        "Employee Name",
        ["Emp Name", "EmployeeName", "Name"],
        1,
      ) || changed;
    changed =
      ensureColumnAppend(
        ws,
        "Site Engaged",
        ["Site", "SiteEngaged", "Location"],
        1,
      ) || changed;
    changed =
      ensureColumnAppend(ws, "Starting Date", ["Start Date", "StartDate"], 1) ||
      changed;
    changed =
      ensureColumnAppend(ws, "End Date", ["Ending Date", "EndDate"], 1) ||
      changed;
    changed =
      ensureColumnAppend(
        ws,
        "Duration (Days)",
        ["Duration", "Days", "Duration Days"],
        1,
      ) || changed;
    changed =
      ensureColumnAppend(ws, "status", ["Status", "Phase"], 1) || changed;
    changed = ensureColumnAppend(ws, "id", ["ID"], 1) || changed;
    changed =
      ensureColumnAppend(ws, "Employee ID", ["Emp ID", "EmployeeID"], 1) ||
      changed;
    return changed;
  }

  if (sheet === "Sites_List") {
    changed =
      ensureColumnCanonical(
        ws,
        "Location",
        1,
        ["location", "Site", "Site Name"],
        1,
      ) || changed;
    changed =
      ensureColumnAppend(
        ws,
        "Status",
        ["status", "Phase", "phase", "Site Status"],
        1,
      ) || changed;
    return changed;
  }

  if (sheet === "Seti_Updated") {
    changed = ensureColumnCanonical(ws, "id", 1, ["ID", "Id"], 1) || changed;
    return changed;
  }

  if (sheet === "Tools") {
    const headerRowIndex = 2;

    changed =
      ensureColumnCanonical(
        ws,
        "S.N",
        1,
        [
          "SN",
          "S N",
          "S.No",
          "S. No",
          "Serial No",
          "Serial Number",
          "id",
          "ID",
        ],
        headerRowIndex,
      ) || changed;

    changed =
      ensureColumnCanonical(
        ws,
        "Item",
        2,
        ["Tool", "Tools", "Name"],
        headerRowIndex,
      ) || changed;

    changed =
      ensureColumnCanonical(
        ws,
        "UoM",
        3,
        ["UOM", "Unit", "Unit of Measure"],
        headerRowIndex,
      ) || changed;

    changed =
      ensureColumnCanonical(
        ws,
        "Qty",
        4,
        ["QTY", "Quantity"],
        headerRowIndex,
      ) || changed;

    changed =
      ensureColumnCanonical(
        ws,
        "Location",
        5,
        ["location", "Site"],
        headerRowIndex,
      ) || changed;

    changed =
      ensureColumnCanonical(
        ws,
        "Remarks",
        6,
        ["Remark", "Comments", "Comment", "Notes", "Note"],
        headerRowIndex,
      ) || changed;

    changed = ensureColumnAppend(ws, "__rowKey", [], headerRowIndex) || changed;
    changed =
      ensureColumnAppend(ws, "__rowType", [], headerRowIndex) || changed;
    changed =
      ensureColumnAppend(ws, "__parentSn", [], headerRowIndex) || changed;

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

  for (const ws of wb.worksheets) {
    const name = ws.name;

    if (
      name === "profiles" ||
      name === "Engagements" ||
      name === "Sites_List" ||
      name === "Tools"
    ) {
      changed = prepareSheet(ws, name) || changed;
      continue;
    }

    if (isScheduleProjectSheetName(name)) {
      changed = prepareScheduleProjectSheet(ws) || changed;
    }
  }

  if (changed) await atomicSave(wb);
}

export async function ensureMasterExists() {
  ensureDirs();

  let syncInfo = { synced: false };

  if (!fs.existsSync(MASTER_PATH)) {
    syncInfo = syncMasterFromUploads();

    if (!fs.existsSync(MASTER_PATH)) {
      const wb = createDefaultWorkbook();
      await wb.xlsx.writeFile(MASTER_PATH);
    }
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

function backfillIds(ws, prefix, headerRowIndex = 1) {
  const idx = headerIndexMapNormalized(ws, headerRowIndex);
  const idCol = idx.get("id");
  if (!idCol) return false;

  const startRow = headerRowIndex + 1;
  const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
  let max = 0;

  for (let r = startRow; r <= last; r++) {
    const val = text(cellToPlain(ws.getRow(r).getCell(idCol)));
    const m = val.match(new RegExp(`^${prefix}-(\\d+)$`, "i"));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }

  let changed = false;
  for (let r = startRow; r <= last; r++) {
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

function backfillToolRowKeys(ws) {
  const headerRowIndex = 2;
  const idx = headerIndexMapNormalized(ws, headerRowIndex);
  const rowKeyCol = idx.get("rowkey");

  if (!rowKeyCol) return false;

  const itemCol = idx.get("item");
  const uomCol = idx.get("uom");
  const qtyCol = idx.get("qty");
  const locCol = idx.get("location");
  const remarksCol = idx.get("remarks");
  const snCol = idx.get("s.n");

  const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
  let max = 0;

  for (let r = headerRowIndex + 1; r <= last; r++) {
    const raw = text(cellToPlain(ws.getRow(r).getCell(rowKeyCol)));
    const m = raw.match(/^TOOLROW-(\d+)$/i);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }

  let changed = false;

  for (let r = headerRowIndex + 1; r <= last; r++) {
    const row = ws.getRow(r);
    const current = text(cellToPlain(row.getCell(rowKeyCol)));

    const hasAny =
      text(cellToPlain(row.getCell(snCol))) ||
      text(cellToPlain(row.getCell(itemCol))) ||
      text(cellToPlain(row.getCell(uomCol))) ||
      text(cellToPlain(row.getCell(qtyCol))) ||
      text(cellToPlain(row.getCell(locCol))) ||
      text(cellToPlain(row.getCell(remarksCol)));

    if (!hasAny) continue;
    if (current) continue;

    max += 1;
    row.getCell(rowKeyCol).value = `TOOLROW-${String(max).padStart(6, "0")}`;
    row.commit?.();
    changed = true;
  }

  return changed;
}

function backfillToolsIds(ws) {
  const headerRowIndex = 2;
  const idx = headerIndexMapNormalized(ws, headerRowIndex);
  const snCol = idx.get("s.n");
  if (!snCol) return false;

  const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
  let seq = 0;
  let changed = false;

  for (let r = headerRowIndex + 1; r <= last; r++) {
    const row = ws.getRow(r);

    const currentSN = text(cellToPlain(row.getCell(snCol)));
    const itemCol = idx.get("item") || 2;
    const qtyCol = idx.get("qty") || 4;
    const locCol = idx.get("location") || 5;
    const remarksCol = idx.get("remarks") || 6;

    const item = text(cellToPlain(row.getCell(itemCol)));
    const qty = text(cellToPlain(row.getCell(qtyCol)));
    const loc = text(cellToPlain(row.getCell(locCol)));
    const remarks = text(cellToPlain(row.getCell(remarksCol)));

    const hasAny = item || qty || loc || remarks;
    if (!hasAny) continue;

    // child row: must stay blank
    if (!currentSN) continue;

    seq += 1;

    if (currentSN !== String(seq)) {
      row.getCell(snCol).value = String(seq);
      row.commit?.();
      changed = true;
    }
  }

  return changed;
}

function findRowByLocation(ws, location) {
  const idx = headerIndexMapNormalized(ws, 1);
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

function findRowByToolsId(ws, id) {
  const headerRowIndex = 2;
  const idx = headerIndexMapNormalized(ws, headerRowIndex);
  const idCol = idx.get("s.n");
  if (!idCol) return -1;

  const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
  for (let r = 3; r <= last; r++) {
    const v = text(cellToPlain(ws.getRow(r).getCell(idCol)));
    if (v === String(id)) return r;
  }

  return -1;
}

function rowToObject(ws, rowNumber, headerRowIndex = 1) {
  const headerRow = ws.getRow(headerRowIndex);
  const row = ws.getRow(rowNumber);

  const obj = {};
  const maxCol = ws.columnCount || headerRow.cellCount || 0;

  for (let col = 1; col <= maxCol; col++) {
    const keyRaw = cellToPlain(headerRow.getCell(col));
    const key = String(keyRaw ?? "").trim();
    if (!key) continue;

    let v = cellToPlain(row.getCell(col));

    if (v == null) v = "";
    if (v instanceof Date) {
      v = toISODate(v);
    } else if (typeof v === "string") {
      v = v.trim();
    }

    obj[key] = v;
  }

  return obj;
}

function isScheduleProjectSheetName(sheet) {
  const s = String(sheet || "")
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

function prepareScheduleProjectSheet(ws) {
  let changed = false;
  const headerRowIndex = getHeaderRowIndexBySheet(ws.name, ws);

  changed =
    ensureColumnAppend(ws, "id", ["ID", "Id"], headerRowIndex) || changed;

  changed =
    ensureColumnAppend(
      ws,
      "Task Name",
      [
        "TASK NAME",
        "Task",
        "Activity",
        "Activities",
        "Work Item",
        "Name",
        "Task Name / Milestone",
        "Milestone",
      ],
      headerRowIndex,
    ) || changed;

  changed =
    ensureColumnAppend(
      ws,
      "% Completion",
      [
        "% Complete",
        "% complete",
        "Completion",
        "Progress",
        "Progress %",
        "Percent Complete",
      ],
      headerRowIndex,
    ) || changed;

  changed =
    ensureColumnAppend(
      ws,
      "Duration",
      [
        "Duration (days)",
        "DURATION",
        "Days",
        "No. of Days",
        "Planned Duration",
      ],
      headerRowIndex,
    ) || changed;

  changed =
    ensureColumnAppend(
      ws,
      "Start Date",
      [
        "Start",
        "START",
        "Starting Date",
        "Planned Start",
        "Begin Date",
        "Begin",
      ],
      headerRowIndex,
    ) || changed;

  changed =
    ensureColumnAppend(
      ws,
      "End Date",
      [
        "Finish",
        "FINISH",
        "End",
        "END",
        "Finish Date",
        "FINISH DATE",
        "Completion Date",
        "Planned Finish",
      ],
      headerRowIndex,
    ) || changed;

  changed =
    ensureColumnAppend(
      ws,
      "Remarks",
      [
        "Remark",
        "Notes",
        "Comments",
        "Comment",
        "Description",
        "Notes / Section",
      ],
      headerRowIndex,
    ) || changed;

  return changed;
}

function findHeaderRowIndex(ws) {
  const want = ["ID", "Task Name", "Start", "Finish"];
  const maxScan = Math.min(25, ws.rowCount || 25);

  for (let r = 1; r <= maxScan; r++) {
    const row = ws.getRow(r);
    const values = [];
    for (let c = 1; c <= Math.min(30, ws.columnCount || 30); c++) {
      const v = row.getCell(c).value;
      if (v == null) continue;
      values.push(String(v).trim());
    }

    const hit = want.every((k) =>
      values.some((x) => x.toLowerCase() === k.toLowerCase()),
    );
    if (hit) return r;
  }

  return 1;
}

function setRowFromObject(ws, rowNumber, patch, headerRowIndex = 1) {
  const idx = headerIndexMapNormalized(ws, headerRowIndex);
  const row = ws.getRow(rowNumber);

  const incoming = new Map();
  for (const [k, v] of Object.entries(patch || {})) incoming.set(normKey(k), v);

  for (const [kLower, col] of idx.entries()) {
    if (!incoming.has(kLower)) continue;
    row.getCell(col).value = incoming.get(kLower) ?? "";
  }

  row.commit?.();
}

function mapDataToHeaderRow(ws, dataObj, headerRowIndex = 1) {
  const incoming = new Map();
  for (const [k, v] of Object.entries(dataObj || {})) {
    incoming.set(normKey(k), v);
  }

  const header = ws.getRow(headerRowIndex);
  const maxCol = ws.columnCount || header.cellCount || 0;

  const vals = [];
  for (let col = 1; col <= maxCol; col++) {
    const k = normKey(cellToPlain(header.getCell(col)));
    vals.push(incoming.has(k) ? (incoming.get(k) ?? "") : "");
  }
  return vals;
}
function makeToolRowKey(ws) {
  const headerRowIndex = 2;
  const idx = headerIndexMapNormalized(ws, headerRowIndex);
  const keyCol = idx.get("rowkey");
  if (!keyCol) return `TOOLROW-${Date.now()}`;

  const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
  let max = 0;

  for (let r = headerRowIndex + 1; r <= last; r++) {
    const raw = text(cellToPlain(ws.getRow(r).getCell(keyCol)));
    const m = raw.match(/^TOOLROW-(\d+)$/i);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }

  return `TOOLROW-${String(max + 1).padStart(6, "0")}`;
}

function findRowByToolRowKey(ws, rowKey) {
  const headerRowIndex = 2;
  const idx = headerIndexMapNormalized(ws, headerRowIndex);
  const keyCol = idx.get("rowkey");
  if (!keyCol) return -1;

  const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
  for (let r = headerRowIndex + 1; r <= last; r++) {
    const v = text(cellToPlain(ws.getRow(r).getCell(keyCol)));
    if (v === String(rowKey)) return r;
  }

  return -1;
}

async function listRows({ sheet = "profiles", page = 1, pageSize = 50 } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();
    const ws = getSheetOrThrow(wb, sheet);

    let changed = false;
    changed = prepareSheet(ws, sheet) || changed;

    if (sheet === "profiles") changed = backfillIds(ws, "EMP", 1) || changed;
    if (sheet === "Engagements") changed = backfillIds(ws, "ENG", 1) || changed;
    if (sheet === "Tools") {
      changed = backfillToolsIds(ws) || changed;
      changed = backfillToolRowKeys(ws) || changed;
    }

    if (isScheduleProjectSheetName(sheet)) {
      changed = prepareScheduleProjectSheet(ws) || changed;

      const headerRowIndex = getHeaderRowIndexBySheet(sheet, ws);
      const idx = headerIndexMapNormalized(ws, headerRowIndex);
      const idCol = idx.get("id");

      if (!idCol) throw new Error(`Sheet ${sheet} has no id column`);

      const dataStartRow = headerRowIndex + 1;
      const last = ws.actualRowCount || ws.rowCount || headerRowIndex;

      let maxId = 0;
      for (let r = dataStartRow; r <= last; r++) {
        const raw = text(cellToPlain(ws.getRow(r).getCell(idCol)));
        const n = parseInt(String(raw).replace(/[^\d]/g, ""), 10);
        if (!isNaN(n)) maxId = Math.max(maxId, n);
      }

      for (let r = dataStartRow; r <= last; r++) {
        const rowObj = rowToObject(ws, r, headerRowIndex);
        const hasAny = Object.values(rowObj).some(
          (v) => String(v ?? "").trim() !== "",
        );
        if (!hasAny) continue;

        const currentId = text(cellToPlain(ws.getRow(r).getCell(idCol)));
        if (!currentId) {
          maxId += 1;
          ws.getRow(r).getCell(idCol).value = String(maxId);
          ws.getRow(r).commit?.();
          changed = true;
        }
      }

      if (changed) await atomicSave(wb);

      const total = Math.max(0, last - headerRowIndex);
      const p = Math.max(1, parseInt(page, 10) || 1);
      const ps = Math.min(500, Math.max(1, parseInt(pageSize, 10) || 50));

      const startRow = dataStartRow + (p - 1) * ps;
      const endRow = Math.min(last, startRow + ps - 1);

      const items = [];
      for (let r = startRow; r <= endRow; r++) {
        const obj = rowToObject(ws, r, headerRowIndex);
        const hasAny = Object.values(obj).some(
          (v) => String(v ?? "").trim() !== "",
        );
        if (!hasAny) continue;
        items.push(obj);
      }
      return { ok: true, total, page: p, pageSize: ps, items };
    }

    if (sheet === "Engagements") {
      const wsProfiles = wb.getWorksheet("profiles");
      if (wsProfiles) {
        changed = backfillEngagementEmployeeIds(ws, wsProfiles) || changed;
      }
    }

    if (changed) await atomicSave(wb);

    const headerRowIndex = getHeaderRowIndexBySheet(sheet, ws);
    const dataStartRow = headerRowIndex + 1;
    const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
    const total = Math.max(0, last - headerRowIndex);

    const p = Math.max(1, parseInt(page, 10) || 1);
    const ps = Math.min(500, Math.max(1, parseInt(pageSize, 10) || 50));

    const startRow = dataStartRow + (p - 1) * ps;
    const endRow = Math.min(last, startRow + ps - 1);

    const items = [];
    for (let r = startRow; r <= endRow; r++) {
      const obj = rowToObject(ws, r, headerRowIndex);

      const hasAny = Object.values(obj).some(
        (v) => String(v ?? "").trim() !== "",
      );
      if (!hasAny) continue;

      if (sheet === "Engagements") {
        items.push(computeEngagementFields(obj));
      } else if (sheet === "Sites_List") {
        const out = { ...obj };
        for (const key of Object.keys(out)) {
          if (/status|phase/i.test(key)) {
            out[key] = normalizeSiteStatus(out[key]);
          }
        }
        out.id = out.Location;
        items.push(out);
      } else if (sheet === "Tools") {
        const out = { ...obj };

        const sn = String(out["S.N"] ?? "").trim();
        const rowTypeRaw = String(out["__rowType"] ?? "")
          .trim()
          .toLowerCase();
        const parentSn = String(out["__parentSn"] ?? "").trim();
        const rowKey = String(out["__rowKey"] ?? "").trim();

        out.__rowType = rowTypeRaw || (sn ? "parent" : "child");
        out.__parentSn = parentSn;
        out.__rowKey = rowKey || `TOOLS-TEMP-${r}`;
        out.id = out.__rowKey;

        items.push(out);
      } else {
        items.push(obj);
      }
    }

    return { ok: true, total, page: p, pageSize: ps, items };
  });
}

export async function createRow({
  sheet = "profiles",
  idPrefix,
  data = {},
} = {}) {
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

      const sitePatch = { ...data, Location: location };

      for (const key of Object.keys(sitePatch)) {
        if (/status|phase/i.test(key)) {
          sitePatch[key] = normalizeSiteStatus(sitePatch[key]);
        }
      }

      ws.addRow(mapDataToHeaderRow(ws, sitePatch, 1));
      await atomicSave(wb);
      return { ...sitePatch, id: location };
    }
    if (sheet === "Tools") {
      const headerRowIndex = 2;
      const patch = { ...data };

      let changedTools = false;
      changedTools = prepareSheet(ws, sheet) || changedTools;
      changedTools = backfillToolRowKeys(ws) || changedTools;
      changedTools = backfillToolsIds(ws) || changedTools;

      const idx = headerIndexMapNormalized(ws, headerRowIndex);
      const snCol = idx.get("s.n");

      const isChild =
        String(patch.__isChild ?? "")
          .trim()
          .toLowerCase() === "true";

      const parentSn = text(patch.__parentSn);

      delete patch.__isChild;
      delete patch.id;
      delete patch.ID;
      delete patch.Id;
      delete patch["S.N"];
      delete patch["s.n"];

      if (isChild && !parentSn) {
        throw new Error("Parent S.N is required for child row");
      }

      let nextSN = 1;
      if (!isChild && snCol) {
        const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
        for (let r = headerRowIndex + 1; r <= last; r++) {
          const currentSN = text(cellToPlain(ws.getRow(r).getCell(snCol)));
          const n = parseInt(currentSN, 10);
          if (!isNaN(n)) nextSN = Math.max(nextSN, n + 1);
        }
      }

      patch["S.N"] = isChild ? "" : String(nextSN);
      patch["__rowKey"] = makeToolRowKey(ws);
      patch["__rowType"] = isChild ? "child" : "parent";
      patch["__parentSn"] = isChild ? parentSn : "";

      const rowValues = mapDataToHeaderRow(ws, patch, headerRowIndex);

      if (isChild) {
        const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
        let insertAt = last + 1;
        let parentRow = -1;

        for (let r = headerRowIndex + 1; r <= last; r++) {
          const currentSN = text(cellToPlain(ws.getRow(r).getCell(snCol)));
          if (currentSN === parentSn) {
            parentRow = r;
            insertAt = r + 1;
            continue;
          }
          if (parentRow !== -1 && currentSN) {
            insertAt = r;
            break;
          }
          if (parentRow !== -1) {
            insertAt = r + 1;
          }
        }

        ws.spliceRows(insertAt, 0, rowValues);
      } else {
        ws.addRow(rowValues);
      }

      backfillToolRowKeys(ws);
      backfillToolsIds(ws);

      await atomicSave(wb);

      return {
        ...patch,
        id: patch["__rowKey"],
      };
    }

    if (sheet === "profiles") idPrefix = "EMP";
    if (sheet === "Engagements") idPrefix = "ENG";

    const patch = { ...data };

    if (isScheduleProjectSheetName(sheet)) {
      let changedSchedule = false;
      changedSchedule = prepareScheduleProjectSheet(ws) || changedSchedule;
      if (changedSchedule) await atomicSave(wb);

      const headerRowIndex = getHeaderRowIndexBySheet(sheet, ws);
      const idx = headerIndexMapNormalized(ws, headerRowIndex);
      const idCol = idx.get("id");

      const patch = { ...data };
      if (idCol && !patch.id) {
        const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
        let max = 0;

        for (let r = headerRowIndex + 1; r <= last; r++) {
          const raw = text(cellToPlain(ws.getRow(r).getCell(idCol)));
          const n = parseInt(String(raw).replace(/[^\d]/g, ""), 10);
          if (!isNaN(n)) max = Math.max(max, n);
        }

        patch.id = String(max + 1);
      }

      ws.addRow(mapDataToHeaderRow(ws, patch, headerRowIndex));
      await atomicSave(wb);
      return patch;
    }

    if (idPrefix) {
      const idx = headerIndexMapNormalized(ws, 1);
      const idCol = idx.get("id");
      if (idCol) {
        const last = ws.actualRowCount || ws.rowCount || 1;
        let max = 0;
        for (let r = 2; r <= last; r++) {
          const val = text(cellToPlain(ws.getRow(r).getCell(idCol)));
          const m = val.match(new RegExp(`^${idPrefix}-(\\d+)$`, "i"));
          if (m) max = Math.max(max, parseInt(m[1], 10));
        }
        const next = `${idPrefix}-${String(max + 1).padStart(5, "0")}`;
        patch.id = patch.id || next;
      }
    }

    ws.addRow(mapDataToHeaderRow(ws, patch, 1));
    await atomicSave(wb);
    return patch;
  });
}

export async function updateRow({
  sheet = "profiles",
  id,
  key = "id",
  patch = {},
  data,
} = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();
    const ws = getSheetOrThrow(wb, sheet);

    prepareSheet(ws, sheet);

    const payload = { ...(data ?? patch ?? {}) };

    if (sheet === "Sites_List") {
      const idx = headerIndexMapNormalized(ws, 1);
      const locCol = idx.get("location");
      if (!locCol) throw new Error(`Sheet ${sheet} has no Location column`);

      const last = ws.actualRowCount || ws.rowCount || 1;
      let rowNumber = -1;

      const target = normKey(id);
      for (let r = 2; r <= last; r++) {
        const v = text(cellToPlain(ws.getRow(r).getCell(locCol)));
        if (normKey(v) === target) {
          rowNumber = r;
          break;
        }
      }

      if (rowNumber === -1) throw new Error(`Location not found: ${id}`);

      if (payload.Location && normKey(payload.Location) !== target) {
        const newLoc = text(payload.Location);
        if (!newLoc) throw new Error("Location is required");
        if (findRowByLocation(ws, newLoc) !== -1) {
          throw new Error(`Duplicate Location: ${newLoc}`);
        }
      }

      for (const k of Object.keys(payload)) {
        if (/status|phase/i.test(k)) {
          payload[k] = normalizeSiteStatus(payload[k]);
        }
      }

      setRowFromObject(ws, rowNumber, payload, 1);
      await atomicSave(wb);

      const out = rowToObject(ws, rowNumber, 1);
      out.id = out.Location;
      return out;
    }

    if (sheet === "Tools") {
      const headerRowIndex = 2;

      let changedTools = false;
      changedTools = prepareSheet(ws, sheet) || changedTools;
      changedTools = backfillToolRowKeys(ws) || changedTools;
      changedTools = backfillToolsIds(ws) || changedTools;

      const rowNumber = findRowByToolRowKey(ws, id);
      if (rowNumber === -1) {
        throw new Error(`Row not found for tool row key: ${id}`);
      }

      const idx = headerIndexMapNormalized(ws, headerRowIndex);
      const snCol = idx.get("s.n");

      const existing = rowToObject(ws, rowNumber, headerRowIndex);
      const existingType = String(existing["__rowType"] ?? "")
        .trim()
        .toLowerCase();

      const isChild =
        String(payload.__isChild ?? payload.__rowType ?? existingType)
          .trim()
          .toLowerCase() === "true" ||
        String(payload.__rowType ?? existingType)
          .trim()
          .toLowerCase() === "child";

      const parentSn = text(payload.__parentSn || existing["__parentSn"]);

      delete payload.id;
      delete payload["id"];
      delete payload["S.N"];
      delete payload["s.n"];
      delete payload["__isChild"];

      if (isChild) {
        if (!parentSn) throw new Error("Parent item is required for child row");
        payload["__rowType"] = "child";
        payload["__parentSn"] = parentSn;
        payload["S.N"] = "";
      } else {
        payload["__rowType"] = "parent";
        payload["__parentSn"] = "";

        const currentSN = text(
          cellToPlain(ws.getRow(rowNumber).getCell(snCol)),
        );
        if (currentSN) {
          payload["S.N"] = currentSN;
        } else {
          let nextSN = 1;
          const last = ws.actualRowCount || ws.rowCount || headerRowIndex;

          for (let r = headerRowIndex + 1; r <= last; r++) {
            if (r === rowNumber) continue;
            const v = text(cellToPlain(ws.getRow(r).getCell(snCol)));
            const n = parseInt(v, 10);
            if (!isNaN(n)) nextSN = Math.max(nextSN, n + 1);
          }

          payload["S.N"] = String(nextSN);
        }
      }

      setRowFromObject(ws, rowNumber, payload, headerRowIndex);

      changedTools = backfillToolsIds(ws) || changedTools;
      if (changedTools) await atomicSave(wb);
      else await atomicSave(wb);

      const out = rowToObject(ws, rowNumber, headerRowIndex);
      out.id = String(out["__rowKey"] ?? "");
      return out;
    }

    if (isScheduleProjectSheetName(sheet)) {
      prepareScheduleProjectSheet(ws);

      const headerRowIndex = getHeaderRowIndexBySheet(sheet, ws);
      const idx = headerIndexMapNormalized(ws, headerRowIndex);
      const idCol = idx.get("id");

      if (!idCol) throw new Error(`Sheet ${sheet} has no id column`);

      const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
      let rowNumber = -1;

      for (let r = headerRowIndex + 1; r <= last; r++) {
        const v = text(cellToPlain(ws.getRow(r).getCell(idCol)));
        if (v === String(id)) {
          rowNumber = r;
          break;
        }
      }

      if (rowNumber === -1) {
        throw new Error(`Row not found for id: ${id}`);
      }

      setRowFromObject(ws, rowNumber, payload, headerRowIndex);
      await atomicSave(wb);

      return rowToObject(ws, rowNumber, headerRowIndex);
    }
    const idx = headerIndexMapNormalized(ws, 1);
    const idCol = idx.get(normKey(key || "id"));
    if (!idCol) throw new Error(`Sheet ${sheet} has no ${key || "id"} column`);

    const last = ws.actualRowCount || ws.rowCount || 1;
    let rowNumber = -1;

    for (let r = 2; r <= last; r++) {
      const v = text(cellToPlain(ws.getRow(r).getCell(idCol)));
      if (v === id) {
        rowNumber = r;
        break;
      }
    }

    if (rowNumber === -1) {
      throw new Error(`Row not found for ${key || "id"}: ${id}`);
    }

    setRowFromObject(ws, rowNumber, payload, 1);
    await atomicSave(wb);

    return rowToObject(ws, rowNumber, 1);
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

    if (sheet === "Tools") {
      let changedTools = false;
      changedTools = prepareSheet(ws, sheet) || changedTools;
      changedTools = backfillToolRowKeys(ws) || changedTools;
      changedTools = backfillToolsIds(ws) || changedTools;

      const rowNumber = findRowByToolRowKey(ws, id);
      if (rowNumber === -1) {
        throw new Error(`Row not found for tool row key: ${id}`);
      }

      ws.spliceRows(rowNumber, 1);

      changedTools = backfillToolsIds(ws) || changedTools;
      await atomicSave(wb);

      return { ok: true };
    }
    if (isScheduleProjectSheetName(sheet)) {
      const headerRowIndex = getHeaderRowIndexBySheet(sheet, ws);
      const idx = headerIndexMapNormalized(ws, headerRowIndex);
      const idCol = idx.get("id");

      if (!idCol) throw new Error(`Sheet ${sheet} has no id column`);

      const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
      let rowNumber = -1;

      for (let r = headerRowIndex + 1; r <= last; r++) {
        const v = text(cellToPlain(ws.getRow(r).getCell(idCol)));
        if (v === String(id)) {
          rowNumber = r;
          break;
        }
      }

      if (rowNumber === -1) throw new Error(`Row not found for id: ${id}`);

      ws.spliceRows(rowNumber, 1);
      await atomicSave(wb);

      return { ok: true };
    }

    const idx = headerIndexMapNormalized(ws, 1);
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

function argbToCss(argb) {
  const raw = String(argb || "").trim();
  if (!raw) return "";
  const hex = raw.length === 8 ? raw.slice(2) : raw;
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return "";
  return `#${hex}`;
}

function formatGridCellText(value, colNumber) {
  if (value == null || value === "") return "";
  if (value instanceof Date) {
    const d = value;
    const month = d.toLocaleString("en-US", { month: "short" });
    if (colNumber >= 6) {
      return `${d.getDate()}-${month}-${String(d.getFullYear()).slice(-2)}`;
    }
    return `${String(d.getDate()).padStart(2, "0")}-${month}-${String(d.getFullYear()).slice(-2)}`;
  }
  if (typeof value === "object" && value?.formula != null) {
    if (value.result != null) return String(value.result);
    return `=${value.formula}`;
  }
  return String(value);
}

function detectHeaderRow(rows) {
  const wanted = [
    "id",
    "task name",
    "task name / milestone",
    "duration",
    "duration (days)",
    "start date",
    "end date",
    "finish date",
    "remarks",
    "notes / section",
    "% completion",
  ];

  for (let i = 0; i < rows.length; i++) {
    const texts = (rows[i].cells || [])
      .filter((c) => c && !c.skip)
      .map((c) =>
        String(c.text || "")
          .trim()
          .toLowerCase(),
      )
      .filter(Boolean);

    if (!texts.length) continue;

    const score = texts.filter((t) => wanted.includes(t)).length;

    const hasId = texts.includes("id");
    const hasTask =
      texts.includes("task name") || texts.includes("task name / milestone");
    const hasStart = texts.includes("start date");
    const hasEnd = texts.includes("end date") || texts.includes("finish date");

    if (score >= 3 || (hasId && hasTask && (hasStart || hasEnd))) {
      return i;
    }
  }

  return 2;
}

export async function getSheetGrid({ sheet } = {}) {
  return withWriteLock(async () => {
    const wb = await loadWorkbook();
    const ws = getSheetOrThrow(wb, sheet);

    let maxRow = 0;
    let maxCol = 0;

    for (let r = 1; r <= (ws.actualRowCount || ws.rowCount || 0); r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= (ws.columnCount || row.cellCount || 0); c++) {
        const cell = row.getCell(c);
        const plain = cellToPlain(cell);
        const fill = argbToCss(
          cell.fill?.fgColor?.argb || cell.fill?.bgColor?.argb,
        );
        const hasValue = plain !== "" && plain != null;
        if (hasValue || fill) {
          if (r > maxRow) maxRow = r;
          if (c > maxCol) maxCol = c;
        }
      }
    }

    maxRow = Math.max(maxRow, 1);
    maxCol = Math.max(maxCol, 1);

    const widths = [];
    for (let c = 1; c <= maxCol; c++) {
      if (c === 1) widths.push(260);
      else if (c === 2) widths.push(72);
      else if (c === 3) widths.push(92);
      else if (c === 4) widths.push(92);
      else if (c === 5) widths.push(92);
      else widths.push(22);
    }

    const mergeMap = new Map();
    const mergeSkip = new Set();
    for (const range of ws.model?.merges || []) {
      const m = String(range).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
      if (!m) continue;
      const colToNum = (letters) =>
        letters
          .toUpperCase()
          .split("")
          .reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      const c1 = colToNum(m[1]);
      const r1 = Number(m[2]);
      const c2 = colToNum(m[3]);
      const r2 = Number(m[4]);
      mergeMap.set(`${r1}:${c1}`, {
        rowSpan: r2 - r1 + 1,
        colSpan: c2 - c1 + 1,
      });
      for (let rr = r1; rr <= r2; rr++) {
        for (let cc = c1; cc <= c2; cc++) {
          if (rr === r1 && cc === c1) continue;
          mergeSkip.add(`${rr}:${cc}`);
        }
      }
    }

    const rows = [];
    for (let r = 1; r <= maxRow; r++) {
      const row = ws.getRow(r);
      const rowOut = { cells: [] };
      let nonEmptyCount = 0;

      for (let c = 1; c <= maxCol; c++) {
        if (mergeSkip.has(`${r}:${c}`)) {
          rowOut.cells.push({ skip: true });
          continue;
        }

        const cell = row.getCell(c);
        const raw = cellToPlain(cell);
        const txt = formatGridCellText(raw, c);
        const bg = argbToCss(
          cell.fill?.fgColor?.argb || cell.fill?.bgColor?.argb,
        );
        const align =
          cell.alignment?.horizontal ||
          (c === 1 ? "left" : c >= 6 ? "center" : "center");
        const merge = mergeMap.get(`${r}:${c}`) || { rowSpan: 1, colSpan: 1 };

        if (String(txt).trim() !== "" || bg) nonEmptyCount += 1;

        rowOut.cells.push({
          text: txt,
          bg,
          align,
          bold: !!cell.font?.bold,
          color: argbToCss(cell.font?.color?.argb),
          isDate: raw instanceof Date,
          width: widths[c - 1] || 100,
          stickyLeft: widths
            .slice(0, Math.max(0, c - 1))
            .reduce((a, b) => a + b, 0),
          colSpan: merge.colSpan,
          rowSpan: merge.rowSpan,
        });
      }

      const onlyFirstText =
        rowOut.cells.filter(
          (x) => !x.skip && String(x.text || "").trim() !== "",
        ).length === 1 && String(rowOut.cells[0]?.text || "").trim() !== "";
      rowOut.isTitle = onlyFirstText && r <= 4;
      rowOut.isSection = onlyFirstText && r > 4;
      rowOut.nonEmptyCount = nonEmptyCount;
      rows.push(rowOut);
    }

    const headerRowIndex = detectHeaderRow(rows);
    rows.forEach((row, idx) => {
      row.isHeader = idx === headerRowIndex;
      row.cells.forEach((cell) => {
        if (!cell || cell.skip) return;
        if (row.isHeader && cell.isDate) cell.bg = cell.bg || "#8fdd6a";
        if (row.isHeader && !cell.isDate) cell.bg = cell.bg || "#e5e7eb";
        if (row.isTitle) cell.isTitle = true;
        if (row.isSection) cell.isSection = true;
      });
    });

    const title =
      rows
        .find((r) => r.isTitle)
        ?.cells?.find((c) => !c.skip && String(c.text || "").trim())?.text ||
      sheet;
    const subtitle =
      rows
        .find((r, idx) => idx > 0 && r.isSection)
        ?.cells?.find((c) => !c.skip && String(c.text || "").trim())?.text ||
      "";

    return {
      ok: true,
      sheet,
      title,
      subtitle,
      leftFixedCount: 5,
      colWidths: widths,
      rows,
    };
  });
}

function normalizeSiteStatus(value) {
  const v = String(value ?? "")
    .trim()
    .toLowerCase();

  if (!v) return "";

  if (v === "active") return "Installation";
  if (v === "completed") return "Commissioned";
  if (v === "tbd") return "Pre Commissioning";

  if (v === "installation") return "Installation";
  if (v === "pre commissioning" || v === "pre-commissioning") {
    return "Pre Commissioning";
  }
  if (v === "wet commissioning" || v === "wet-commissioning") {
    return "Wet Commissioning";
  }
  if (v === "commissioned") return "Commissioned";

  return String(value ?? "").trim();
}

function excelColLetters(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cloneStyle(obj) {
  return obj ? JSON.parse(JSON.stringify(obj)) : obj;
}

function isBlankScheduleRow(row) {
  return ![
    row.id,
    row["Task Name"],
    row["% Completion"],
    row["Duration"],
    row["Start Date"],
    row["End Date"],
    row["Remarks"],
  ].some((v) => String(v ?? "").trim() !== "");
}

function buildScheduleExportRows(ws, sheetName) {
  const headerRowIndex = getHeaderRowIndexBySheet(sheetName, ws);
  const last = ws.actualRowCount || ws.rowCount || headerRowIndex;
  const rows = [];

  for (let r = headerRowIndex + 1; r <= last; r++) {
    const obj = rowToObject(ws, r, headerRowIndex);

    const hasAny = Object.values(obj).some(
      (v) => String(v ?? "").trim() !== "",
    );
    if (!hasAny) continue;

    const taskName = String(
      obj["Task Name"] ??
        obj["Task Name / Milestone"] ??
        obj["Milestone"] ??
        obj["Task"] ??
        "",
    ).trim();

    const startDate = parseAnyDate(obj["Start Date"] ?? obj["Start"]);
    const endDate = parseAnyDate(
      obj["End Date"] ?? obj["Finish Date"] ?? obj["Finish"],
    );

    let duration = parseInt(
      String(obj["Duration"] ?? obj["Duration (days)"] ?? "").replace(
        /[^\d]/g,
        "",
      ),
      10,
    );
    if (!Number.isFinite(duration) && startDate && endDate) {
      duration = daysInclusive(startDate, endDate);
    }
    if (!Number.isFinite(duration)) duration = "";

    rows.push({
      id: String(obj["id"] ?? "").trim(),
      taskName,
      completion: String(obj["% Completion"] ?? "").trim(),
      duration,
      startDate,
      endDate,
      remarks: String(obj["Remarks"] ?? obj["Notes / Section"] ?? "").trim(),
      __raw: obj,
    });
  }

  return rows;
}

function applyCellBorder(cell) {
  cell.border = {
    top: { style: "thin", color: { argb: "D7DEE8" } },
    left: { style: "thin", color: { argb: "D7DEE8" } },
    bottom: { style: "thin", color: { argb: "D7DEE8" } },
    right: { style: "thin", color: { argb: "D7DEE8" } },
  };
}

function paintTimelineCell(cell, color = "2E7D32") {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color },
  };
  applyCellBorder(cell);
}

function styleScheduleExportSheet(wsOut, totalTimelineCols, title, sheetName) {
  const leftCols = 7;
  const totalCols = leftCols + totalTimelineCols;

  wsOut.getCell("A1").value = title || sheetName.replace(/_timeline$/i, "");
  wsOut.getCell("A1").font = { bold: true, size: 18 };
  wsOut.mergeCells(1, 1, 1, 3);

  wsOut.getCell("A2").value = `Sheet: ${sheetName}`;
  wsOut.getCell("A2").font = { italic: true, color: { argb: "666666" } };
  wsOut.mergeCells(2, 1, 2, 3);

  wsOut.getRow(3).height = 34;
  wsOut.getRow(4).height = 28;

  wsOut.columns = [
    { width: 14 },
    { width: 36 },
    { width: 14 },
    { width: 12 },
    { width: 14 },
    { width: 14 },
    { width: 18 },
    ...Array.from({ length: totalTimelineCols }, () => ({ width: 4.2 })),
  ];

  const headerFillDark = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "1F6F2A" },
  };

  const headerFillLight = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "4CAF50" },
  };

  const headerFont = {
    bold: true,
    color: { argb: "FFFFFF" },
  };

  const headers = [
    "ID",
    "TASK NAME / MILESTONE",
    "% COMPLETION",
    "DURATION (DAYS)",
    "START DATE",
    "FINISH DATE",
    "NOTES / SECTION",
  ];

  headers.forEach((label, i) => {
    const cell = wsOut.getRow(3).getCell(i + 1);
    cell.value = label;
    cell.font = headerFont;
    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    cell.fill = headerFillDark;
    applyCellBorder(cell);
  });

  for (let c = 1; c <= totalCols; c++) {
    applyCellBorder(wsOut.getRow(4).getCell(c));
  }

  wsOut.views = [
    {
      state: "frozen",
      xSplit: 7,
      ySplit: 4,
    },
  ];

  return {
    headerFillLight,
    headerFillDark,
    headerFont,
  };
}

async function buildScheduleProjectWorkbookBuffer(
  sheetName,
  wbOverride = null,
) {
  const wb = wbOverride || (await loadWorkbook());
  const ws = getSheetOrThrow(wb, sheetName);

  if (!isScheduleProjectSheetName(sheetName)) {
    throw new Error(`Not a schedule project sheet: ${sheetName}`);
  }

  prepareScheduleProjectSheet(ws);

  const title =
    String(cellToPlain(ws.getCell("A1")) || "").trim() ||
    sheetName.replace(/_timeline$/i, "");

  const rows = buildScheduleExportRows(ws, sheetName);

  const out = new ExcelJS.Workbook();
  out.created = new Date();
  out.creator = "Troyer Nepal Dashboard";

  const wsOut = out.addWorksheet(
    String(sheetName)
      .replace(/[\\/*?:[\]]/g, "_")
      .slice(0, 31),
  );

  const taskRows = rows.filter((r) => r.startDate && r.endDate);

  if (!taskRows.length) {
    wsOut.addRow([title || sheetName]);
    wsOut.addRow([`Sheet: ${sheetName}`]);
    wsOut.addRow([
      "ID",
      "Task Name",
      "% Completion",
      "Duration",
      "Start Date",
      "End Date",
      "Remarks",
    ]);

    rows.forEach((r) => {
      wsOut.addRow([
        r.id,
        r.taskName,
        r.completion,
        r.duration,
        r.startDate ? toISODate(r.startDate) : "",
        r.endDate ? toISODate(r.endDate) : "",
        r.remarks,
      ]);
    });

    wsOut.columns = [
      { width: 14 },
      { width: 36 },
      { width: 14 },
      { width: 12 },
      { width: 14 },
      { width: 14 },
      { width: 18 },
    ];

    return out.xlsx.writeBuffer();
  }

  let minDate = taskRows[0].startDate;
  let maxDate = taskRows[0].endDate;

  for (const r of taskRows) {
    if (r.startDate < minDate) minDate = r.startDate;
    if (r.endDate > maxDate) maxDate = r.endDate;
  }

  const totalTimelineCols = daysInclusive(minDate, maxDate);
  const timelineStartCol = 8;

  const { headerFillLight, headerFillDark, headerFont } =
    styleScheduleExportSheet(wsOut, totalTimelineCols, title, sheetName);

  let currentMonthStart = 0;
  let currentMonthLabel = "";

  for (let i = 0; i < totalTimelineCols; i++) {
    const d = new Date(minDate);
    d.setDate(d.getDate() + i);

    const col = timelineStartCol + i;
    const monthLabel = d
      .toLocaleString("en-US", {
        month: "short",
        year: "2-digit",
      })
      .toUpperCase();

    const dayCell = wsOut.getRow(4).getCell(col);
    dayCell.value = d.getDate();
    dayCell.font = headerFont;
    dayCell.alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    dayCell.fill = headerFillLight;
    applyCellBorder(dayCell);

    if (monthLabel !== currentMonthLabel) {
      if (currentMonthLabel) {
        wsOut.mergeCells(3, timelineStartCol + currentMonthStart, 3, col - 1);
      }
      currentMonthStart = i;
      currentMonthLabel = monthLabel;
      wsOut.getRow(3).getCell(col).value = monthLabel;
    }

    wsOut.getColumn(col).width = 4.2;
  }

  wsOut.mergeCells(
    3,
    timelineStartCol + currentMonthStart,
    3,
    timelineStartCol + totalTimelineCols - 1,
  );

  for (
    let c = timelineStartCol;
    c < timelineStartCol + totalTimelineCols;
    c++
  ) {
    const cell = wsOut.getRow(3).getCell(c);
    cell.font = headerFont;
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = headerFillDark;
    applyCellBorder(cell);
  }

  let outRow = 5;

  for (const r of rows) {
    wsOut.getRow(outRow).height = 24;

    wsOut.getCell(outRow, 1).value = r.id || "";
    wsOut.getCell(outRow, 2).value = r.taskName || "";
    wsOut.getCell(outRow, 3).value = r.completion || "";
    wsOut.getCell(outRow, 4).value = r.duration || "";
    wsOut.getCell(outRow, 7).value = r.remarks || "";

    const toExcelSafeDate = (d) =>
      d ? new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0) : "";

    wsOut.getCell(outRow, 5).value = r.startDate
      ? toExcelSafeDate(r.startDate)
      : "";
    wsOut.getCell(outRow, 6).value = r.endDate
      ? toExcelSafeDate(r.endDate)
      : "";

    if (r.startDate) wsOut.getCell(outRow, 5).numFmt = "d-mmm-yy";
    if (r.endDate) wsOut.getCell(outRow, 6).numFmt = "d-mmm-yy";

    for (let c = 1; c <= 7; c++) {
      const cell = wsOut.getCell(outRow, c);
      cell.alignment = {
        vertical: "middle",
        horizontal: c === 2 || c === 7 ? "left" : "center",
      };
      applyCellBorder(cell);
    }

    for (let i = 0; i < totalTimelineCols; i++) {
      applyCellBorder(wsOut.getCell(outRow, timelineStartCol + i));
    }

    if (r.startDate && r.endDate) {
      const startOffset = Math.floor(
        (startOfDay(r.startDate).getTime() - startOfDay(minDate).getTime()) /
          86400000,
      );

      const span = daysInclusive(r.startDate, r.endDate);

      for (let i = 0; i < span; i++) {
        const col = timelineStartCol + startOffset + i;
        paintTimelineCell(wsOut.getCell(outRow, col), "388E3C");
      }
    }

    outRow += 1;
  }

  const today = startOfDay(new Date());
  if (today >= minDate && today <= maxDate) {
    const todayOffset = Math.floor(
      (today.getTime() - startOfDay(minDate).getTime()) / 86400000,
    );
    const todayCol = timelineStartCol + todayOffset;

    for (let r = 4; r < outRow; r++) {
      const cell = wsOut.getCell(r, todayCol);
      cell.border = {
        top: { style: "thin", color: { argb: "FF4D4F" } },
        left: { style: "thin", color: { argb: "FF4D4F" } },
        bottom: { style: "thin", color: { argb: "FF4D4F" } },
        right: { style: "thin", color: { argb: "FF4D4F" } },
      };
    }
  }

  return out.xlsx.writeBuffer();
}

export async function exportScheduleProjectWorkbookBuffer(sheetName) {
  return withWriteLock(async () => {
    return buildScheduleProjectWorkbookBuffer(sheetName);
  });
}

function copyWorksheetContents(srcWs, destWs) {
  const maxCol = srcWs.columnCount || 0;
  const maxRow = srcWs.actualRowCount || srcWs.rowCount || 0;

  for (let c = 1; c <= maxCol; c++) {
    const srcCol = srcWs.getColumn(c);
    const destCol = destWs.getColumn(c);
    destCol.width = srcCol.width;
    destCol.hidden = !!srcCol.hidden;
  }

  for (let r = 1; r <= maxRow; r++) {
    const srcRow = srcWs.getRow(r);
    const destRow = destWs.getRow(r);

    if (srcRow.height) destRow.height = srcRow.height;
    destRow.hidden = !!srcRow.hidden;

    for (let c = 1; c <= maxCol; c++) {
      const srcCell = srcRow.getCell(c);
      const destCell = destRow.getCell(c);

      destCell.value = srcCell.value;

      if (srcCell.numFmt) destCell.numFmt = srcCell.numFmt;
      if (srcCell.font)
        destCell.font = JSON.parse(JSON.stringify(srcCell.font));
      if (srcCell.alignment)
        destCell.alignment = JSON.parse(JSON.stringify(srcCell.alignment));
      if (srcCell.fill)
        destCell.fill = JSON.parse(JSON.stringify(srcCell.fill));
      if (srcCell.border)
        destCell.border = JSON.parse(JSON.stringify(srcCell.border));
      if (srcCell.protection)
        destCell.protection = JSON.parse(JSON.stringify(srcCell.protection));
    }
  }

  const merges = srcWs.model?.merges || [];
  for (const range of merges) {
    try {
      destWs.mergeCells(range);
    } catch {}
  }

  if (srcWs.views) {
    destWs.views = JSON.parse(JSON.stringify(srcWs.views));
  }

  if (srcWs.autoFilter) {
    destWs.autoFilter = JSON.parse(JSON.stringify(srcWs.autoFilter));
  }
}

export async function exportFullWorkbookWithGanttBuffer() {
  return withWriteLock(async () => {
    const src = await loadWorkbook();
    const out = new ExcelJS.Workbook();

    out.created = new Date();
    out.creator = "Troyer Nepal Dashboard";

    for (const ws of src.worksheets) {
      const sheetName = String(ws.name || "").trim();
      if (!sheetName) continue;

      if (
        isScheduleProjectSheetName(sheetName) &&
        sheetName !== "Seti_Updated"
      ) {
        const ganttBuf = await buildScheduleProjectWorkbookBuffer(
          sheetName,
          src,
        );

        const tmp = new ExcelJS.Workbook();
        await tmp.xlsx.load(ganttBuf);

        const ganttSheet = tmp.worksheets[0];
        const newWs = out.addWorksheet(
          String(sheetName)
            .replace(/[\\/*?:[\]]/g, "_")
            .slice(0, 31),
        );

        copyWorksheetContents(ganttSheet, newWs);
      } else {
        const newWs = out.addWorksheet(
          String(sheetName)
            .replace(/[\\/*?:[\]]/g, "_")
            .slice(0, 31),
        );
        copyWorksheetContents(ws, newWs);
      }
    }

    return out.xlsx.writeBuffer();
  });
}
