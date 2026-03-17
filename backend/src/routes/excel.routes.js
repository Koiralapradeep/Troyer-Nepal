import express from "express";
import ExcelJS from "exceljs";
import {
  getMasterPath,
  listRows,
  createRow,
  updateRow,
  deleteRow,
  ensureMasterExists,
  exportEmployeeWorkbookBuffer,
  exportScheduleProjectWorkbookBuffer,
  exportFullWorkbookWithGanttBuffer,
  getSheetGrid,
  createScheduleProjectSheet,
  renameScheduleProjectSheet,
  deleteScheduleProjectSheet,
} from "../excel/excelstore.js";


const router = express.Router();

router.get("/excel/status", async (req, res) => {
  try {
    const info = await ensureMasterExists();
    res.json({ ok: true, ...info });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.get("/excel/grid", async (req, res) => {
  try {
    const { sheet } = req.query;
    if (!sheet) {
      return res.status(400).json({ ok: false, error: "sheet is required" });
    }

    const out = await getSheetGrid({ sheet });
    res.json(out);
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.get("/excel/sheets", async (req, res) => {
  try {
    const p = await getMasterPath();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(p);
    res.json({ ok: true, sheets: wb.worksheets.map((w) => w.name) });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.get("/excel/download", async (req, res) => {
  try {
    const mode = String(req.query.mode || "").trim().toLowerCase();
    const sheet = String(req.query.sheet || "").trim();

    // Full database export with gantt for every schedule project
    if (mode === "full") {
      const buf = await exportFullWorkbookWithGanttBuffer();

      res.setHeader(
        "Content-Disposition",
        `attachment; filename="troyer_full_dashboard_export.xlsx"`
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      return res.send(Buffer.from(buf));
    }

    // Single schedule project export
    if (sheet && sheet.toLowerCase().endsWith("_timeline")) {
      const buf = await exportScheduleProjectWorkbookBuffer(sheet);
      const safe = sheet.replace(/[^a-z0-9_-]/gi, "_");

      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${safe}_dashboard_export.xlsx"`
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      return res.send(Buffer.from(buf));
    }

    const p = await getMasterPath();
    return res.download(p, "data.xlsx");
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.get("/rows", async (req, res) => {
  try {
    const { sheet, page, pageSize } = req.query;
    const out = await listRows({ sheet, page, pageSize });
    res.json(out);
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.post("/rows", async (req, res) => {
  try {
    const { sheet, idPrefix } = req.query;
    const item = await createRow({ sheet, idPrefix, data: req.body || {} });
    res.json({ ok: true, item });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.put("/rows/:id", async (req, res) => {
  try {
    const { sheet, key } = req.query;
    const item = await updateRow({
      sheet,
      key,                
      id: req.params.id,
      patch: req.body || {},
    });
    res.json({ ok: true, item });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.delete("/rows/:id", async (req, res) => {
  try {
    const { sheet } = req.query;
    const out = await deleteRow({ sheet, id: req.params.id });
    res.json(out);
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

router.get("/employees/:id/export", async (req, res) => {
  try {
    const employeeId = req.params.id;

    const buf = await exportEmployeeWorkbookBuffer(employeeId);

    const safe = employeeId.replace(/[^a-z0-9_-]/gi, "_");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${safe}_bundle.xlsx"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(400).json({ ok: false, error: String(e.message || e) });
  }
});

router.post("/excel/schedule-project",  async (req, res) => {
  try {
    const name = String(req.body?.name || "").trim();
    if (!name) {
      return res.status(400).json({
        ok: false,
        error: "Project name is required",
      });
    }

    const safeBase = name.replace(/[^\w\s-]/g, "").trim().replace(/\s+/g, "_");
    const sheetName = `${safeBase}_Timeline`;

    const result = await createScheduleProjectSheet({
      sheetName,
      title: name,
    });

    return res.json({
      ok: true,
      sheet: result.sheet,
      title: name,
    });
  } catch (err) {
    return res.status(500).json({
      ok: false,
      error: err.message || "Failed to create project",
    });
  }
});

router.put("/excel/schedule-project", async (req, res) => {
  try {
    const oldSheetName = String(req.body?.sheet || "").trim();
    const name = String(req.body?.name || "").trim();

    if (!oldSheetName) {
      return res.status(400).json({ ok: false, error: "Current project sheet is required" });
    }

    if (!name) {
      return res.status(400).json({ ok: false, error: "Project name is required" });
    }

    const result = await renameScheduleProjectSheet({
      oldSheetName,
      newTitle: name,
    });

    return res.json(result);
  } catch (err) {
    return res.status(500).json({
      ok: false,
      error: err.message || "Failed to rename project",
    });
  }
});

router.delete("/excel/schedule-project", async (req, res) => {
  try {
    const sheetName = String(req.query?.sheet || "").trim();

    if (!sheetName) {
      return res.status(400).json({ ok: false, error: "Project sheet is required" });
    }

    const result = await deleteScheduleProjectSheet({ sheetName });

    return res.json(result);
  } catch (err) {
    return res.status(500).json({
      ok: false,
      error: err.message || "Failed to delete project",
    });
  }
});


export default router;