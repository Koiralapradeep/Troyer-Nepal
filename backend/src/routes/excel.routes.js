import express from "express";
import ExcelJS from "exceljs";
import {
  getMasterPath,
  listRows,
  createRow,
  updateRow,
  deleteRow,
  ensureMasterExists,
} from "../excel/excelstore.js";
import { exportEmployeeWorkbookBuffer } from "../excel/excelstore.js";


const router = express.Router();

router.get("/excel/status", async (req, res) => {
  try {
    const info = await ensureMasterExists();
    res.json({ ok: true, ...info });
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
    const p = await getMasterPath();
    res.download(p, "data.xlsx");
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

export default router;