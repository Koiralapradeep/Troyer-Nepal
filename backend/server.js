import path from "path";
import { fileURLToPath } from "url";
import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import jwt from "jsonwebtoken";
import "dotenv/config";

import excelRoutes from "./src/routes/excel.routes.js";
import authRoutes, { requireAuth } from "./src/routes/auth.routes.js";

const app = express();

app.set("trust proxy", 1);
app.use(cors({ origin: true, credentials: true }));
app.use(express.json());
app.use(cookieParser());

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const CLIENT_PATH = path.join(__dirname, "../Client");

function isAuthed(req) {
  try {
    const token = req.cookies?.tn_token;
    if (!token) return false;
    jwt.verify(token, process.env.JWT_SECRET);
    return true;
  } catch {
    return false;
  }
}

app.get("/", (req, res) => {
  return res.redirect(isAuthed(req) ? "/app" : "/login");
});

app.get("/login", (req, res) => {
  if (isAuthed(req)) return res.redirect("/app");
  return res.sendFile(path.join(CLIENT_PATH, "login.html"));
});

app.get("/app", (req, res) => {
  if (!isAuthed(req)) return res.redirect("/login");
  return res.sendFile(path.join(CLIENT_PATH, "index.html"));
});

app.get("/health", (req, res) => res.json({ ok: true }));

app.use("/api/auth", authRoutes);
app.use("/api", requireAuth, excelRoutes);

app.use(express.static(CLIENT_PATH));

const PORT = process.env.PORT ? Number(process.env.PORT) : 3001;

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});