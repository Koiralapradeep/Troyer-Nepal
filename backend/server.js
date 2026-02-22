import path from "path";
import { fileURLToPath } from "url";
import express from "express";
import cors from "cors";
import excelRoutes from "./src/routes/excel.routes.js";

const app = express();
app.use(cors());
app.use(express.json());

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const CLIENT_PATH = path.join(__dirname, "../Client");
app.use(express.static(CLIENT_PATH));

app.get("/", (req, res) => {
  res.sendFile(path.join(CLIENT_PATH, "index.html"));
});

app.get("/health", (req, res) => res.json({ ok: true }));
app.use("/api", excelRoutes);

app.listen(3001, () => console.log("Server running at http://localhost:3001"));