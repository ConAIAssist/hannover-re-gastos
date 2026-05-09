import express from "express";
import jwt from "jsonwebtoken";
import multer from "multer";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { existsSync, mkdirSync } from "fs";
import { db, hashPw } from "./db.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3001;
const SECRET = process.env.JWT_SECRET || "hannover-re-gastos-secret-2026";

app.use(express.json({ limit: "50mb" }));

// File uploads
const uploadsDir = join(__dirname, "uploads");
if (!existsSync(uploadsDir)) mkdirSync(uploadsDir);
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname.replace(/[^a-zA-Z0-9.-]/g, "_")),
});
const upload = multer({ storage, limits: { fileSize: 10 * 1024 * 1024 } });

// Auth middleware
function auth(req, res, next) {
  const token = req.headers.authorization?.replace("Bearer ", "");
  if (!token) return res.status(401).json({ error: "No autorizado" });
  try { req.user = jwt.verify(token, SECRET); next(); }
  catch (e) { res.status(401).json({ error: "Token invalido" }); }
}

// ══════ AUTH ══════
app.post("/api/auth/login", (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ error: "Email y contrasena requeridos" });
  const user = db.getUserByEmail(email.trim());
  if (!user) return res.status(401).json({ error: "Correo no registrado en el sistema" });
  if (user.password !== hashPw(password)) return res.status(401).json({ error: "Contrasena incorrecta" });
  const token = jwt.sign({ id: user.id, email: user.email, role: user.role, name: user.name, area: user.area }, SECRET, { expiresIn: "24h" });
  const { password: _, ...safe } = user;
  res.json({ token, user: safe });
});

// ══════ USERS ══════
app.get("/api/users", auth, (req, res) => res.json(db.getUsers()));
app.post("/api/users", auth, (req, res) => {
  try { res.json(db.createUser(req.body)); }
  catch (e) { res.status(400).json({ error: e.message }); }
});
app.put("/api/users/:id", auth, (req, res) => {
  const u = db.updateUser(parseInt(req.params.id), req.body);
  u ? res.json(u) : res.status(404).json({ error: "No encontrado" });
});
app.delete("/api/users/:id", auth, (req, res) => { db.deleteUser(parseInt(req.params.id)); res.json({ ok: true }); });

// ══════ REPORTS ══════
app.get("/api/reports", auth, (req, res) => res.json(db.getReports()));

app.post("/api/reports", auth, (req, res) => {
  const { destination, tripPurpose, dateFrom, dateTo, currency, costCenter, type, status, expenses } = req.body;
  const user = db.getUserById(req.user.id);
  const isDraft = status === "draft";
  const r = db.createReport({
    employee_id: req.user.id, employee_name: user.name, area: user.area, cc: user.cc,
    destination, trip_purpose: tripPurpose, date_from: dateFrom, date_to: dateTo,
    presentation_date: new Date().toISOString().split("T")[0],
    currency: currency || "COP", cost_center: costCenter, type: type || "anticipo",
    status: isDraft ? "draft" : "submitted", current_step: isDraft ? 1 : 2,
    expenses: (expenses || []).map(e => ({ date: e.date, currency: e.currency, category: e.category, amount: e.amount, obs: e.obs, has_receipt: e.hasReceipt ? 1 : 0 })),
    approvals: isDraft ? [] : [{ step: 1, by_name: user.name, date: new Date().toISOString().split("T")[0], comment: "" }],
  });
  res.json(r);
});

app.put("/api/reports/:id", auth, (req, res) => {
  const { destination, tripPurpose, dateFrom, dateTo, currency, costCenter, type, expenses } = req.body;
  const r = db.updateReport(parseInt(req.params.id), {
    destination, trip_purpose: tripPurpose, date_from: dateFrom, date_to: dateTo,
    currency, cost_center: costCenter, type,
    expenses: (expenses || []).map(e => ({ id: e.id, date: e.date, currency: e.currency, category: e.category, amount: e.amount, obs: e.obs, has_receipt: e.has_receipt || e.hasReceipt ? 1 : 0, attachments: e.attachments || [] })),
  });
  r ? res.json(r) : res.status(404).json({ error: "No encontrado" });
});

app.post("/api/reports/:id/approve", auth, (req, res) => {
  const report = db.getReport(parseInt(req.params.id));
  if (!report) return res.status(404).json({ error: "No encontrado" });
  const user = db.getUserById(req.user.id);
  let ns = report.status, nc = report.current_step;
  if (nc === 2) { ns = "approved_leader"; nc = 3; }
  else if (nc === 3) { ns = "approved_admin"; nc = 4; }
  else if (nc === 4) { ns = "approved_accountant"; nc = 5; }
  else if (nc === 5) { ns = "approved_final"; }
  db.updateReport(report.id, { status: ns, current_step: nc });
  db.addApproval(report.id, { step: report.current_step, by_name: user.name, date: new Date().toISOString().split("T")[0], comment: req.body.comment || "" });
  res.json(db.getReport(report.id));
});

app.post("/api/reports/:id/reject", auth, (req, res) => {
  const report = db.getReport(parseInt(req.params.id));
  if (!report) return res.status(404).json({ error: "No encontrado" });
  const user = db.getUserById(req.user.id);
  db.updateReport(report.id, { status: "rejected" });
  db.addApproval(report.id, { step: report.current_step, by_name: user.name, date: new Date().toISOString().split("T")[0], comment: req.body.comment || "Rechazado" });
  res.json(db.getReport(report.id));
});

// ══════ ATTACHMENTS ══════
app.post("/api/expenses/:id/attach", auth, upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Archivo requerido" });
  const eid = parseInt(req.params.id);
  const atts = db.addAttachment(eid, { original_name: req.file.originalname, filename: req.file.filename, size: (req.file.size / 1024).toFixed(0) + " KB", mime_type: req.file.mimetype });
  atts ? res.json(atts) : res.status(404).json({ error: "Gasto no encontrado" });
});

app.get("/api/attachments/:id/download", auth, (req, res) => {
  const att = db.getAttachment(parseInt(req.params.id));
  if (!att) return res.status(404).json({ error: "No encontrado" });
  const fp = join(uploadsDir, att.filename);
  if (!existsSync(fp)) return res.status(404).json({ error: "Archivo no encontrado" });
  res.setHeader("Content-Disposition", `attachment; filename="${att.original_name}"`);
  res.sendFile(fp);
});

app.get("/api/attachments/:id/preview", auth, (req, res) => {
  const att = db.getAttachment(parseInt(req.params.id));
  if (!att) return res.status(404).json({ error: "No encontrado" });
  const fp = join(uploadsDir, att.filename);
  if (!existsSync(fp)) return res.status(404).json({ error: "No encontrado" });
  res.setHeader("Content-Type", att.mime_type || "application/octet-stream");
  res.sendFile(fp);
});

// ══════ CURRENCIES ══════
app.get("/api/currencies", auth, (req, res) => res.json(db.getCurrencies()));
app.post("/api/currencies", auth, (req, res) => {
  try { db.addCurrency({ code: req.body.code?.toUpperCase(), name: req.body.name, rate_to_cop: req.body.rateToCOP, symbol: req.body.symbol || req.body.code }); res.json({ ok: true }); }
  catch (e) { res.status(400).json({ error: e.message }); }
});
app.delete("/api/currencies/:code", auth, (req, res) => { db.removeCurrency(req.params.code); res.json({ ok: true }); });

// ══════ SERVE FRONTEND ══════
const distPath = join(__dirname, "dist");
console.log("Looking for dist at:", distPath, "exists:", existsSync(distPath));

app.use(express.static(distPath));
app.get("*", (req, res) => {
  if (req.path.startsWith("/api")) return res.status(404).json({ error: "Ruta no encontrada" });
  const indexPath = join(distPath, "index.html");
  if (existsSync(indexPath)) {
    res.sendFile(indexPath);
  } else {
    res.send(`<html><body><h1>Hannover Re Gastos API</h1><p>Backend activo. Frontend no encontrado en: ${distPath}</p><p>Ejecute: npm run build</p></body></html>`);
  }
});

app.listen(PORT, () => console.log(`Hannover Re Gastos API: http://localhost:${PORT}`));
