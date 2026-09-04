const express = require("express");
const session = require("express-session");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const app = express();
const PORT = process.env.PORT || 3000;
// DATA_DIR can be overridden (e.g. to point at a Render persistent disk mount)
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "pilelog-data.json");

// --- Login setup (username + password) ---
// Set these as environment variables in production (e.g. on Render).
// If not set, the app falls back to the defaults below - change them!
const AUTH_USERNAME = process.env.AUTH_USERNAME || "admin";
const AUTH_PASSWORD = process.env.AUTH_PASSWORD || "changeme123";
const SESSION_SECRET = process.env.SESSION_SECRET || "please-change-this-session-secret";

fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) {
  fs.writeFileSync(DATA_FILE, JSON.stringify({ piles: [] }, null, 2), "utf8");
}

function loadData() {
  try {
    const value = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
    return value && Array.isArray(value.piles) ? value : { piles: [] };
  } catch {
    return { piles: [] };
  }
}

function saveData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), "utf8");
}

app.use(express.json({ limit: "5mb" }));

app.set("trust proxy", 1); // needed on Render/behind a proxy for secure cookies
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    maxAge: 1000 * 60 * 60 * 12 // 12 hours
  }
}));

// Public: login page assets and the login/logout API must work with no session yet.
app.get("/login.html", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "login.html"));
});
app.get("/styles.css", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "styles.css"));
});

app.post("/api/login", (req, res) => {
  const username = String((req.body && req.body.username) || "");
  const password = String((req.body && req.body.password) || "");
  if (username === AUTH_USERNAME && password === AUTH_PASSWORD) {
    req.session.authed = true;
    return res.json({ ok: true });
  }
  return res.status(401).json({ error: "Incorrect username or password." });
});

app.post("/api/logout", (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// Everything below this line requires a valid session.
function requireAuth(req, res, next) {
  if (req.session && req.session.authed) return next();
  if (req.path.startsWith("/api/")) {
    return res.status(401).json({ error: "Not authenticated" });
  }
  return res.redirect("/login.html?next=" + encodeURIComponent(req.originalUrl));
}
app.use(requireAuth);

app.use(express.static(path.join(__dirname, "public")));

app.get("/api/piles", (req, res) => {
  res.json(loadData().piles);
});

app.get("/api/piles/:id", (req, res) => {
  const piles = loadData().piles;
  const pile = piles.find(
    p => String(p.id) === String(req.params.id) ||
         String(p.pile_id || "").trim().toLowerCase() === String(req.params.id).trim().toLowerCase()
  );
  if (!pile) return res.status(404).json({ error: "Pile not found" });
  res.json(pile);
});

app.post("/api/piles", (req, res) => {
  const data = loadData();
  const pile = req.body || {};
  if (!String(pile.pile_id || "").trim()) {
    return res.status(400).json({ error: "Pile ID is required" });
  }
  if (data.piles.some(p =>
    String(p.pile_id || "").trim().toLowerCase() === String(pile.pile_id).trim().toLowerCase()
  )) {
    return res.status(409).json({ error: "Pile ID already exists" });
  }
  pile.pile_id = String(pile.pile_id).trim();
  pile.id = Date.now();
  data.piles.push(pile);
  saveData(data);
  res.status(201).json(pile);
});

app.put("/api/piles/:id", (req, res) => {
  const data = loadData();
  const index = data.piles.findIndex(
    p => String(p.id) === String(req.params.id) ||
         String(p.pile_id || "").trim().toLowerCase() === String(req.params.id).trim().toLowerCase()
  );
  if (index < 0) return res.status(404).json({ error: "Pile not found" });

  const updated = { ...data.piles[index], ...(req.body || {}), id: data.piles[index].id };
  if (!String(updated.pile_id || "").trim()) {
    return res.status(400).json({ error: "Pile ID is required" });
  }
  updated.pile_id = String(updated.pile_id).trim();
  data.piles[index] = updated;
  saveData(data);
  res.json(updated);
});

app.delete("/api/piles/:id", (req, res) => {
  const data = loadData();
  const before = data.piles.length;
  data.piles = data.piles.filter(
    p => String(p.id) !== String(req.params.id) &&
         String(p.pile_id || "").trim().toLowerCase() !== String(req.params.id).trim().toLowerCase()
  );
  if (data.piles.length === before) return res.status(404).json({ error: "Pile not found" });
  saveData(data);
  res.json({ ok: true });
});

function safeSheetName(name, used) {
  let base = String(name || "Pile").replace(/[\\/*?:[\]]/g, "_").slice(0, 31) || "Pile";
  let result = base;
  let n = 1;
  while (used.has(result)) {
    const suffix = "_" + n++;
    result = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(result);
  return result;
}

const COL_WIDTHS = [22,20,12,12,16,18,16,10,16,12,12,12,16,22,10,10,12];

function styleCell(ws, address, { bold = false, align, border = true } = {}) {
  const cell = ws.getCell(address);
  if (bold) cell.font = { bold: true };
  if (align) cell.alignment = { vertical: "middle", horizontal: align, wrapText: true };
  else cell.alignment = { vertical: "middle", wrapText: true };
  if (border) {
    cell.border = {
      top: { style: "thin" }, left: { style: "thin" },
      bottom: { style: "thin" }, right: { style: "thin" }
    };
  }
  return cell;
}

function mergeAndLabel(ws, range, value, opts = {}) {
  ws.mergeCells(range);
  const first = range.split(":")[0];
  const cell = styleCell(ws, first, opts);
  cell.value = value ?? "";
  return cell;
}

function toNum(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function toExcelTime(v) {
  if (!v) return null;
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(v).trim());
  if (!m) return null;
  const h = parseInt(m[1], 10), mm = parseInt(m[2], 10);
  return (h * 3600 + mm * 60) / 86400;
}

function setTimeCell(cell, raw) {
  const t = toExcelTime(raw);
  if (t != null) { cell.value = t; cell.numFmt = "HH:MM"; }
  else cell.value = raw || "";
}

function addPileSheet(workbook, pile, used) {
  const ws = workbook.addWorksheet(safeSheetName(pile.pile_id, used));
  ws.pageSetup.orientation = "landscape";
  ws.pageSetup.fitToPage = true;
  ws.pageSetup.fitToWidth = 1;
  ws.pageSetup.fitToHeight = 0;
  ws.columns = COL_WIDTHS.map(width => ({ width }));

  const boringLogs = Array.isArray(pile.boring_logs) ? pile.boring_logs : [];
  const boringCount = Math.max(1, boringLogs.length);
  const boringStart = 8;
  const boringEnd = boringStart + boringCount - 1;

  const reinforcement = Array.isArray(pile.reinforcement) ? pile.reinforcement : [];
  const reinCount = Math.max(1, reinforcement.length);

  // ---- Layout planning (rows below the fixed 4-11 general info block) ----
  const cageStartBase = 12;
  function layoutFrom(cageStart) {
    const cageRows = { header: cageStart, r1: cageStart + 1, r2: cageStart + 2, r3: cageStart + 3, r4: cageStart + 4 };
    const flushHeader = cageStart + 5;
    const flushRows = { header: flushHeader, r1: flushHeader + 1, r2: flushHeader + 2, r3: flushHeader + 3, r4: flushHeader + 4 };
    const reinHeader = flushHeader + 5;
    const reinSub1 = reinHeader + 1, reinSub2 = reinHeader + 2;
    const reinDataStart = reinSub2 + 1;
    const reinDataEnd = reinDataStart + reinCount - 1;
    const reinTotalLen = reinDataEnd + 1;
    const reinUnitWeight = reinDataEnd + 2;
    const reinTotalWeight = reinDataEnd + 3;
    const concHeader = reinTotalWeight + 2;
    const concRows = {
      header: concHeader,
      cement: concHeader + 1,
      grade: concHeader + 2,
      mix: concHeader + 3,
      slump: concHeader + 4,
      volumeTop: concHeader + 5,
      volumeBottom: concHeader + 6
    };
    return { cageRows, flushRows, reinHeader, reinSub1, reinSub2, reinDataStart, reinDataEnd,
      reinTotalLen, reinUnitWeight, reinTotalWeight, concRows, signatureStart: concRows.mix, signatureEnd: concRows.volumeBottom };
  }
  let layout = layoutFrom(cageStartBase);
  if (boringEnd >= layout.signatureStart) {
    const shift = boringEnd - layout.signatureStart + 1;
    layout = layoutFrom(cageStartBase + shift);
  }

  // ---- Header ----
  mergeAndLabel(ws, "A1:N1", pile.company_name || "Company Name", { bold: true, border: false });
  styleCell(ws, "O1", { bold: true, border: false }).value = "SI.No";
  styleCell(ws, "P1", { border: false }).value = pile.serial_no || "";
  mergeAndLabel(ws, "A2:N2", "Project : " + (pile.project_location || ""), { border: false });
  styleCell(ws, "O2", { bold: true, border: false }).value = "Date";
  const dateCell = styleCell(ws, "P2", { border: false });
  dateCell.value = pile.report_date ? new Date(pile.report_date) : "";
  if (pile.report_date) dateCell.numFmt = "dd-mmm-yyyy";
  mergeAndLabel(ws, "B3:I3", "PILE BORE LOG DETAILS", { bold: true, align: "center" });
  mergeAndLabel(ws, "J3:P3", "BORING DETAILS", { bold: true, align: "center" });

  // ---- General info (fixed rows 4-11) ----
  const leftInfo = [
    [4, "Rig No", pile.rig_no],
    [5, "Location", pile.location],
    [6, "Drawing no.", pile.drawing_no],
    [7, "Pile No.", pile.pile_id],
    [8, "Dia. of Pile", pile.dia_of_pile],
    [9, "Depth of Pile from CTL  (M)", pile.depth_of_pile_ctl],
    [10, "Depth of Pile from Cut Of Level (M)", null], // formula below
    [11, "Founding Level", pile.founding_level]
  ];
  leftInfo.forEach(([row, label, value]) => {
    styleCell(ws, `A${row}`, { bold: true }).value = label;
    const vcell = mergeAndLabel(ws, `B${row}:E${row}`, value);
    if (row === 10) vcell.value = { formula: "G9-(G8-G6)" };
  });

  const rightInfo = [
    [4, "Type", pile.type],
    [5, "Co-ordinates", pile.coordinates],
    [6, "Cut-off level", toNum(pile.cutoff_level) ?? pile.cutoff_level],
    [7, "Top of Casing/Liner", toNum(pile.top_of_casing) ?? pile.top_of_casing],
    [8, "Existing GL", toNum(pile.existing_gl) ?? pile.existing_gl],
    [9, "Bore Depth from GL (M)", null], // formula below
    [10, "Cage Length", toNum(pile.cage_length) ?? pile.cage_length],
    [11, "Boring Completed Date", null]
  ];
  rightInfo.forEach(([row, label, value]) => {
    styleCell(ws, `F${row}`, { bold: true }).value = label;
    const vcell = mergeAndLabel(ws, `G${row}:I${row}`, value);
    if (row === 9) vcell.value = { formula: `SUM(O${boringStart}:O${boringEnd})` };
    if (row === 11) {
      if (pile.boring_completed_date) {
        vcell.value = new Date(pile.boring_completed_date);
        vcell.numFmt = "dd-mmm-yyyy";
      } else vcell.value = "";
    }
  });

  // ---- Boring details table ----
  mergeAndLabel(ws, `J4:J${boringStart - 1}`, "Date", { bold: true, align: "center" });
  mergeAndLabel(ws, `K4:L5`, "Time (Hrs)", { bold: true, align: "center" });
  mergeAndLabel(ws, `K6:K7`, "From", { bold: true, align: "center" });
  mergeAndLabel(ws, `L6:L7`, "To", { bold: true, align: "center" });
  mergeAndLabel(ws, `M4:M${boringStart - 1}`, "Depth (M) from EGL", { bold: true, align: "center" });
  mergeAndLabel(ws, `N4:N${boringStart - 1}`, "Description of Soil", { bold: true, align: "center" });
  mergeAndLabel(ws, `O4:O${boringStart - 1}`, "Penetration", { bold: true, align: "center" });
  mergeAndLabel(ws, `P4:P${boringStart - 1}`, "Remarks", { bold: true, align: "center" });

  for (let i = 0; i < boringCount; i++) {
    const row = boringStart + i;
    const item = boringLogs[i] || {};
    styleCell(ws, `J${row}`).value = item.date ? new Date(item.date) : "";
    if (item.date) ws.getCell(`J${row}`).numFmt = "dd-mmm-yyyy";
    setTimeCell(styleCell(ws, `K${row}`), item.time_from);
    setTimeCell(styleCell(ws, `L${row}`), item.time_to);
    const depth = toNum(item.depth);
    styleCell(ws, `M${row}`).value = depth ?? (item.depth || "");
    styleCell(ws, `N${row}`).value = item.description || "";
    const oCell = styleCell(ws, `O${row}`);
    if (depth != null) {
      oCell.value = i === 0 ? { formula: `M${row}` } : { formula: `M${row}-M${row - 1}` };
    } else {
      oCell.value = "";
    }
    styleCell(ws, `P${row}`).value = item.remarks || "";
  }

  // ---- Cage / Trimmer Lowering ----
  const { cageRows } = layout;
  mergeAndLabel(ws, `A${cageRows.header}:E${cageRows.header}`, "CAGE LOWERING DETAILS", { bold: true, align: "center" });
  mergeAndLabel(ws, `F${cageRows.header}:I${cageRows.header}`, "TRIMMER LOWERING DETAILS", { bold: true, align: "center" });

  const dateOrBlank = v => v ? new Date(v) : "";
  const cageTrimmerRows = [
    [cageRows.r1, "Started Date", pile.cage_started_date, "Started Date", pile.trimmer_started_date, true],
    [cageRows.r2, "Time Started", pile.cage_started_time, "Time Started", pile.trimmer_started_time, false],
    [cageRows.r3, "Completed Date", pile.cage_completed_date, "Completed Date", pile.trimmer_completed_date, true],
    [cageRows.r4, "Time Completed", pile.cage_completed_time, "Time Completed", pile.trimmer_completed_time, false]
  ];
  cageTrimmerRows.forEach(([row, labelL, valL, labelR, valR, isDate]) => {
    mergeAndLabel(ws, `A${row}:B${row}`, labelL, { bold: true });
    ws.mergeCells(`C${row}:E${row}`);
    const c1 = styleCell(ws, `C${row}`);
    if (isDate) { c1.value = dateOrBlank(valL); if (valL) c1.numFmt = "dd-mmm-yyyy"; }
    else setTimeCell(c1, valL);
    mergeAndLabel(ws, `F${row}:G${row}`, labelR, { bold: true });
    ws.mergeCells(`H${row}:I${row}`);
    const c2 = styleCell(ws, `H${row}`);
    if (isDate) { c2.value = dateOrBlank(valR); if (valR) c2.numFmt = "dd-mmm-yyyy"; }
    else setTimeCell(c2, valR);
  });

  // ---- Flushing ----
  const { flushRows } = layout;
  mergeAndLabel(ws, `A${flushRows.header}:I${flushRows.header}`, "FLUSHING DETAILS", { bold: true, align: "center" });
  mergeAndLabel(ws, `A${flushRows.r1}:B${flushRows.r1}`, "Started Date", { bold: true });
  const fsd = mergeAndLabel(ws, `C${flushRows.r1}:E${flushRows.r1}`, dateOrBlank(pile.flushing_started_date));
  if (pile.flushing_started_date) fsd.numFmt = "dd-mmm-yyyy";
  mergeAndLabel(ws, `F${flushRows.r1}:G${flushRows.r1}`, "Completed Date", { bold: true });
  const fcd = mergeAndLabel(ws, `H${flushRows.r1}:I${flushRows.r1}`, dateOrBlank(pile.flushing_completed_date));
  if (pile.flushing_completed_date) fcd.numFmt = "dd-mmm-yyyy";

  mergeAndLabel(ws, `A${flushRows.r2}:B${flushRows.r2}`, "Time Started", { bold: true });
  ws.mergeCells(`C${flushRows.r2}:E${flushRows.r2}`);
  setTimeCell(styleCell(ws, `C${flushRows.r2}`), pile.flushing_started_time);
  mergeAndLabel(ws, `F${flushRows.r2}:G${flushRows.r2}`, "Time Completed", { bold: true });
  ws.mergeCells(`H${flushRows.r2}:I${flushRows.r2}`);
  setTimeCell(styleCell(ws, `H${flushRows.r2}`), pile.flushing_completed_time);

  mergeAndLabel(ws, `A${flushRows.r3}:B${flushRows.r4}`, "Specific Gravity of Bentonite", { bold: true });
  mergeAndLabel(ws, `C${flushRows.r3}:E${flushRows.r3}`, pile.specific_gravity || "");
  mergeAndLabel(ws, `F${flushRows.r3}:G${flushRows.r3}`, "Before Flushing", { bold: true, align: "center" });
  mergeAndLabel(ws, `H${flushRows.r3}:I${flushRows.r3}`, "After Flushing", { bold: true, align: "center" });
  mergeAndLabel(ws, `C${flushRows.r4}:E${flushRows.r4}`, "");
  mergeAndLabel(ws, `F${flushRows.r4}:G${flushRows.r4}`, pile.bentonite_before || "-");
  mergeAndLabel(ws, `H${flushRows.r4}:I${flushRows.r4}`, pile.bentonite_after || "-");

  // ---- Reinforcement details ----
  mergeAndLabel(ws, `A${layout.reinHeader}:I${layout.reinHeader}`, "REINFORCEMENT DETAILS", { bold: true, align: "center" });
  mergeAndLabel(ws, `A${layout.reinSub1}:A${layout.reinSub2}`, "Description of Bars", { bold: true, align: "center" });
  mergeAndLabel(ws, `B${layout.reinSub1}:B${layout.reinSub2}`, "Diameter", { bold: true, align: "center" });
  mergeAndLabel(ws, `C${layout.reinSub1}:C${layout.reinSub2}`, "Nos", { bold: true, align: "center" });
  mergeAndLabel(ws, `D${layout.reinSub1}:D${layout.reinSub2}`, "Length (m)", { bold: true, align: "center" });
  mergeAndLabel(ws, `E${layout.reinSub1}:I${layout.reinSub1}`, "Total Length (Diameter Wise)", { bold: true, align: "center" });
  const diaCols = { 25: "E", 16: "F", 12: "G", 8: "H" };
  ["25 mm", "16 mm", "12 mm", "8 mm", "Total"].forEach((h, i) => {
    styleCell(ws, `${String.fromCharCode(69 + i)}${layout.reinSub2}`, { bold: true, align: "center" }).value = h;
  });

  for (let i = 0; i < reinCount; i++) {
    const row = layout.reinDataStart + i;
    const item = reinforcement[i] || {};
    styleCell(ws, `A${row}`).value = item.description || "";
    const dia = toNum(item.diameter ?? item.dia);
    styleCell(ws, `B${row}`).value = dia ?? (item.diameter ?? item.dia ?? "");
    const nos = toNum(item.nos);
    styleCell(ws, `C${row}`).value = nos ?? (item.nos || "");
    const len = toNum(item.length);
    styleCell(ws, `D${row}`).value = len ?? (item.length || "");
    "EFGH".split("").forEach(col => styleCell(ws, `${col}${row}`).value = "");
    if (dia && diaCols[dia] && nos != null && len != null) {
      ws.getCell(`${diaCols[dia]}${row}`).value = { formula: `C${row}*D${row}` };
    }
  }

  const sumRange = col => `SUM(${col}${layout.reinDataStart}:${col}${layout.reinDataEnd})`;
  mergeAndLabel(ws, `A${layout.reinTotalLen}:D${layout.reinTotalLen}`, 'Total Length in "M"', { bold: true });
  "EFGH".split("").forEach(col => { ws.getCell(`${col}${layout.reinTotalLen}`).value = { formula: sumRange(col) }; });

  mergeAndLabel(ws, `A${layout.reinUnitWeight}:D${layout.reinUnitWeight}`, 'Steel Unit Weight "Kgs"', { bold: true });
  ws.getCell(`E${layout.reinUnitWeight}`).value = { formula: "25^2/162.2" };
  ws.getCell(`F${layout.reinUnitWeight}`).value = { formula: "16^2/162.2" };
  ws.getCell(`G${layout.reinUnitWeight}`).value = { formula: "12^2/162.2" };
  ws.getCell(`H${layout.reinUnitWeight}`).value = { formula: "8^2/162.2" };

  mergeAndLabel(ws, `A${layout.reinTotalWeight}:D${layout.reinTotalWeight}`, 'Total weight in "Kgs"', { bold: true });
  "EFGH".split("").forEach(col => {
    ws.getCell(`${col}${layout.reinTotalWeight}`).value = { formula: `${col}${layout.reinUnitWeight}*${col}${layout.reinTotalLen}` };
  });
  ws.getCell(`I${layout.reinTotalWeight}`).value = { formula: `SUM(E${layout.reinTotalWeight}:H${layout.reinTotalWeight})` };

  // ---- Concreting details ----
  const { concRows } = layout;
  mergeAndLabel(ws, `A${concRows.header}:I${concRows.header}`, "CONCRETING DETAILS", { bold: true, align: "center" });

  mergeAndLabel(ws, `A${concRows.cement}:B${concRows.cement}`, "Type of Cement", { bold: true });
  mergeAndLabel(ws, `C${concRows.cement}:E${concRows.cement}`, pile.cement_type || "");
  styleCell(ws, `F${concRows.cement}`, { bold: true }).value = "Commence Date";
  const ccd = styleCell(ws, `G${concRows.cement}`);
  ccd.value = dateOrBlank(pile.concrete_commence_date);
  if (pile.concrete_commence_date) ccd.numFmt = "dd-mmm-yyyy";
  styleCell(ws, `H${concRows.cement}`, { bold: true }).value = "Time";
  setTimeCell(styleCell(ws, `I${concRows.cement}`), pile.concrete_commence_time);

  mergeAndLabel(ws, `A${concRows.grade}:B${concRows.grade}`, "Grade of Concrete", { bold: true });
  mergeAndLabel(ws, `C${concRows.grade}:E${concRows.grade}`, pile.concrete_grade || "");
  styleCell(ws, `F${concRows.grade}`, { bold: true }).value = "Completed Date";
  const cpd = styleCell(ws, `G${concRows.grade}`);
  cpd.value = dateOrBlank(pile.concrete_completed_date);
  if (pile.concrete_completed_date) cpd.numFmt = "dd-mmm-yyyy";
  styleCell(ws, `H${concRows.grade}`, { bold: true }).value = "Time";
  setTimeCell(styleCell(ws, `I${concRows.grade}`), pile.concrete_completed_time);

  mergeAndLabel(ws, `A${concRows.mix}:B${concRows.mix}`, "Design Mix Ratio", { bold: true });
  mergeAndLabel(ws, `C${concRows.mix}:E${concRows.mix}`, pile.design_mix || "");
  mergeAndLabel(ws, `F${concRows.mix}:F${concRows.mix + 1}`, "No, of cubes Taken", { bold: true });
  styleCell(ws, `G${concRows.mix}`).value = toNum(pile.cubes_taken) ?? (pile.cubes_taken || "");
  mergeAndLabel(ws, `J${layout.signatureStart}:M${layout.signatureEnd}`, "M/s. " + (pile.company_name || ""), { bold: true, align: "center" });
  mergeAndLabel(ws, `N${layout.signatureStart}:P${layout.signatureEnd}`, "M/s. " + (pile.client_name || ""), { bold: true, align: "center" });

  mergeAndLabel(ws, `A${concRows.slump}:B${concRows.slump}`, "Slump at Site", { bold: true });
  mergeAndLabel(ws, `C${concRows.slump}:E${concRows.slump}`, pile.slump || "");

  mergeAndLabel(ws, `A${concRows.volumeTop}:E${concRows.volumeBottom}`, "Volume of Concrete (M3)", { bold: true, align: "center" });
  mergeAndLabel(ws, `F${concRows.volumeTop}:G${concRows.volumeTop}`, "Theorectical", { bold: true, align: "center" });
  mergeAndLabel(ws, `H${concRows.volumeTop}:I${concRows.volumeTop}`, "Actual", { bold: true, align: "center" });
  const diaMm = toNum(pile.dia_of_pile);
  const radius = diaMm ? (diaMm / 1000) / 2 : 0.3;
  mergeAndLabel(ws, `F${concRows.volumeBottom}:G${concRows.volumeBottom}`, { formula: `PI()*${radius}^2*G9` }, { align: "center" });
  mergeAndLabel(ws, `H${concRows.volumeBottom}:I${concRows.volumeBottom}`, toNum(pile.actual_concrete_volume) ?? (pile.actual_concrete_volume || ""), { align: "center" });

  return ws;
}

app.get("/api/export", async (req, res) => {
  try {
    const data = loadData();
    let piles = data.piles;
    if (req.query.ids) {
      const ids = String(req.query.ids).split(",").map(x=>x.trim().toLowerCase());
      piles = piles.filter(p => ids.includes(String(p.pile_id || "").trim().toLowerCase()));
    }
    if (!piles.length) return res.status(404).json({ error: "No piles available for export" });

    const workbook = new ExcelJS.Workbook();
    const used = new Set();
    piles.forEach(p => addPileSheet(workbook, p, used));

    const file = path.join(DATA_DIR, `pile-bore-logs-${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(file);
    res.download(file, path.basename(file), err => {
      try { fs.unlinkSync(file); } catch {}
      if (err) console.error(err);
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Pile Log Web App running at http://localhost:${PORT}`);
});
