// =====================================================================
//  server.js — DocCode Generator (backend complet)
//  Codification documentaire QMS/ISO — OCG Consulting
//
//  Endpoints :
//    GET    /api/health
//    GET    /api/codes
//    GET    /api/counters
//    POST   /api/codes            body: { type, intitule, customCode? }
//    DELETE /api/codes            (vide tout le tableau)
//    DELETE /api/codes/:id
//    GET    /api/export/excel
//    GET    /api/export/word
//
//  Installation :
//    npm install express cors exceljs docx
//    npm start            (ou: node server.js)
//
//  Stockage : fichier db.json (créé automatiquement à côté de ce fichier)
// =====================================================================

const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const {
  Document, Packer, Table, TableRow, TableCell, Paragraph, TextRun,
  WidthType, AlignmentType, BorderStyle, VerticalAlign, ShadingType,
} = require("docx");

const app = express();
const PORT = 5000;
const DB_FILE = path.join(__dirname, "db.json");

app.use(cors());
app.use(express.json());

// ---------------------------------------------------------------------
//  Référentiel des types
// ---------------------------------------------------------------------
const TYPE_MAP_BACKEND = {
  processus:      "PS",
  procedure:      "PR",
  enregistrement: "ENG",
  formulaire:     "F",
  logigramme:     "L",
  instruction:    "INS",
  modeoperatoire: "MO",
};

// Libellés (pluriel) pour la colonne "Type de document" des exports
const TYPE_LABELS = {
  PS:  "Processus",
  PR:  "Procédures",
  ENG: "Enregistrements",
  F:   "Formulaires",
  L:   "Logigrammes",
  INS: "Instructions",
  MO:  "Modes opératoires",
};
const TYPE_ORDER = ["PS", "PR", "ENG", "F", "L", "INS", "MO"];

const IGNORE = new Set([
  "DE","DES","ET","LA","LE","LES","DU","D","AU","AUX","UN","UNE",
  "EN","PAR","SUR","POUR","AVEC","DANS","L","A","SE","SA","SES",
  "OU","QUE","QUI","SOUS","VERS","ENTRE","DONT","SON","CAS",
  "PROCESSUS","PROCEDURE","PROCEDURES","ENREGISTREMENT","ENREGISTREMENTS",
  "FORMULAIRE","FORMULAIRES","FICHE","FICHES","LOGIGRAMME","LOGIGRAMMES",
  "INSTRUCTION","INSTRUCTIONS","MODE","OPERATOIRE","OPERATOIRES",
]);

// ---------------------------------------------------------------------
//  Persistance (db.json)
// ---------------------------------------------------------------------
function loadCodes() {
  try {
    if (!fs.existsSync(DB_FILE)) return [];
    const raw = fs.readFileSync(DB_FILE, "utf-8");
    const data = JSON.parse(raw || "[]");
    return Array.isArray(data) ? data : (data.codes || []);
  } catch (e) {
    console.error("Lecture db.json :", e.message);
    return [];
  }
}
function saveCodes(codes) {
  fs.writeFileSync(DB_FILE, JSON.stringify(codes, null, 2), "utf-8");
}

// Fonction utilisée par les exports (toujours la même source)
async function getAllCodes() {
  return loadCodes();
}

// ---------------------------------------------------------------------
//  Génération du code
// ---------------------------------------------------------------------
function makeAbbrev(intitule) {
  const words = intitule.toUpperCase().normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z\s]/g, " ")
    .split(/\s+/).filter(w => w && !IGNORE.has(w));
  let abbrev = words.map(w => w[0]).join("").slice(0, 5);
  if (abbrev.length === 1 && words[0]) abbrev = words[0].slice(0, 3);
  return abbrev || "DOC";
}

function nextNumber(codes, typeCode) {
  // numéro = plus grand numéro existant pour ce type + 1 (robuste aux suppressions)
  let max = 0;
  for (const c of codes) {
    if (c.type === typeCode && typeof c.seq === "number" && c.seq > max) max = c.seq;
  }
  return max + 1;
}

// ---------------------------------------------------------------------
//  Routes API
// ---------------------------------------------------------------------
app.get("/api/health", (req, res) => {
  res.json({ success: true, message: "OK", time: new Date().toISOString() });
});

app.get("/api/codes", (req, res) => {
  res.json({ success: true, data: loadCodes() });
});

app.get("/api/counters", (req, res) => {
  const codes = loadCodes();
  const counters = { PS:0, PR:0, ENG:0, F:0, L:0, INS:0, MO:0 };
  for (const c of codes) {
    counters[c.type] = (counters[c.type] || 0) + 1;
  }
  res.json({ success: true, data: counters });
});

app.post("/api/codes", (req, res) => {
  try {
    const { type, intitule, customCode } = req.body;
    if (!type || !intitule || !intitule.trim()) {
      return res.json({ success: false, message: "Type et intitulé requis" });
    }

    // typeCode : soit le code custom (3 lettres), soit la table de correspondance
    const typeCode = customCode
      ? customCode.trim().toUpperCase().slice(0, 3)
      : TYPE_MAP_BACKEND[type];

    if (!typeCode) {
      return res.json({ success: false, message: "Type inconnu" });
    }

    const codes = loadCodes();
    const abbrev = makeAbbrev(intitule);
    const seq = nextNumber(codes, typeCode);
    const numStr = String(seq).padStart(2, "0");
    const code = `${typeCode}-${abbrev}-${numStr}`;

    const entry = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      code,
      intitule: intitule.trim(),
      type: typeCode,
      seq,
      version: 0,
      motif: "Création",
      createdAt: new Date().toISOString(),
    };

    codes.push(entry);
    saveCodes(codes);

    res.json({ success: true, data: entry });
  } catch (e) {
    console.error("POST /api/codes :", e);
    res.status(500).json({ success: false, message: "Erreur serveur" });
  }
});

// Vider tout le tableau (les compteurs sont recalculés à partir de db.json)
// IMPORTANT : déclarée AVANT "/api/codes/:id" pour éviter toute ambiguïté de routage.
app.delete("/api/codes", (req, res) => {
  try {
    saveCodes([]);
    res.json({ success: true, message: "Tous les codes ont été supprimés" });
  } catch (e) {
    console.error("DELETE /api/codes :", e);
    res.status(500).json({ success: false, message: "Erreur suppression" });
  }
});

app.delete("/api/codes/:id", (req, res) => {
  const codes = loadCodes();
  const filtered = codes.filter(c => c.id !== req.params.id);
  saveCodes(filtered);
  res.json({ success: true });
});

// ---------------------------------------------------------------------
//  EXPORT EXCEL  (registre "Liste des Informations Documentées")
// ---------------------------------------------------------------------
const NAVY  = "FF1F3864";
const LIGHT = "FFD9E1F2";
const XFONT = "Times New Roman";

function buildWorkbook(codes) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "OCG Consulting — DocCode Generator";
  const ws = wb.addWorksheet("Informations documentées", {
    views: [{ state: "frozen", ySplit: 5 }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  ws.columns = [
    { width: 22 }, { width: 50 }, { width: 18 },
    { width: 10 }, { width: 13 }, { width: 24 },
  ];

  const thin = { style: "thin", color: { argb: "FFBFBFBF" } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  // Bloc en-tête
  ws.mergeCells("A1:B3");
  const logo = ws.getCell("A1");
  logo.value = "OCG\nCONSULTING";
  logo.font = { name: XFONT, bold: true, size: 14, color: { argb: "FFFFFFFF" } };
  logo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  logo.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.mergeCells("C1:D3");
  const titre = ws.getCell("C1");
  titre.value = "Liste des Informations Documentées";
  titre.font = { name: XFONT, bold: true, size: 13, color: { argb: NAVY } };
  titre.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.mergeCells("E1:F1"); ws.getCell("E1").value = "Code : ENG-LID-01";
  ws.mergeCells("E2:F2"); ws.getCell("E2").value = "Version : 00";
  ws.mergeCells("E3:F3");
  ws.getCell("E3").value = "Date : " + new Date().toLocaleDateString("fr-FR");
  ["E1", "E2", "E3"].forEach((c) => {
    ws.getCell(c).font = { name: XFONT, size: 10 };
    ws.getCell(c).alignment = { horizontal: "left", vertical: "middle" };
  });
  for (let r = 1; r <= 3; r++)
    for (let c = 1; c <= 6; c++) ws.getCell(r, c).border = border;

  // En-tête tableau (ligne 5)
  const HR = 5;
  const headers = ["Type de document", "Intitulé du document", "Code",
                   "Version", "Date", "Motif de mise à jour"];
  const headerRow = ws.getRow(HR);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: XFONT, bold: true, size: 11, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = border;
  });
  headerRow.height = 26;

  // Données groupées par type + fusion verticale
  let r = HR + 1;
  TYPE_ORDER.forEach((t) => {
    const group = codes.filter((c) => c.type === t);
    if (group.length === 0) return;
    const start = r;
    group.forEach((item) => {
      const row = ws.getRow(r);
      row.getCell(1).value = TYPE_LABELS[t];
      row.getCell(2).value = item.intitule || "";
      row.getCell(3).value = item.code || "";
      row.getCell(4).value = item.version ?? 0;
      row.getCell(5).value = item.createdAt ? new Date(item.createdAt) : new Date();
      row.getCell(6).value = item.motif || "Création";
      for (let col = 1; col <= 6; col++) {
        const cell = row.getCell(col);
        cell.border = border;
        cell.font = { name: XFONT, size: 10 };
        cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        if (col === 1) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT } };
          cell.font = { name: XFONT, size: 11, bold: true, color: { argb: NAVY } };
          cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        } else if (col === 3) {
          cell.font = { name: "Consolas", size: 10, bold: true, color: { argb: NAVY } };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (col === 4 || col === 6) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (col === 5) {
          cell.numFmt = "dd/mm/yyyy";
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }
      r++;
    });
    if (r - 1 > start) ws.mergeCells(`A${start}:A${r - 1}`);
  });

  return wb;
}

app.get("/api/export/excel", async (req, res) => {
  try {
    const codes = await getAllCodes();
    const wb = buildWorkbook(codes);
    res.setHeader("Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition",
      'attachment; filename="liste-informations-documentees.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Erreur export Excel :", err);
    res.status(500).json({ success: false, message: "Erreur export Excel" });
  }
});

// ---------------------------------------------------------------------
//  EXPORT WORD  (même registre, en .docx)
// ---------------------------------------------------------------------
const NAVY_HEX = "1F3864";
const LIGHT_HEX = "D9E1F2";

function cell(text, opts = {}) {
  const {
    bold = false, color = "000000", fill = null, align = AlignmentType.LEFT,
    size = 20, width = null, mono = false, rowSpan = null,
  } = opts;
  return new TableCell({
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    rowSpan: rowSpan || undefined,
    shading: fill ? { type: ShadingType.SOLID, color: fill, fill } : undefined,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({
        text: String(text ?? ""),
        bold, size, color,
        font: mono ? "Consolas" : "Times New Roman",
      })],
    })],
  });
}

function buildWordDoc(codes) {
  const rows = [];

  // Ligne d'en-tête
  rows.push(new TableRow({
    tableHeader: true,
    children: ["Type de document","Intitulé du document","Code","Version","Date","Motif de mise à jour"]
      .map((h, i) => cell(h, {
        bold: true, color: "FFFFFF", fill: NAVY_HEX,
        align: AlignmentType.CENTER, size: 20,
        width: [16, 34, 16, 8, 12, 14][i],
      })),
  }));

  // Lignes groupées par type avec fusion verticale (rowSpan)
  TYPE_ORDER.forEach((t) => {
    const group = codes.filter((c) => c.type === t);
    if (group.length === 0) return;
    group.forEach((item, idx) => {
      const cells = [];
      if (idx === 0) {
        cells.push(cell(TYPE_LABELS[t], {
          bold: true, color: NAVY_HEX, fill: LIGHT_HEX,
          align: AlignmentType.CENTER, rowSpan: group.length, width: 16,
        }));
      }
      cells.push(cell(item.intitule, { width: 34 }));
      cells.push(cell(item.code, { bold: true, color: NAVY_HEX, mono: true, align: AlignmentType.CENTER, width: 16 }));
      cells.push(cell(item.version ?? 0, { align: AlignmentType.CENTER, width: 8 }));
      cells.push(cell(
        new Date(item.createdAt || Date.now()).toLocaleDateString("fr-FR"),
        { align: AlignmentType.CENTER, width: 12 }
      ));
      cells.push(cell(item.motif || "Création", { align: AlignmentType.CENTER, width: 14 }));
      rows.push(new TableRow({ children: cells }));
    });
  });

  const tableBorder = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };

  const doc = new Document({
    sections: [{
      properties: { page: { size: { orientation: "landscape" } } },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "Liste des Informations Documentées",
            bold: true, size: 32, color: NAVY_HEX, font: "Times New Roman",
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({
            text: "Code : ENG-LID-01   |   Version : 00   |   Date : " + new Date().toLocaleDateString("fr-FR"),
            size: 18, color: "555555", font: "Times New Roman",
          })],
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder,
            insideHorizontal: tableBorder, insideVertical: tableBorder,
          },
          rows,
        }),
      ],
    }],
  });

  return doc;
}

app.get("/api/export/word", async (req, res) => {
  try {
    const codes = await getAllCodes();
    const doc = buildWordDoc(codes);
    const buffer = await Packer.toBuffer(doc);
    res.setHeader("Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition",
      'attachment; filename="liste-informations-documentees.docx"');
    res.send(buffer);
  } catch (err) {
    console.error("Erreur export Word :", err);
    res.status(500).json({ success: false, message: "Erreur export Word" });
  }
});

// ---------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`✅ Server running on http://localhost:${PORT}`);
});