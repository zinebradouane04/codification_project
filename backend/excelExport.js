const ExcelJS = require("exceljs");

/* Libellé "Type de document" affiché dans la colonne A (pluriel, façon modèle OCG) */
const CAT_LABEL = {
  processus: "Processus",
  procedure: "Procédures",
  instruction: "Instructions",
  formulaire: "Formulaires",
  enregistrement: "Enregistrements",
  plan: "Plans",
  registre: "Registres",
};

const NAVY = "FF1F3864";
const LIGHT = "FFD9E1F2";
const WHITE = "FFFFFFFF";
const TNR = "Times New Roman";
const THIN = { style: "thin", color: { argb: "FF000000" } };
const BORDER = { top: THIN, left: THIN, right: THIN, bottom: THIN };

/**
 * Construit le classeur "Liste des Informations Documentées".
 * @param {Array} entries  [{ code, intitule, cat, createdAt, ... }]
 * @param {Object} meta     { code, version, date }  (en-tête)
 */
function buildWorkbook(entries, meta = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "OCG Consulting";
  const ws = wb.addWorksheet("Informations documentées", {
    views: [{ showGridLines: false }],
  });

  // Largeurs de colonnes (identiques au modèle)
  ws.columns = [
    { width: 22 }, { width: 50 }, { width: 18 },
    { width: 10 }, { width: 13 }, { width: 24 },
  ];

  // ---- Bloc en-tête ----
  ws.mergeCells("A1:B3");
  ws.mergeCells("C1:D3");
  ws.mergeCells("E1:F1");
  ws.mergeCells("E2:F2");
  ws.mergeCells("E3:F3");

  const logo = ws.getCell("A1");
  logo.value = "OCG\nCONSULTING";
  logo.font = { name: TNR, size: 14, bold: true, color: { argb: WHITE } };
  logo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  logo.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  const title = ws.getCell("C1");
  title.value = "Liste des Informations Documentées";
  title.font = { name: TNR, size: 13, bold: true, color: { argb: NAVY } };
  title.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  const today = new Date().toLocaleDateString("fr-FR");
  [
    ["E1", `Code : ${meta.code || "ENG-LID-01"}`],
    ["E2", `Version : ${meta.version || "00"}`],
    ["E3", `Date : ${meta.date || today}`],
  ].forEach(([ref, val]) => {
    const c = ws.getCell(ref);
    c.value = val;
    c.font = { name: TNR, size: 10 };
    c.alignment = { horizontal: "left", vertical: "middle" };
  });

  // Bordures du bloc en-tête (A1:F3)
  for (let r = 1; r <= 3; r++)
    for (let col = 1; col <= 6; col++) ws.getCell(r, col).border = BORDER;

  // ---- Ligne d'en-tête du tableau (ligne 5) ----
  const headers = ["Type de document", "Intitulé du document", "Code", "Version", "Date", "Motif de mise à jour"];
  const hRow = ws.getRow(5);
  hRow.height = 26;
  headers.forEach((h, i) => {
    const c = hRow.getCell(i + 1);
    c.value = h;
    c.font = { name: TNR, size: 11, bold: true, color: { argb: WHITE } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = BORDER;
  });

  // ---- Lignes de données ----
  entries.forEach((e, idx) => {
    const r = ws.getRow(6 + idx);
    const date = e.createdAt ? new Date(e.createdAt) : new Date();
    const cells = [
      CAT_LABEL[e.cat] || e.famLabel || e.cat || "",
      e.intitule || "",
      e.code || "",
      typeof e.version === "number" ? e.version : 0,
      date,
      e.motif || "Création",
    ];
    cells.forEach((val, i) => {
      const c = r.getCell(i + 1);
      c.value = val;
      c.border = BORDER;
      c.font = { name: TNR, size: 11, color: { argb: i === 0 ? NAVY : "FF000000" }, bold: i === 0 };
      if (i === 0) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT } };
      c.alignment = {
        horizontal: i === 1 ? "left" : "center",
        vertical: "middle",
        wrapText: i === 1 || i === 5,
      };
      if (i === 4) c.numFmt = "dd/mm/yyyy hh:mm";
    });
  });

  return wb;
}

module.exports = { buildWorkbook };

/* ----------------------------------------------------------------------
   ROUTE EXPRESS (à mettre dans server.js) :

   const { buildWorkbook } = require("./excelExport");

   app.get("/api/export/excel", async (req, res) => {
     const wb = buildWorkbook(codes);   // `codes` = vos enregistrements
     res.setHeader("Content-Type",
       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
     res.setHeader("Content-Disposition",
       'attachment; filename="liste-informations-documentees.xlsx"');
     await wb.xlsx.write(res);
     res.end();
   });
   ---------------------------------------------------------------------- */