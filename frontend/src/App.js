import React, { useState } from "react";
import ExcelJS from "exceljs";
import {
  Document, Packer, Table, TableRow, TableCell, Paragraph, TextRun,
  WidthType, AlignmentType, BorderStyle, VerticalAlign, ShadingType,
} from "docx";

/* =========================================================================
   APPLICATION 100% AUTONOME (sans backend)
   - Aucune donnée n'est enregistrée : tout vit en mémoire.
   - À chaque ouverture / rafraîchissement → tableau vierge, numérotation à 01.
   - Génération des codes + exports Word/Excel faits dans le navigateur.
   - Déployable sur GitHub Pages tel quel.
   ========================================================================= */

const TYPE_LABELS = {
  processus:      { code: "PS",  color: "#0F6E56", label: "Processus" },
  procedure:      { code: "PR",  color: "#993C1D", label: "Procédure" },
  enregistrement: { code: "ENG", color: "#534AB7", label: "Enregistrement" },
  formulaire:     { code: "F",   color: "#185FA5", label: "Formulaire" },
  logigramme:     { code: "L",   color: "#3B6D11", label: "Logigramme" },
  instruction:    { code: "INS", color: "#854F0B", label: "Instruction" },
  modeoperatoire: { code: "MO",  color: "#993556", label: "Mode opératoire" },
};

const styles = {
  root: {
    fontFamily: "'Syne', sans-serif",
    background: "#F8F8F6",
    minHeight: "100vh",
    color: "#2C2C2A",
    margin: "0",
  },
  header: {
    borderBottom: "1px solid #F9F9F7",
    padding: "20px 40px",
    display: "flex",
    alignItems: "center",
    gap: "16px",
    background: "#FFFFFF",
  },
  logo: {
    width: "36px", height: "36px",
    background: "#0F6E56",
    borderRadius: "8px",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: "18px", fontWeight: "800", color: "#FFFFFF",
  },
  headerTitle: { fontSize: "18px", fontWeight: "800", letterSpacing: "-0.5px", color: "#2C2C2A" },
  headerSub:   { fontSize: "12px", color: "#888780", fontFamily: "'IBM Plex Mono', monospace", marginTop: "2px" },
  main: { display: "grid", gridTemplateColumns: "400px 1fr", minHeight: "calc(100vh - 65px)" },
  sidebar: { borderRight: "1px solid #D3D1C7", padding: "32px 28px", background: "#FFFFFF", overflowY: "auto" },
  sideTitle: {
    fontSize: "11px", fontFamily: "'IBM Plex Mono', monospace",
    color: "#888780", letterSpacing: "2px", textTransform: "uppercase", marginBottom: "20px",
  },
  label: {
    display: "block", fontSize: "11px", fontFamily: "'IBM Plex Mono', monospace",
    color: "#5F5E5A", marginBottom: "6px", letterSpacing: "1px",
  },
  fieldGroup: { marginBottom: "20px" },
  select: {
    width: "100%", background: "#F8F8F6", border: "1px solid #D3D1C7",
    borderRadius: "8px", color: "#2C2C2A", fontFamily: "'Syne', sans-serif",
    fontSize: "15px", padding: "13px 14px", outline: "none", cursor: "pointer",
    appearance: "none",
    backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23888780' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E")`,
    backgroundRepeat: "no-repeat", backgroundPosition: "right 14px center",
    transition: "border-color 0.2s",
    boxSizing: "border-box",
  },
  input: {
    width: "100%", background: "#F8F8F6", border: "1px solid #D3D1C7",
    borderRadius: "8px", color: "#2C2C2A", fontFamily: "'Syne', sans-serif",
    fontSize: "16px", padding: "14px 16px", outline: "none",
    boxSizing: "border-box", transition: "border-color 0.2s",
  },
  inputCustomType: {
    width: "100%", background: "#F8F8F6", border: "1px solid #D3D1C7",
    borderRadius: "8px", color: "#2C2C2A", fontFamily: "'Syne', sans-serif",
    fontSize: "16px", padding: "14px 16px", outline: "none",
    boxSizing: "border-box", transition: "border-color 0.2s",
  },
  inputCustomCode: {
    width: "100%", background: "#EAF3DE", border: "2px solid #639922",
    borderRadius: "8px", color: "#3B6D11",
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: "22px", fontWeight: "700",
    padding: "12px 16px", outline: "none",
    boxSizing: "border-box", transition: "border-color 0.2s",
    textTransform: "uppercase", letterSpacing: "4px",
    textAlign: "center",
  },
  preview: {
    background: "#F1EFE8", border: "1px solid #D3D1C7",
    borderRadius: "8px", padding: "16px", marginBottom: "20px",
    fontFamily: "'IBM Plex Mono', monospace",
  },
  previewLabel: { fontSize: "10px", color: "#888780", letterSpacing: "2px", marginBottom: "8px", textTransform: "uppercase" },
  previewCode:  { fontSize: "28px", fontWeight: "700", letterSpacing: "3px" },
  btn: {
    width: "100%", padding: "14px", borderRadius: "8px", border: "none",
    background: "#0F6E56",
    color: "#FFFFFF", fontFamily: "'Syne', sans-serif",
    fontSize: "15px", fontWeight: "800", cursor: "pointer",
    letterSpacing: "0.5px", transition: "opacity 0.2s",
  },
  content: { padding: "36px", overflow: "auto", background: "#F8F8F6" },
  statsRow: { display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "12px", marginBottom: "16px" },
  statCard: { background: "#FFFFFF", border: "1px solid #D3D1C7", borderRadius: "10px", padding: "16px" },
  statLabel: { fontSize: "10px", fontFamily: "'IBM Plex Mono', monospace", color: "#888780", letterSpacing: "2px", marginBottom: "6px" },
  statValue: { fontSize: "28px", fontWeight: "800", letterSpacing: "-1px" },
  statType:  { fontSize: "11px", fontFamily: "'IBM Plex Mono', monospace", color: "#888780", marginTop: "4px" },
  tableWrap: { background: "#FFFFFF", border: "1px solid #D3D1C7", borderRadius: "12px", overflow: "hidden" },
  tableHeader: {
    display: "grid", gridTemplateColumns: "200px 1fr 130px 160px 44px",
    padding: "12px 20px", borderBottom: "1px solid #D3D1C7",
    fontSize: "10px", fontFamily: "'IBM Plex Mono', monospace",
    color: "#888780", letterSpacing: "2px", textTransform: "uppercase",
    background: "#F8F8F6",
  },
  tableRow: {
    display: "grid", gridTemplateColumns: "200px 1fr 130px 160px 44px",
    padding: "14px 20px", borderBottom: "1px solid #F1EFE8",
    alignItems: "center", transition: "background 0.15s",
  },
  codeBadge: { fontFamily: "'IBM Plex Mono', monospace", fontSize: "14px", fontWeight: "600", letterSpacing: "1px" },
  typeBadge: {
    display: "inline-flex", alignItems: "center",
    padding: "3px 8px", borderRadius: "5px",
    fontSize: "10px", fontFamily: "'IBM Plex Mono', monospace",
    fontWeight: "600", letterSpacing: "1px",
  },
  dateText: { fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: "#888780" },
  deleteBtn: {
    background: "none", border: "1px solid #D3D1C7", color: "#888780",
    borderRadius: "5px", width: "28px", height: "28px",
    cursor: "pointer", fontSize: "14px",
    display: "flex", alignItems: "center", justifyContent: "center",
    transition: "all 0.2s",
  },
  exportBtn: {
    display: "flex", alignItems: "center", gap: "8px",
    padding: "9px 18px", borderRadius: "7px",
    border: "1px solid #B5D4F4", background: "#E6F1FB",
    color: "#185FA5", fontFamily: "'Syne', sans-serif",
    fontSize: "13px", fontWeight: "700", cursor: "pointer",
    transition: "all 0.2s", letterSpacing: "0.3px",
  },
  exportBtnExcel: {
    display: "flex", alignItems: "center", gap: "8px",
    padding: "9px 18px", borderRadius: "7px",
    border: "1px solid #A8D08D", background: "#EAF3DE",
    color: "#3B6D11", fontFamily: "'Syne', sans-serif",
    fontSize: "13px", fontWeight: "700", cursor: "pointer",
    transition: "all 0.2s", letterSpacing: "0.3px",
  },
  clearBtn: {
    display: "flex", alignItems: "center", gap: "8px",
    padding: "9px 18px", borderRadius: "7px",
    border: "1px solid #E8B4B4", background: "#FBEAEA",
    color: "#A53030", fontFamily: "'Syne', sans-serif",
    fontSize: "13px", fontWeight: "700", cursor: "pointer",
    transition: "all 0.2s", letterSpacing: "0.3px",
  },
  toast: {
    position: "fixed", bottom: "28px", right: "28px",
    background: "#FFFFFF", border: "1px solid #D3D1C7",
    borderRadius: "10px", padding: "14px 20px",
    fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px",
    color: "#0F6E56", zIndex: 999,
    animation: "slideUp 0.3s ease",
    boxShadow: "0 4px 20px rgba(0,0,0,0.1)",
  },
  empty: {
    textAlign: "center", padding: "56px",
    color: "#B4B2A9", fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px",
  },
  refRow: {
    display: "flex", justifyContent: "space-between",
    marginBottom: "8px", fontSize: "12px",
    padding: "6px 10px", borderRadius: "6px",
    background: "#F8F8F6",
  },
  autreBox: {
    background: "#F1EFE8", border: "1px dashed #B4B2A9",
    borderRadius: "8px", padding: "14px", marginTop: "10px",
  },
  autreBoxLabel: {
    fontSize: "10px", fontFamily: "'IBM Plex Mono', monospace",
    color: "#888780", letterSpacing: "2px", textTransform: "uppercase",
    marginBottom: "12px",
  },
  codeAutoTag: {
    fontSize: "10px", fontFamily: "'IBM Plex Mono', monospace",
    color: "#3B6D11", letterSpacing: "1px",
    marginTop: "6px", textAlign: "center",
  },
};

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
const TYPE_LABELS_PLURAL = {
  PS:  "Processus",
  PR:  "Procédures",
  ENG: "Enregistrements",
  F:   "Formulaires",
  L:   "Logigrammes",
  INS: "Instructions",
  MO:  "Modes opératoires",
};
const TYPE_ORDER = ["PS", "PR", "ENG", "F", "L", "INS", "MO"];

/* =========================================================================
   RÉFÉRENTIEL DOCUMENTAIRE — alimente la liste "Document"
   ========================================================================= */
const DOC_CATALOG = {
  processus: [
    { group: "Processus de Management", items: [
      "Pilotage stratégique", "Management du système intégré", "Gestion des risques et opportunités",
      "Revue de direction", "Veille réglementaire", "Amélioration continue",
    ] },
    { group: "Processus Réalisation", items: [
      "Gestion commerciale", "Gestion des appels d'offres", "Conception et développement", "Achats",
      "Production / Réalisation du service", "Contrôle qualité", "Logistique", "Livraison",
      "Service après-vente", "Gestion des réclamations",
    ] },
    { group: "Processus Support", items: [
      "Ressources humaines", "Formation", "Maintenance", "Système d'information",
      "Finance et comptabilité", "Gestion documentaire", "Métrologie", "Gestion des infrastructures",
    ] },
  ],
  procedure: [
    { group: "Management", items: [
      "Maîtrise des informations documentées", "Gestion des risques et opportunités", "Audit interne",
      "Revue de direction", "Actions correctives", "Traitement des non-conformités", "Veille réglementaire",
      "Communication interne et externe", "Gestion des changements", "Gestion des objectifs",
    ] },
    { group: "Qualité — ISO 9001", items: [
      "Gestion des commandes", "Gestion des réclamations", "Satisfaction client", "Évaluation fournisseurs",
      "Conception et développement", "Contrôle qualité", "Maîtrise des équipements de mesure",
    ] },
    { group: "Environnement — ISO 14001", items: [
      "Identification des aspects environnementaux", "Gestion des déchets", "Gestion des produits chimiques",
      "Gestion des situations d'urgence environnementale", "Contrôle des consommations",
    ] },
    { group: "SST — ISO 45001", items: [
      "Identification des dangers", "Évaluation des risques SST", "Gestion des accidents", "Gestion des EPI",
      "Gestion des entreprises extérieures", "Gestion des permis de travail", "Gestion des situations d'urgence",
    ] },
    { group: "Sécurité alimentaire — ISO 22000", items: [
      "HACCP", "Gestion des CCP", "Gestion des PRP", "Gestion des PRPo", "Traçabilité",
      "Gestion des rappels", "Gestion des allergènes",
    ] },
    { group: "Dispositifs médicaux — ISO 13485", items: [
      "Gestion des risques ISO 14971", "Validation des procédés", "Validation des logiciels",
      "Matériovigilance", "Traçabilité dispositifs médicaux",
    ] },
    { group: "Sécurité de l'information — ISO 27001", items: [
      "Gestion des actifs informationnels", "Gestion des accès", "Gestion des mots de passe",
      "Gestion des sauvegardes", "Gestion des incidents de sécurité", "Gestion des vulnérabilités",
      "Gestion des fournisseurs IT", "Continuité d'activité", "Gestion de la cryptographie",
    ] },
  ],
  instruction: [
    { group: "Instructions de travail", items: [
      "Traitement des commandes", "Contrôle qualité", "Gestion des déchets", "Évacuation d'urgence",
      "Utilisation des EPI", "Nettoyage et désinfection", "Sauvegarde informatique", "Gestion des mots de passe",
      "Gestion des visiteurs", "Produits non conformes", "Surveillance CCP", "Gestion des incidents SSI",
    ] },
  ],
  formulaire: [
    { group: "Formulaires", items: [
      "Fiche de non-conformité", "Fiche d'action corrective", "Fiche d'audit", "Fiche de réclamation client",
      "Fiche d'évaluation fournisseur", "Fiche d'évaluation formation", "Fiche d'analyse des risques",
      "Fiche d'accident de travail", "Fiche de contrôle qualité", "Fiche de contrôle environnemental",
    ] },
  ],
  enregistrement: [
    { group: "Enregistrements", items: [
      "Registre des risques", "Registre des opportunités", "Registre des non-conformités",
      "Registre des actions correctives", "Registre des audits internes", "Registre des réclamations",
      "Registre des accidents", "Registre des déchets", "Registre des formations", "Registre des compétences",
      "Registre des fournisseurs", "Registre des équipements", "Registre des CCP", "Registre de traçabilité",
      "Registre des incidents SSI", "Inventaire des actifs informationnels", "Registre des sauvegardes",
      "Registre des accès utilisateurs",
    ] },
  ],
};

const IGNORE = new Set([
  "DE","DES","ET","LA","LE","LES","DU","D","AU","AUX","UN","UNE",
  "EN","PAR","SUR","POUR","AVEC","DANS","L","A","SE","SA","SES",
  "OU","QUE","QUI","SOUS","VERS","ENTRE","DONT","SON","CAS",
  "PROCESSUS","PROCEDURE","PROCEDURES","ENREGISTREMENT","ENREGISTREMENTS",
  "FORMULAIRE","FORMULAIRES","FICHE","FICHES","LOGIGRAMME","LOGIGRAMMES",
  "INSTRUCTION","INSTRUCTIONS","MODE","OPERATOIRE","OPERATOIRES",
]);

function autoGenerateCode(name) {
  if (!name.trim()) return "";
  const cleaned = name.toUpperCase().normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z\s]/g, "").trim();
  const words = cleaned.split(/\s+/).filter(w => w.length > 0);
  if (words.length === 0) return "";
  if (words.length > 1) return words.map(w => w[0]).join("").slice(0, 3);
  const w = words[0];
  if (w.length <= 4) return w.slice(0, 1);
  if (w.length <= 7) return w.slice(0, 2);
  return w.slice(0, 3);
}

// Abréviation à partir de l'intitulé (même logique que le backend)
function makeAbbrev(intitule) {
  const words = intitule.toUpperCase().normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z\s]/g, " ")
    .split(/\s+/).filter(w => w && !IGNORE.has(w));
  let abbrev = words.map(w => w[0]).join("").slice(0, 5);
  if (abbrev.length === 1 && words[0]) abbrev = words[0].slice(0, 3);
  return abbrev || "DOC";
}

function previewCode(typeKey, customCode, intitule, counters) {
  if (!typeKey || !intitule.trim()) return "_ _ _";
  let typeCode;
  if (typeKey === "__autre__") {
    typeCode = customCode.trim().toUpperCase().slice(0, 3) || "??";
  } else {
    typeCode = TYPE_MAP_BACKEND[typeKey];
  }
  if (!typeCode) return "???";
  const abbrev = makeAbbrev(intitule);
  const next = String((counters[typeCode] || 0) + 1).padStart(2, "0");
  return `${typeCode}-${abbrev}-${next}`;
}

// Téléchargement d'un Blob dans le navigateur
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

// Ordre des types pour les exports (types fixes puis types personnalisés rencontrés)
function orderedTypeCodes(codes) {
  const extras = [];
  for (const c of codes) {
    if (!TYPE_ORDER.includes(c.type) && !extras.includes(c.type)) extras.push(c.type);
  }
  return [...TYPE_ORDER, ...extras];
}
function pluralLabel(typeCode, group) {
  return TYPE_LABELS_PLURAL[typeCode] || (group[0] && group[0].typeLabel) || typeCode;
}

/* =========================================================================
   EXPORT EXCEL (navigateur, via ExcelJS) — registre branché OCG
   ========================================================================= */
const NAVY  = "FF1F3864";
const LIGHT = "FFD9E1F2";
const XFONT = "Times New Roman";

async function exportExcel(codes) {
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

  let r = HR + 1;
  orderedTypeCodes(codes).forEach((t) => {
    const group = codes.filter((c) => c.type === t).slice().sort((a, b) => (a.seq || 0) - (b.seq || 0));
    if (group.length === 0) return;
    const start = r;
    group.forEach((item) => {
      const row = ws.getRow(r);
      row.getCell(1).value = pluralLabel(t, group);
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

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  downloadBlob(blob, "liste-informations-documentees.xlsx");
}

/* =========================================================================
   EXPORT WORD (navigateur, via docx)
   ========================================================================= */
const NAVY_HEX = "1F3864";
const LIGHT_HEX = "D9E1F2";

function wcell(text, opts = {}) {
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

async function exportWord(codes) {
  const rows = [];

  rows.push(new TableRow({
    tableHeader: true,
    children: ["Type de document","Intitulé du document","Code","Version","Date","Motif de mise à jour"]
      .map((h, i) => wcell(h, {
        bold: true, color: "FFFFFF", fill: NAVY_HEX,
        align: AlignmentType.CENTER, size: 20,
        width: [16, 34, 16, 8, 12, 14][i],
      })),
  }));

  orderedTypeCodes(codes).forEach((t) => {
    const group = codes.filter((c) => c.type === t).slice().sort((a, b) => (a.seq || 0) - (b.seq || 0));
    if (group.length === 0) return;
    group.forEach((item, idx) => {
      const cells = [];
      if (idx === 0) {
        cells.push(wcell(pluralLabel(t, group), {
          bold: true, color: NAVY_HEX, fill: LIGHT_HEX,
          align: AlignmentType.CENTER, rowSpan: group.length, width: 16,
        }));
      }
      cells.push(wcell(item.intitule, { width: 34 }));
      cells.push(wcell(item.code, { bold: true, color: NAVY_HEX, mono: true, align: AlignmentType.CENTER, width: 16 }));
      cells.push(wcell(item.version ?? 0, { align: AlignmentType.CENTER, width: 8 }));
      cells.push(wcell(
        new Date(item.createdAt || Date.now()).toLocaleDateString("fr-FR"),
        { align: AlignmentType.CENTER, width: 12 }
      ));
      cells.push(wcell(item.motif || "Création", { align: AlignmentType.CENTER, width: 14 }));
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

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, "liste-informations-documentees.docx");
}

/* =========================================================================
   APPLICATION
   ========================================================================= */
export default function App() {
  const [type, setType]             = useState("");
  const [customType, setCustomType] = useState("");
  const [customCode, setCustomCode] = useState("");
  const [docModel, setDocModel]     = useState("");
  const [intitule, setIntitule]     = useState("");
  const [codes, setCodes]           = useState([]);     // en mémoire uniquement
  const [counters, setCounters]     = useState({ PS:0, PR:0, ENG:0, F:0, L:0, INS:0, MO:0 });
  const [toast, setToast]           = useState(null);
  const [hover, setHover]           = useState(null);

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 3000); };

  const handleSubmit = () => {
    if (!type) { showToast("⚠️ Veuillez sélectionner un type"); return; }
    if (type === "__autre__" && (!customType.trim() || !customCode.trim())) {
      showToast("⚠️ Veuillez saisir le nom et le code du type");
      return;
    }
    if (!intitule.trim()) { showToast("⚠️ Veuillez saisir un intitulé"); return; }

    const typeCode = type === "__autre__"
      ? customCode.trim().toUpperCase().slice(0, 3)
      : TYPE_MAP_BACKEND[type];
    const typeLabel = type === "__autre__"
      ? customType.trim()
      : TYPE_LABELS_PLURAL[typeCode];

    const abbrev = makeAbbrev(intitule);
    const seq = (counters[typeCode] || 0) + 1;
    const code = `${typeCode}-${abbrev}-${String(seq).padStart(2, "0")}`;

    const entry = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      code,
      intitule: intitule.trim(),
      type: typeCode,
      typeLabel,
      seq,
      version: 0,
      motif: "Création",
      createdAt: new Date().toISOString(),
    };

    setCodes((prev) => [entry, ...prev]);
    setCounters((prev) => ({ ...prev, [typeCode]: seq }));
    showToast(`✅ Code généré : ${code}`);

    setIntitule(""); setType(""); setCustomType(""); setCustomCode(""); setDocModel("");
  };

  const handleDelete = (id) => {
    setCodes((prev) => prev.filter((c) => c.id !== id));
    showToast("🗑️ Code supprimé");
  };

  const handleClearAll = () => {
    if (codes.length === 0) { showToast("⚠️ Le tableau est déjà vide"); return; }
    if (!window.confirm(`Supprimer les ${codes.length} codes et réinitialiser les compteurs ?`)) return;
    setCodes([]);
    setCounters({ PS:0, PR:0, ENG:0, F:0, L:0, INS:0, MO:0 });
    showToast("🗑️ Tableau vidé — compteurs réinitialisés");
  };

  const handleExportWord = async () => {
    if (codes.length === 0) { showToast("⚠️ Aucun code à exporter"); return; }
    try { await exportWord(codes); showToast("📄 Word téléchargé"); }
    catch (e) { console.error(e); showToast("❌ Erreur export Word"); }
  };

  const handleExportExcel = async () => {
    if (codes.length === 0) { showToast("⚠️ Aucun code à exporter"); return; }
    try { await exportExcel(codes); showToast("📊 Excel téléchargé"); }
    catch (e) { console.error(e); showToast("❌ Erreur export Excel"); }
  };

  const preview = previewCode(type, customCode, intitule, counters);
  const selectedColor = type && type !== "__autre__" ? TYPE_LABELS[type]?.color : "#3B6D11";

  return (
    <div style={styles.root}>
      <style>{`
        * { box-sizing: border-box; }
        select:focus, input:focus { border-color: #0F6E56 !important; outline: none; }
        @keyframes slideUp { from { transform:translateY(16px); opacity:0 } to { transform:translateY(0); opacity:1 } }
        ::-webkit-scrollbar { width:5px; }
        ::-webkit-scrollbar-track { background:#F8F8F6; }
        ::-webkit-scrollbar-thumb { background:#D3D1C7; border-radius:3px; }
        button:hover { opacity: 0.85; }
      `}</style>

      {/* Header */}
      <div style={styles.header}>
        <div style={styles.logo}>D</div>
        <div>
          <div style={styles.headerTitle}>DocCode Generator</div>
          <div style={styles.headerSub}>Système de codification documentaire — QMS/ISO</div>
        </div>
      </div>

      <div style={styles.main}>
        {/* Sidebar */}
        <div style={styles.sidebar}>
          <div style={styles.sideTitle}>// Nouveau document</div>

          <div style={styles.fieldGroup}>
            <label style={styles.label}>TYPE DE DOCUMENT</label>
            <select
              style={styles.select}
              value={type}
              onChange={e => {
                setType(e.target.value);
                setCustomType(""); setCustomCode(""); setDocModel("");
              }}
            >
              <option value="">-- Sélectionner --</option>
              <option value="processus">Processus (PS)</option>
              <option value="procedure">Procédure (PR)</option>
              <option value="enregistrement">Enregistrement (ENG)</option>
              <option value="formulaire">Formulaire (F)</option>
              <option value="logigramme">Logigramme (L)</option>
              <option value="instruction">Instruction (INS)</option>
              <option value="modeoperatoire">Mode opératoire (MO)</option>
              <option value="__autre__">✏️ Autre type (personnalisé)...</option>
            </select>

            {type === "__autre__" && (
              <div style={styles.autreBox}>
                <div style={styles.autreBoxLabel}>Définir un nouveau type</div>

                <div style={{ marginBottom: "12px" }}>
                  <label style={{ ...styles.label, marginBottom: "5px" }}>NOM DU TYPE</label>
                  <input
                    style={styles.inputCustomType}
                    placeholder="ex: Rapport d'audit"
                    value={customType}
                    onChange={e => {
                      const val = e.target.value;
                      setCustomType(val);
                      setCustomCode(autoGenerateCode(val));
                    }}
                  />
                </div>

                <div>
                  <label style={{ ...styles.label, marginBottom: "5px" }}>CODE (max 3 lettres)</label>
                  <input
                    style={styles.inputCustomCode}
                    maxLength={3}
                    value={customCode}
                    onChange={e => setCustomCode(e.target.value.toUpperCase().replace(/[^A-Z]/g, ""))}
                  />
                  <div style={styles.codeAutoTag}>↑ généré automatiquement — modifiable</div>
                </div>
              </div>
            )}
          </div>

          {/* Liste des documents du référentiel (selon le type choisi) */}
          {type && type !== "__autre__" && DOC_CATALOG[type] && (
            <div style={styles.fieldGroup}>
              <label style={styles.label}>DOCUMENT (depuis le référentiel)</label>
              <select
                style={styles.select}
                value={docModel}
                onChange={e => {
                  const v = e.target.value;
                  setDocModel(v);
                  if (v === "__autre__") setIntitule("");
                  else if (v !== "") setIntitule(v);
                }}
              >
                <option value="">-- Choisir un document --</option>
                {DOC_CATALOG[type].map((grp, gi) => (
                  <optgroup key={gi} label={grp.group}>
                    {grp.items.map((it, ii) => (
                      <option key={ii} value={it}>{it}</option>
                    ))}
                  </optgroup>
                ))}
                <option value="__autre__">✏️ Autre (saisie libre)…</option>
              </select>
            </div>
          )}

          <div style={styles.fieldGroup}>
            <label style={styles.label}>INTITULÉ DU DOCUMENT</label>
            <input
              style={styles.input}
              placeholder="ex: Gestion des non conformités"
              value={intitule}
              onChange={e => setIntitule(e.target.value)}
              onKeyDown={e => e.key === "Enter" && handleSubmit()}
            />
          </div>

          <div style={styles.preview}>
            <div style={styles.previewLabel}>Aperçu du code</div>
            <div style={{ ...styles.previewCode, color: selectedColor }}>{preview}</div>
          </div>

          <button style={styles.btn} onClick={handleSubmit}>
            → Générer le code
          </button>

          <div style={{ marginTop: "32px" }}>
            <div style={styles.sideTitle}>// Référence</div>
            {Object.entries(TYPE_LABELS).map(([key, val]) => (
              <div key={key} style={styles.refRow}>
                <span style={{ color: "#5F5E5A" }}>{val.label}</span>
                <span style={{ fontFamily:"'IBM Plex Mono', monospace", color: val.color, fontSize:"11px", fontWeight:"700" }}>{val.code}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Contenu principal */}
        <div style={styles.content}>

          {/* Stats rangée 1 */}
          <div style={styles.statsRow}>
            {["processus","procedure","enregistrement","formulaire"].map(key => {
              const val = TYPE_LABELS[key];
              return (
                <div key={key} style={styles.statCard}>
                  <div style={styles.statLabel}>TOTAL</div>
                  <div style={{ ...styles.statValue, color: val.color }}>
                    {String(counters[val.code] || 0).padStart(2, "0")}
                  </div>
                  <div style={styles.statType}>{val.label}s</div>
                </div>
              );
            })}
          </div>

          {/* Stats rangée 2 */}
          <div style={{ ...styles.statsRow, gridTemplateColumns: "repeat(3, 1fr)", marginBottom: "28px" }}>
            {["logigramme","instruction","modeoperatoire"].map(key => {
              const val = TYPE_LABELS[key];
              return (
                <div key={key} style={{ ...styles.statCard, padding: "12px 16px" }}>
                  <div style={styles.statLabel}>TOTAL</div>
                  <div style={{ ...styles.statValue, fontSize: "22px", color: val.color }}>
                    {String(counters[val.code] || 0).padStart(2, "0")}
                  </div>
                  <div style={styles.statType}>{val.label}</div>
                </div>
              );
            })}
          </div>

          {/* Barre tableau */}
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"16px" }}>
            <div style={{ fontSize:"12px", fontFamily:"'IBM Plex Mono', monospace", color:"#888780" }}>
              {codes.length} code{codes.length !== 1 ? "s" : ""} enregistré{codes.length !== 1 ? "s" : ""}
            </div>
            <div style={{ display: "flex", gap: "10px" }}>
              <button style={styles.exportBtn} onClick={handleExportWord}>
                📄 Exporter Word
              </button>
              <button style={styles.exportBtnExcel} onClick={handleExportExcel}>
                📊 Exporter Excel
              </button>
              <button style={styles.clearBtn} onClick={handleClearAll}>
                🗑️ Tout effacer
              </button>
            </div>
          </div>

          {/* Tableau */}
          <div style={styles.tableWrap}>
            <div style={styles.tableHeader}>
              <span>Code</span><span>Intitulé</span><span>Type</span><span>Date</span><span></span>
            </div>

            {codes.length === 0 && (
              <div style={styles.empty}>Aucun code généré — commencez par créer un document</div>
            )}

            {codes.map((c) => {
              const typeInfo = Object.values(TYPE_LABELS).find(t => t.code === c.type);
              const color = typeInfo?.color || "#888780";
              return (
                <div
                  key={c.id}
                  style={{ ...styles.tableRow, background: hover === c.id ? "#F8F8F6" : "transparent" }}
                  onMouseEnter={() => setHover(c.id)}
                  onMouseLeave={() => setHover(null)}
                >
                  <span style={{ ...styles.codeBadge, color }}>{c.code}</span>
                  <span style={{ fontSize:"13px", color:"#444441" }}>{c.intitule}</span>
                  <span>
                    <span style={{ ...styles.typeBadge, background:`${color}18`, color }}>
                      {c.type}
                    </span>
                  </span>
                  <span style={styles.dateText}>
                    {new Date(c.createdAt).toLocaleString("fr-FR", { dateStyle:"short", timeStyle:"short" })}
                  </span>
                  <span>
                    <button style={styles.deleteBtn} onClick={() => handleDelete(c.id)} title="Supprimer">×</button>
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {toast && <div style={styles.toast}>{toast}</div>}
    </div>
  );
}