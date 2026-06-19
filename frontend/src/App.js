import React, { useState, useEffect, useRef } from "react";
import ExcelJS from "exceljs";
import brandLogo from "./logo_codoc.png";
import logoOcg from "./logo_ocg.png";
import {
  Document, Packer, Table, TableRow, TableCell, Paragraph, TextRun, ImageRun,
  WidthType, AlignmentType, BorderStyle, VerticalAlign, ShadingType,
} from "docx";

/* =========================================================================
   LOGO DE MARQUE (carré en haut de la sidebar)
   ⬇️ Collez ici VOTRE image PNG. Deux options :
   • Option A (recommandée, marche partout y compris GitHub Pages) :
       une data URI base64 → "data:image/png;base64,iVBORw0KGgoAAAANS...."
   • Option B (fichier dans le projet) :
       en haut du fichier : import brandLogo from "./logo.png";
       puis ci-dessous :   const BRAND_LOGO = brandLogo;
   Laissez à null pour garder la lettre « D ».
   ========================================================================= */
const BRAND_LOGO = brandLogo; // ex : "data:image/png;base64,iVBORw0KGgo...."

/* =========================================================================
   APPLICATION 100% AUTONOME (sans backend)
   - Multi-profils : écran de connexion / création de compte.
     ⚠ Sécurité LOCALE uniquement (localStorage du navigateur) — adapté à un
       usage interne. Pas de vraie protection cryptographique ni de partage
       entre postes. Pour du multi-utilisateur réel → backend requis.
   - Données séparées par utilisateur : tableau, compteurs et historique
     sont stockés et rechargés par compte (clé = email).
   - Logo d'entreprise par compte : importé à l'inscription ou modifié après
     connexion (clic sur l'avatar). Affiché en-tête, accueil + exports.
   - À la 1re connexion d'un compte → programme vierge + « Bonjour, [entreprise] ».
   - Exports Word / Excel faits dans le navigateur. Déployable sur GitHub Pages.
   ========================================================================= */

/* ----------------------- AUTHENTIFICATION (locale) ----------------------- */
const ACCOUNTS_KEY = "doccode_accounts_v1";
const SESSION_KEY  = "doccode_session_v1";

function loadAccounts() {
  try { return JSON.parse(localStorage.getItem(ACCOUNTS_KEY)) || {}; }
  catch (e) { return {}; }
}
function saveAccounts(acc) {
  try { localStorage.setItem(ACCOUNTS_KEY, JSON.stringify(acc)); } catch (e) {}
}
function loadSession() {
  try { return localStorage.getItem(SESSION_KEY) || null; } catch (e) { return null; }
}
function saveSession(email) { try { localStorage.setItem(SESSION_KEY, email); } catch (e) {} }
function clearSession()     { try { localStorage.removeItem(SESSION_KEY); } catch (e) {} }

// Hachage simple — NON cryptographique, juste pour éviter le mot de passe en clair.
function hashPassword(pw) {
  let h = 0;
  for (let i = 0; i < pw.length; i++) h = (Math.imul(31, h) + pw.charCodeAt(i)) | 0;
  return "h" + (h >>> 0).toString(36);
}

/* ------------------ GESTION DU LOGO (image → base64 PNG) ----------------- */
// Lit un fichier image, le redimensionne (max `maxSize` px, côté le plus long)
// et renvoie un data URL PNG + ses dimensions. Garde le localStorage léger.
function readImageFile(file, maxSize = 256) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const ratio = Math.min(maxSize / img.width, maxSize / img.height, 1);
        const w = Math.max(1, Math.round(img.width * ratio));
        const h = Math.max(1, Math.round(img.height * ratio));
        const canvas = document.createElement("canvas");
        canvas.width = w; canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, w, h);
        try {
          resolve({ dataUrl: canvas.toDataURL("image/png"), width: w, height: h });
        } catch (e) { reject(e); }
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// Extrait la base64 brute d'un data URL PNG.
function dataUrlToBase64(dataUrl) {
  if (!dataUrl || typeof dataUrl !== "string") return null;
  const i = dataUrl.indexOf("base64,");
  return i === -1 ? null : dataUrl.slice(i + 7);
}
// base64 → Uint8Array (pour docx ImageRun).
function base64ToUint8(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
// Met une image (w,h) à l'échelle dans une boîte (maxW,maxH) sans déformation.
function scaleDims(w, h, maxW, maxH) {
  if (!w || !h) return { width: maxW, height: maxH };
  const r = Math.min(maxW / w, maxH / h, 1);
  return { width: Math.max(1, Math.round(w * r)), height: Math.max(1, Math.round(h * r)) };
}

/* ------------------ DONNÉES PAR UTILISATEUR (localStorage) ---------------- */
const stateKey = (email) => `doccode_state_v1__${email}`;
const histKey  = (email) => `doccode_history_v1__${email}`;
const snapKey  = (email) => `doccode_snapshots_v1__${email}`;

function loadUserState(email) {
  try {
    const raw = localStorage.getItem(stateKey(email));
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return null;
}
function loadUserHistory(email) {
  try {
    const raw = localStorage.getItem(histKey(email));
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return [];
}
function loadUserSnapshots(email) {
  try {
    const raw = localStorage.getItem(snapKey(email));
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return [];
}

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
    padding: "16px 40px",
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
  userName:  { fontSize: "13px", fontWeight: "700", color: "#2C2C2A", textAlign: "right" },
  userEmail: { fontSize: "11px", color: "#888780", fontFamily: "'IBM Plex Mono', monospace", textAlign: "right", marginTop: "1px" },
  avatar: {
    width: "46px", height: "46px", borderRadius: "50%",
    background: "#0F6E56", color: "#FFFFFF",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontWeight: "800", fontSize: "14px", flexShrink: 0,
    overflow: "hidden", cursor: "pointer",
    border: "1px solid #E2E0D7",
  },
  logoutBtn: {
    padding: "9px 15px", borderRadius: "7px",
    border: "1px solid #E8B4B4", background: "#FBEAEA", color: "#A53030",
    fontFamily: "'Syne', sans-serif", fontSize: "12px", fontWeight: "700",
    cursor: "pointer", whiteSpace: "nowrap",
  },
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
  welcomeBanner: {
    background: "linear-gradient(120deg, #0F6E56 0%, #0B4F3E 100%)",
    color: "#FFFFFF", borderRadius: "12px", padding: "18px 24px",
    marginBottom: "24px", display: "flex", alignItems: "center",
    justifyContent: "space-between", flexWrap: "wrap", gap: "8px",
  },
  welcomeTitle: { fontSize: "20px", fontWeight: "800", letterSpacing: "-0.5px" },
  welcomeSub: { fontSize: "12px", fontFamily: "'IBM Plex Mono', monospace", opacity: 0.85, marginTop: "3px" },
  bannerLogo: {
    width: "72px", height: "72px", borderRadius: "14px", background: "#FFFFFF",
    display: "flex", alignItems: "center", justifyContent: "center",
    padding: "7px", flexShrink: 0, boxSizing: "border-box",
    boxShadow: "0 2px 10px rgba(0,0,0,0.12)",
  },
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
  histBtn: {
    display: "flex", alignItems: "center", gap: "8px",
    padding: "9px 18px", borderRadius: "7px",
    border: "1px solid #C9C4DF", background: "#EFEDF7",
    color: "#534AB7", fontFamily: "'Syne', sans-serif",
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

  /* --- Fenêtres modales (historique / paramètres logo) --- */
  modalOverlay: {
    position: "fixed", inset: 0, background: "rgba(44,44,42,0.45)",
    display: "flex", alignItems: "center", justifyContent: "center",
    zIndex: 1000, padding: "24px",
  },
  modal: {
    background: "#FFFFFF", borderRadius: "14px", width: "100%", maxWidth: "820px",
    maxHeight: "82vh", display: "flex", flexDirection: "column",
    boxShadow: "0 12px 48px rgba(0,0,0,0.22)", overflow: "hidden",
  },
  modalHead: {
    display: "flex", justifyContent: "space-between", alignItems: "center",
    padding: "20px 24px", borderBottom: "1px solid #D3D1C7", background: "#F8F8F6",
  },
  modalTitle: { fontSize: "16px", fontWeight: "800", color: "#2C2C2A", letterSpacing: "-0.3px" },
  modalSub: { fontSize: "11px", fontFamily: "'IBM Plex Mono', monospace", color: "#888780", marginTop: "3px" },
  closeX: {
    background: "none", border: "1px solid #D3D1C7", borderRadius: "6px",
    width: "32px", height: "32px", cursor: "pointer", fontSize: "16px", color: "#888780",
    display: "flex", alignItems: "center", justifyContent: "center",
  },
  modalBody: { overflowY: "auto", flex: 1 },
  histHeaderRow: {
    display: "grid", gridTemplateColumns: "190px 1fr 90px 150px",
    padding: "10px 24px", borderBottom: "1px solid #D3D1C7",
    fontSize: "10px", fontFamily: "'IBM Plex Mono', monospace",
    color: "#888780", letterSpacing: "2px", textTransform: "uppercase",
    background: "#FCFCFB", position: "sticky", top: 0,
  },
  histRow: {
    display: "grid", gridTemplateColumns: "190px 1fr 90px 150px",
    padding: "12px 24px", borderBottom: "1px solid #F1EFE8", alignItems: "center",
  },
  modalFoot: {
    display: "flex", justifyContent: "space-between", alignItems: "center",
    padding: "14px 24px", borderTop: "1px solid #D3D1C7", background: "#F8F8F6",
  },

  /* --- Fenêtre Logo --- */
  logoPreviewBox: {
    width: "180px", height: "180px", margin: "0 auto 20px",
    borderRadius: "14px", border: "1px solid #D3D1C7", background: "#F8F8F6",
    display: "flex", alignItems: "center", justifyContent: "center",
    overflow: "hidden", padding: "14px", boxSizing: "border-box",
  },
  logoUploadBtn: {
    display: "inline-flex", alignItems: "center", justifyContent: "center", gap: "8px",
    padding: "12px 22px", borderRadius: "9px", cursor: "pointer",
    border: "1px solid #0F6E56", background: "#0F6E56", color: "#FFFFFF",
    fontFamily: "'Syne', sans-serif", fontSize: "14px", fontWeight: "800",
  },
  logoRemoveBtn: {
    display: "inline-flex", alignItems: "center", justifyContent: "center", gap: "8px",
    padding: "12px 22px", borderRadius: "9px", cursor: "pointer",
    border: "1px solid #E8B4B4", background: "#FBEAEA", color: "#A53030",
    fontFamily: "'Syne', sans-serif", fontSize: "14px", fontWeight: "700",
    marginLeft: "10px",
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
    { group: "Catalogue OCG", items: [
      "Management stratégique", "Pilotage stratégique", "Pilotage SMQ", "Pilotage SMI", "Pilotage SMSDA",
      "Management de la qualité", "Achats", "Achats et approvisionnement", "Production", "Fabrication",
      "Vente", "Comptabilité", "Comptabilité et finance", "Logistique et stockage", "Maintenance",
      "Ressources Humaines", "Système d'information", "Marketing", "Management de l'environnement",
      "Santé et sécurité au travail", "Commercial", "Commercial et gestion des commandes",
      "Contrôle qualité", "Contrôle qualité et analyse laboratoire", "Amélioration du SMQ",
      "Amélioration SMI", "Amélioration SMSDA", "Conditionnement et gestion des stocks",
      "Gestion de l'entreprise et développement", "Gestion IT",
    ] },
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
    { group: "Catalogue OCG", items: [
      "Revue de direction", "Audit interne", "Non-conformité et action corrective", "Amélioration continue",
      "Communication interne et externe", "Gestion de changement", "Gestion des risques et opportunités",
      "Gestion des risques SST", "Gestion des aspects et impacts environnementaux", "Gestion des objectifs et KPI",
      "Commercial et gestion des commandes", "Achats", "Sélection et évaluation des fournisseurs", "Vente",
      "Comptabilité", "Logistique et stockage", "Conditionnement", "Gestion IT", "Retrait / rappel",
      "Gestion des situations d'urgence", "Maîtrise des informations documentées",
      "Gestion des compétences et de la formation", "Maintenance des équipements",
      "Gestion des ressources humaines", "Achats et évaluation des fournisseurs", "Contrôle qualité",
      "Gestion des réclamations clients", "Mesure de la satisfaction client", "Gestion des produits non conformes",
      "Nettoyage et désinfection", "Lutte contre les nuisibles", "Hygiène du personnel", "Gestion des déchets",
      "Stockage et transport", "Prévention des contaminations croisées", "Traçabilité",
    ] },
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
    { group: "Catalogue OCG", items: [
      "Fiche de non-conformité", "Fiche de réclamation client", "Fiche d'évaluation des fournisseurs",
      "Fiche de contrôle à la réception", "Fiche de contrôle de l'efficacité du nettoyage", "Fiche d'étalonnage",
      "Fiche de présence à la formation", "Fiche d'évaluation de la formation", "Bon de commande",
      "Bon de livraison", "Bon de réception", "Ordre de fabrication", "Fiche de gestion de changement",
      "Questionnaire de satisfaction client", "Bon de sortie", "Test de traçabilité",
    ] },
    { group: "Formulaires", items: [
      "Fiche de non-conformité", "Fiche d'action corrective", "Fiche d'audit", "Fiche de réclamation client",
      "Fiche d'évaluation fournisseur", "Fiche d'évaluation formation", "Fiche d'analyse des risques",
      "Fiche d'accident de travail", "Fiche de contrôle qualité", "Fiche de contrôle environnemental",
    ] },
  ],
  enregistrement: [
    { group: "Catalogue OCG", items: [
      "Liste des informations documentées", "Cartographie des processus", "Politique qualité", "Politique SMI",
      "Politique SMSDA", "Tableau des parties intéressées", "Registre des risques et opportunités",
      "Registre des réclamations client", "Rapport d'audit interne", "Compte rendu de revue de direction",
      "Registre des non-conformités et actions correctives", "Registre des retraits / rappels de produits",
      "Plan de communication", "Programme d'audit interne", "Planning audit interne", "Plan d'action",
      "Tableau de bord KPI", "Plan HACCP", "Check-list des PRP", "Plan de formation", "Plan de maintenance",
    ] },
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

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

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
   EXPORT EXCEL (navigateur, via ExcelJS) — registre branché par entreprise
   ========================================================================= */
const NAVY  = "FF1F3864";
const LIGHT = "FFD9E1F2";
const XFONT = "Times New Roman";

async function exportExcel(codes, branding = {}) {
  const company = (branding.company || "OCG CONSULTING").trim();
  const wb = new ExcelJS.Workbook();
  wb.creator = `${company} — DocCode Generator`;
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

  // --- Bloc logo (A1:B3) ---
  ws.mergeCells("A1:B3");
  const logo = ws.getCell("A1");
  ws.getRow(1).height = 22; ws.getRow(2).height = 22; ws.getRow(3).height = 22;

  if (branding.logo && dataUrlToBase64(branding.logo)) {
    // Logo image sur fond blanc, ancré en haut-gauche du bloc fusionné.
    const imgId = wb.addImage({ base64: branding.logo, extension: "png" });
    const { width, height } = scaleDims(branding.logoW, branding.logoH, 150, 56);
    logo.value = "";
    logo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
    ws.addImage(imgId, {
      tl: { col: 0.15, row: 0.15 },
      ext: { width, height },
      editAs: "oneCell",
    });
  } else {
    // Pas de logo : nom de l'entreprise sur fond navy.
    logo.value = company.toUpperCase();
    logo.font = { name: XFONT, bold: true, size: 13, color: { argb: "FFFFFFFF" } };
    logo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  }
  logo.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  // --- Titre (C1:D3) avec nom de l'entreprise en sous-titre ---
  ws.mergeCells("C1:D3");
  const titre = ws.getCell("C1");
  titre.value = {
    richText: [
      { text: "Liste des Informations Documentées",
        font: { name: XFONT, bold: true, size: 13, color: { argb: NAVY } } },
      { text: "\n" + company,
        font: { name: XFONT, italic: true, size: 10, color: { argb: "FF555555" } } },
    ],
  };
  titre.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  ws.mergeCells("E1:F1"); ws.getCell("E1").value = "Code : ENG-LID-01";
  ws.mergeCells("E2:F2"); ws.getCell("E2").value = "Version : 00";
  ws.mergeCells("E3:F3");
  ws.getCell("E3").value = "Date : " + new Date().toLocaleDateString("fr-FR");
  ["E1", "E2", "E3"].forEach((c) => {
    ws.getCell(c).font = { name: XFONT, size: 10 };
    ws.getCell(c).alignment = { horizontal: "left", vertical: "middle" };
  });
  for (let r0 = 1; r0 <= 3; r0++)
    for (let c0 = 1; c0 <= 6; c0++) ws.getCell(r0, c0).border = border;

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
   EXPORT WORD (navigateur, via docx) — en-tête branché par entreprise
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

async function exportWord(codes, branding = {}) {
  const company = (branding.company || "OCG CONSULTING").trim();
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

  // En-tête du document : logo (si présent) + nom entreprise + titre + ligne code
  const headerChildren = [];

  const b64 = dataUrlToBase64(branding.logo);
  if (b64) {
    const { width, height } = scaleDims(branding.logoW, branding.logoH, 150, 75);
    headerChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new ImageRun({
        type: "png",
        data: base64ToUint8(b64),
        transformation: { width, height },
      })],
    }));
  }

  headerChildren.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({
      text: company, bold: true, size: 26, color: NAVY_HEX, font: "Times New Roman",
    })],
  }));

  headerChildren.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({
      text: "Liste des Informations Documentées",
      bold: true, size: 32, color: NAVY_HEX, font: "Times New Roman",
    })],
  }));

  headerChildren.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({
      text: "Code : ENG-LID-01   |   Version : 00   |   Date : " + new Date().toLocaleDateString("fr-FR"),
      size: 18, color: "555555", font: "Times New Roman",
    })],
  }));

  headerChildren.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder,
      insideHorizontal: tableBorder, insideVertical: tableBorder,
    },
    rows,
  }));

  const doc = new Document({
    sections: [{
      properties: { page: { size: { orientation: "landscape" } } },
      children: headerChildren,
    }],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, "liste-informations-documentees.docx");
}

/* =========================================================================
   ÉCRAN DE CONNEXION / CRÉATION DE COMPTE
   ========================================================================= */
const authStyles = {
  wrap: {
    fontFamily: "'Syne', sans-serif", minHeight: "100vh",
    display: "flex", alignItems: "center", justifyContent: "center",
    background: "#EDEBE4", padding: "24px", color: "#2C2C2A",
  },
  brand: {
    background: "linear-gradient(155deg, #0F6E56 0%, #0A4A3A 100%)",
    color: "#FFFFFF", padding: "52px 44px",
    display: "flex", flexDirection: "column", justifyContent: "space-between",
    position: "relative", overflow: "hidden",
  },
  brandLogo: {
    width: "52px", height: "52px", background: "rgba(255,255,255,0.15)",
    border: "1px solid rgba(255,255,255,0.25)",
    borderRadius: "12px", display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: "26px", fontWeight: "800", marginBottom: "26px",
  },
  brandTitle: { fontSize: "30px", fontWeight: "800", lineHeight: 1.15, letterSpacing: "-0.5px" },
  brandText: { fontSize: "13px", opacity: 0.82, marginTop: "16px", lineHeight: 1.6 },
  brandFoot: { fontSize: "11px", fontFamily: "'IBM Plex Mono', monospace", opacity: 0.7, letterSpacing: "1px" },
  blob1: { position: "absolute", width: "260px", height: "260px", borderRadius: "50%",
    background: "rgba(255,255,255,0.06)", bottom: "-80px", right: "-70px" },
  blob2: { position: "absolute", width: "160px", height: "160px", borderRadius: "50%",
    background: "rgba(255,255,255,0.05)", top: "-50px", left: "-40px" },
  formSide: { padding: "48px 44px", display: "flex", flexDirection: "column", justifyContent: "center" },
  tabs: { display: "flex", gap: "6px", background: "#F1EFE8", padding: "5px", borderRadius: "10px", marginBottom: "28px" },
  tab: {
    flex: 1, textAlign: "center", padding: "10px", borderRadius: "7px",
    fontSize: "13px", fontWeight: "700", cursor: "pointer", border: "none",
    background: "transparent", color: "#888780", fontFamily: "'Syne', sans-serif",
    transition: "all 0.2s",
  },
  tabActive: { background: "#FFFFFF", color: "#0F6E56", boxShadow: "0 1px 4px rgba(0,0,0,0.08)" },
  formTitle: { fontSize: "22px", fontWeight: "800", letterSpacing: "-0.5px", marginBottom: "4px" },
  formSub: { fontSize: "12px", color: "#888780", marginBottom: "26px" },
  fGroup: { marginBottom: "16px" },
  fLabel: {
    display: "block", fontSize: "10px", fontFamily: "'IBM Plex Mono', monospace",
    color: "#5F5E5A", marginBottom: "6px", letterSpacing: "1px", textTransform: "uppercase",
  },
  fInput: {
    width: "100%", background: "#F8F8F6", border: "1px solid #D3D1C7",
    borderRadius: "9px", color: "#2C2C2A", fontFamily: "'Syne', sans-serif",
    fontSize: "15px", padding: "13px 15px", outline: "none",
    boxSizing: "border-box", transition: "border-color 0.2s",
  },
  error: {
    background: "#FBEAEA", border: "1px solid #E8B4B4", color: "#A53030",
    borderRadius: "8px", padding: "10px 14px", fontSize: "12px", marginBottom: "16px",
  },
  submit: {
    width: "100%", padding: "14px", borderRadius: "9px", border: "none",
    background: "#0F6E56", color: "#FFFFFF", fontFamily: "'Syne', sans-serif",
    fontSize: "15px", fontWeight: "800", cursor: "pointer", marginTop: "6px",
    letterSpacing: "0.3px",
  },
  switchHint: { fontSize: "12px", color: "#888780", textAlign: "center", marginTop: "18px" },
  switchLink: { color: "#0F6E56", fontWeight: "700", cursor: "pointer" },
  note: { fontSize: "10px", color: "#B4B2A9", textAlign: "center", marginTop: "20px", lineHeight: 1.5 },
  logoPick: {
    width: "54px", height: "54px", borderRadius: "10px", border: "1px solid #D3D1C7",
    background: "#F8F8F6", display: "flex", alignItems: "center", justifyContent: "center",
    overflow: "hidden", flexShrink: 0,
  },
  logoPickBtn: {
    flex: 1, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
    background: "#EAF3DE", border: "1px solid #A8D08D", borderRadius: "9px",
    color: "#0F6E56", fontFamily: "'Syne', sans-serif", fontSize: "13px", fontWeight: "700",
    padding: "13px 12px", height: "54px", boxSizing: "border-box",
  },
};

/* --- Petites icônes (SVG inline, stroke = currentColor) --- */
const IcUser = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
    strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" /><circle cx="12" cy="7" r="4" />
  </svg>
);
const IcMail = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
    strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="2" y="4" width="20" height="16" rx="2" /><path d="m22 6-10 7L2 6" />
  </svg>
);
const IcLock = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
    strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="11" width="18" height="11" rx="2" /><path d="M7 11V7a5 5 0 0 1 10 0v4" />
  </svg>
);
const IcBuild = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
    strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 21h18" /><path d="M5 21V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v16" />
    <path d="M9 9h.01M9 13h.01M9 17h.01M15 9h.01M15 13h.01M15 17h.01" />
  </svg>
);

function AuthScreen({ onAuthenticated }) {
  const [mode, setMode]       = useState("login"); // "login" | "signup"
  const [name, setName]       = useState("");
  const [company, setCompany] = useState("");
  const [email, setEmail]     = useState("");
  const [pw, setPw]           = useState("");
  const [pw2, setPw2]         = useState("");
  const [error, setError]     = useState("");
  const [logo, setLogo]       = useState(null);
  const [logoW, setLogoW]     = useState(null);
  const [logoH, setLogoH]     = useState(null);

  const reset = () => { setError(""); setPw(""); setPw2(""); };
  const goLogin  = () => { setMode("login"); reset(); };
  const goSignup = () => { setMode("signup"); reset(); };

  const handleLogoPick = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      const { dataUrl, width, height } = await readImageFile(file, 256);
      setLogo(dataUrl); setLogoW(width); setLogoH(height); setError("");
    } catch (err) { setError("Image illisible. Choisissez un PNG ou un JPG."); }
    e.target.value = "";
  };
  const clearLogo = () => { setLogo(null); setLogoW(null); setLogoH(null); };

  const doLogin = () => {
    setError("");
    const key = email.trim().toLowerCase();
    if (!key || !pw) { setError("Veuillez saisir votre email et votre mot de passe."); return; }
    const accounts = loadAccounts();
    const acc = accounts[key];
    if (!acc) { setError("Aucun compte n'existe avec cet email sur ce navigateur."); return; }
    if (acc.passHash !== hashPassword(pw)) { setError("Mot de passe incorrect."); return; }
    onAuthenticated(acc.email);
  };

  const doSignup = () => {
    setError("");
    if (!name.trim() || !company.trim() || !email.trim() || !pw) {
      setError("Le nom, l'entreprise, l'email et le mot de passe sont obligatoires."); return;
    }
    if (!/^\S+@\S+\.\S+$/.test(email.trim())) { setError("Adresse email invalide."); return; }
    if (pw.length < 4) { setError("Le mot de passe doit faire au moins 4 caractères."); return; }
    if (pw !== pw2) { setError("Les mots de passe ne correspondent pas."); return; }
    const key = email.trim().toLowerCase();
    const accounts = loadAccounts();
    if (accounts[key]) { setError("Un compte existe déjà avec cet email."); return; }
    accounts[key] = {
      email: key, name: name.trim(), company: company.trim(),
      passHash: hashPassword(pw), createdAt: new Date().toISOString(),
      logo: logo || null, logoW: logoW || null, logoH: logoH || null,
    };
    saveAccounts(accounts);
    onAuthenticated(key);
  };

  return (
    <div className="auth-wrap">
      <style>{`
        .auth-wrap * { box-sizing: border-box; }
        .auth-wrap {
          position: relative; min-height: 100vh; width: 100%;
          display: flex; align-items: center; justify-content: center;
          background: #EDEFF0; font-family: 'Syne', sans-serif;
          padding: 24px; overflow: hidden;
        }
        .bg-shape { position: absolute; z-index: 0; }
        .bg-red {
          top: -70px; right: -70px; width: 280px; height: 280px;
          background: linear-gradient(135deg,#F26D7D,#E8576B);
          border-radius: 44px; transform: rotate(45deg);
          box-shadow: 0 20px 50px rgba(232,87,107,.25);
        }
        .bg-yellow {
          bottom: -120px; left: -120px; width: 330px; height: 330px;
          background: linear-gradient(135deg,#FFD45E,#F6C544); border-radius: 50%;
        }

        .auth-container {
          position: relative; z-index: 1;
          width: 880px; max-width: 100%; min-height: 600px;
          background: #FFFFFF; border-radius: 20px; overflow: hidden;
          box-shadow: 0 30px 80px rgba(15,90,70,.20);
        }

        .form-container {
          position: absolute; top: 0; height: 100%; width: 50%;
          transition: all .6s ease-in-out; overflow-y: auto;
        }
        .sign-in-container { left: 0; z-index: 2; }
        .sign-up-container { left: 0; opacity: 0; z-index: 1; }
        .auth-container.right-panel-active .sign-in-container { transform: translateX(100%); opacity: 0; }
        .auth-container.right-panel-active .sign-up-container { transform: translateX(100%); opacity: 1; z-index: 5; animation: ac-show .6s; }
        @keyframes ac-show { 0%,49.99% { opacity:0; z-index:1; } 50%,100% { opacity:1; z-index:5; } }

        .auth-form {
          display: flex; flex-direction: column; align-items: center; justify-content: center;
          height: 100%; padding: 34px 48px; text-align: center;
        }
        .brand-mini {
          font-size: 11px; letter-spacing: 2px; text-transform: uppercase;
          color: #0F6E56; font-weight: 800; margin-bottom: 12px;
          display: flex; align-items: center; gap: 6px;
        }
        .brand-glyph { color: #1AAE86; }
        .auth-title { font-size: 27px; font-weight: 800; color: #0F6E56; margin: 0; letter-spacing: -.6px; }
        .auth-sub { font-size: 12px; color: #9AA0A6; margin: 6px 0 16px; }

        .logo-drop {
          width: 92px; height: 92px; border-radius: 50%;
          border: 2px dashed #CBD0CC; background: #F5F6F5;
          display: flex; align-items: center; justify-content: center;
          cursor: pointer; overflow: hidden; flex-shrink: 0;
          transition: border-color .2s, transform .15s;
        }
        .logo-drop:hover { transform: scale(1.04); border-color: #1AAE86; }
        .logo-drop.has-logo { border-style: solid; border-color: #1AAE86; background: #FFFFFF; padding: 6px; }
        .logo-drop img { width: 100%; height: 100%; object-fit: contain; }
        .logo-drop-empty { font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #B4BAB5; line-height: 1.3; }
        .logo-cap { font-size: 11px; color: #9AA0A6; margin: 8px 0 6px; }
        .logo-link { background: none; border: none; color: #0F6E56; font-weight: 700; font-size: 12px;
          cursor: pointer; margin: 8px 0 4px; font-family: 'Syne', sans-serif; }

        .in-row {
          display: flex; align-items: center; gap: 10px; width: 100%;
          background: #EFF0F0; border: 1px solid #EFF0F0; border-radius: 10px;
          padding: 0 14px; margin: 6px 0; transition: border-color .2s, background .2s;
        }
        .in-row:focus-within { border-color: #1AAE86; background: #FFFFFF; }
        .in-row svg { color: #A2A8A4; flex-shrink: 0; }
        .in-row input {
          flex: 1; border: none; background: transparent; outline: none;
          padding: 13px 0; font-size: 14px; color: #2C2C2A; font-family: 'Syne', sans-serif;
        }
        .auth-error {
          width: 100%; background: #FCEBEB; border: 1px solid #F1C7C7; color: #C0392B;
          border-radius: 9px; padding: 9px 12px; font-size: 12px; margin: 6px 0 4px;
        }

        .btn-solid {
          margin-top: 18px; border: 1px solid #0F6E56;
          background: linear-gradient(135deg,#1AAE86,#0F6E56); color: #fff;
          border-radius: 24px; padding: 13px 52px; cursor: pointer;
          font-size: 12px; font-weight: 800; letter-spacing: 1.5px; text-transform: uppercase;
          font-family: 'Syne', sans-serif; transition: transform .1s, box-shadow .2s;
          box-shadow: 0 8px 20px rgba(15,110,86,.25);
        }
        .btn-solid:hover { box-shadow: 0 10px 26px rgba(15,110,86,.32); }
        .btn-solid:active { transform: scale(.96); }

        .btn-ghost {
          margin-top: 22px; border: 1.5px solid rgba(255,255,255,.9); background: transparent; color: #fff;
          border-radius: 24px; padding: 12px 46px; cursor: pointer;
          font-size: 12px; font-weight: 800; letter-spacing: 1.5px; text-transform: uppercase;
          font-family: 'Syne', sans-serif; transition: background .2s, transform .1s;
        }
        .btn-ghost:hover { background: rgba(255,255,255,.14); }
        .btn-ghost:active { transform: scale(.96); }

        .overlay-container {
          position: absolute; top: 0; left: 50%; width: 50%; height: 100%;
          overflow: hidden; transition: transform .6s ease-in-out; z-index: 100;
        }
        .auth-container.right-panel-active .overlay-container { transform: translateX(-100%); }
        .overlay {
          position: relative; left: -100%; height: 100%; width: 200%;
          background: linear-gradient(135deg,#1AAE86 0%,#0E6A52 100%);
          color: #fff; transform: translateX(0); transition: transform .6s ease-in-out;
        }
        .auth-container.right-panel-active .overlay { transform: translateX(50%); }
        .overlay-panel {
          position: absolute; top: 0; height: 100%; width: 50%;
          display: flex; flex-direction: column; align-items: center; justify-content: center;
          text-align: center; padding: 0 44px;
        }
        .overlay-left { left: 0; transform: translateX(-18%); }
        .auth-container.right-panel-active .overlay-left { transform: translateX(0); }
        .overlay-right { right: 0; transform: translateX(0); }
        .auth-container.right-panel-active .overlay-right { transform: translateX(18%); }
        .overlay h1 { font-size: 30px; font-weight: 800; margin: 14px 0 0; letter-spacing: -.5px; }
        .overlay p { font-size: 13px; line-height: 1.6; margin: 16px 0 0; opacity: .92; max-width: 260px; }
        .ov-logo {
          width: 90px; height: 90px; border-radius: 16px;
          background: rgba(255,255,255,.16); border: 1px solid rgba(255,255,255,.3);
          display: flex; align-items: center; justify-content: center;
          font-size: 26px; font-weight: 800;
        }
        .ov-foot {
          position: absolute; bottom: 26px; left: 0; right: 0;
          font-size: 10px; letter-spacing: 1px; opacity: .7;
          font-family: 'IBM Plex Mono', monospace; padding: 0 30px;
        }
        .ov-diamond { position: absolute; background: rgba(255,255,255,.08); transform: rotate(45deg); border-radius: 8px; z-index: 0; }
        .ov-diamond.d1 { width: 66px; height: 66px; top: 14%; left: 30%; }
        .ov-diamond.d2 { width: 32px; height: 32px; bottom: 20%; left: 22%; }
        .ov-diamond.d3 { width: 46px; height: 46px; top: 24%; right: 26%; }

        .mobile-switch { display: none; margin-top: 18px; font-size: 12px; color: #9AA0A6; }
        .mobile-switch span { color: #0F6E56; font-weight: 800; cursor: pointer; }

        @media (max-width: 780px) {
          .auth-container { min-height: auto; }
          .overlay-container { display: none; }
          .form-container { position: relative; width: 100%; height: auto; transform: none !important; }
          .sign-in-container { opacity: 1; }
          .sign-up-container { display: none; }
          .auth-container.right-panel-active .sign-in-container { display: none; }
          .auth-container.right-panel-active .sign-up-container { display: block; opacity: 1; transform: none !important; animation: none; }
          .mobile-switch { display: block; }
          .auth-form { padding: 36px 26px; }
        }
      `}</style>

      <span className="bg-shape bg-red" />
      <span className="bg-shape bg-yellow" />

      <div className={"auth-container" + (mode === "signup" ? " right-panel-active" : "")}>

        {/* ----- Formulaire INSCRIPTION (panneau droit) ----- */}
        <div className="form-container sign-up-container">
          <div className="auth-form">
            <div className="brand-mini"><span className="brand-glyph">◆</span> CoDoc Generator</div>
            <h2 className="auth-title">Créer un compte</h2>
            <p className="auth-sub">Votre espace de codification documentaire</p>

            <label className={"logo-drop" + (logo ? " has-logo" : "")} title="Importer le logo de l'entreprise">
              {logo
                ? <img src={logo} alt="logo" />
                : <span className="logo-drop-empty">＋<br />LOGO</span>}
              <input type="file" accept="image/*" style={{ display: "none" }} onChange={handleLogoPick} />
            </label>
            {logo
              ? <button type="button" className="logo-link" onClick={clearLogo}>Retirer le logo</button>
              : <div className="logo-cap">Logo de l'entreprise (optionnel)</div>}

            {error && mode === "signup" && <div className="auth-error">{error}</div>}

            <div className="in-row"><IcUser /><input placeholder="Nom complet"
              value={name} onChange={e => setName(e.target.value)} /></div>
            <div className="in-row"><IcBuild /><input placeholder="Entreprise"
              value={company} onChange={e => setCompany(e.target.value)} /></div>
            <div className="in-row"><IcMail /><input type="email" placeholder="Email"
              value={email} onChange={e => setEmail(e.target.value)} /></div>
            <div className="in-row"><IcLock /><input type="password" placeholder="Mot de passe"
              value={pw} onChange={e => setPw(e.target.value)} /></div>
            <div className="in-row"><IcLock /><input type="password" placeholder="Confirmer le mot de passe"
              value={pw2} onChange={e => setPw2(e.target.value)}
              onKeyDown={e => e.key === "Enter" && doSignup()} /></div>

            <button className="btn-solid" onClick={doSignup}>S'inscrire</button>

            <div className="mobile-switch">
              Déjà inscrit ? <span onClick={goLogin}>Se connecter</span>
            </div>
          </div>
        </div>

        {/* ----- Formulaire CONNEXION (panneau gauche par défaut) ----- */}
        <div className="form-container sign-in-container">
          <div className="auth-form">
            <div className="brand-mini"><span className="brand-glyph">◆</span> CoDoc Generator</div>
            <h2 className="auth-title">Connexion</h2>
            <p className="auth-sub">Heureux de vous revoir</p>

            {error && mode === "login" && <div className="auth-error">{error}</div>}

            <div className="in-row"><IcMail /><input type="email" placeholder="Email"
              value={email} onChange={e => setEmail(e.target.value)}
              onKeyDown={e => e.key === "Enter" && doLogin()} /></div>
            <div className="in-row"><IcLock /><input type="password" placeholder="Mot de passe"
              value={pw} onChange={e => setPw(e.target.value)}
              onKeyDown={e => e.key === "Enter" && doLogin()} /></div>

            <button className="btn-solid" onClick={doLogin}>Se connecter</button>

            <div className="mobile-switch">
              Pas de compte ? <span onClick={goSignup}>Créer un compte</span>
            </div>
          </div>
        </div>

        {/* ----- Overlay vert coulissant ----- */}
        <div className="overlay-container">
          <div className="overlay">
            <span className="ov-diamond d1" />
            <span className="ov-diamond d2" />
            <span className="ov-diamond d3" />

            {/* Affiché en mode INSCRIPTION → invite à se connecter */}
            <div className="overlay-panel overlay-left">
              <div className="ov-logo" style={{ background: "white", padding: "6px" }}>
              <img src={logoOcg} alt="logo" style={{ width: "100%", height: "100%", objectFit: "contain" }} />
            </div>
              <h1>Déjà membre ?</h1>
              <p>Connectez-vous pour retrouver votre tableau de codes et votre historique.</p>
              <button className="btn-ghost" onClick={goLogin}>Se connecter</button>
              <div className="ov-foot">OCG CONSULTING — ISO 9001 · 14001 · 45001 · 27001</div>
            </div>

            {/* Affiché en mode CONNEXION → invite à s'inscrire */}
            <div className="overlay-panel overlay-right">
              <div className="ov-logo" style={{ background: "white", padding: "6px" }}>
              <img src={logoOcg} alt="logo" style={{ width: "100%", height: "100%", objectFit: "contain" }} />
            </div>
              <h1>Bonjour 👋</h1>
              <p>Créez votre espace pour générer et archiver vos codes documentaires QMS / ISO.</p>
              <button className="btn-ghost" onClick={goSignup}>Créer un compte</button>
              <div className="ov-foot">OCG CONSULTING — ISO 9001 · 14001 · 45001 · 27001</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   APPLICATION PRINCIPALE (par utilisateur connecté)
   ========================================================================= */
const EMPTY_COUNTERS = { PS:0, PR:0, ENG:0, F:0, L:0, INS:0, MO:0 };

/* --- Jeu d'icônes (SVG inline, stroke = currentColor) --- */
const mkIcon = (children) => ({ size = 18, ...rest }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
    stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...rest}>
    {children}
  </svg>
);
const IcSparkles = mkIcon(<>
  <path d="M12 3l1.8 4.7L18.5 9.5 13.8 11.3 12 16l-1.8-4.7L5.5 9.5l4.7-1.8z" />
  <path d="M19 14l.6 1.6 1.6.6-1.6.6-.6 1.6-.6-1.6-1.6-.6 1.6-.6z" />
</>);
const IcFile = mkIcon(<><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><path d="M14 2v6h6" /></>);
const IcBookOpen = mkIcon(<>
  <path d="M2 4h7a3 3 0 0 1 3 3v13a2.5 2.5 0 0 0-2.5-2.5H2z" />
  <path d="M22 4h-7a3 3 0 0 0-3 3v13a2.5 2.5 0 0 1 2.5-2.5H22z" />
</>);
const IcInfo = mkIcon(<><circle cx="12" cy="12" r="10" /><path d="M12 16v-4M12 8h.01" /></>);
const IcTrendUp = mkIcon(<><path d="M22 7l-8.5 8.5-5-5L2 17" /><path d="M16 7h6v6" /></>);
const IcCalendar = mkIcon(<><rect x="3" y="4" width="18" height="18" rx="2" /><path d="M16 2v4M8 2v4M3 10h18" /></>);
const IcSearch = mkIcon(<><circle cx="11" cy="11" r="7" /><path d="m21 21-4.3-4.3" /></>);
const IcHistory = mkIcon(<><path d="M3 12a9 9 0 1 0 3-6.7L3 8" /><path d="M3 3v5h5" /><path d="M12 7v5l3 2" /></>);
const IcFileText = mkIcon(<>
  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><path d="M14 2v6h6" />
  <path d="M9 13h6M9 17h4" />
</>);
const IcTrash = mkIcon(<><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6" /></>);
const IcClipboardCheck = mkIcon(<>
  <rect x="8" y="2" width="8" height="4" rx="1" />
  <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
  <path d="m9 14 2 2 4-4" />
</>);
const IcGitBranch = mkIcon(<>
  <rect x="3" y="3" width="6" height="6" rx="1" /><rect x="15" y="15" width="6" height="6" rx="1" />
  <path d="M6 9v6a3 3 0 0 0 3 3h6" />
</>);
const IcDatabase = mkIcon(<>
  <ellipse cx="12" cy="5" rx="8" ry="3" />
  <path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5" /><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6" />
</>);
const IcListChecks = mkIcon(<><path d="m3 6 2 2 3-3" /><path d="m3 14 2 2 3-3" /><path d="M11 6h10M11 15h10" /></>);
const IcSitemap = mkIcon(<>
  <rect x="9" y="3" width="6" height="5" rx="1" /><rect x="3" y="16" width="6" height="5" rx="1" />
  <rect x="15" y="16" width="6" height="5" rx="1" /><path d="M12 8v4M6 16v-2h12v2" />
</>);
const IcGear = mkIcon(<>
  <circle cx="12" cy="12" r="3.2" />
  <path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M19.1 4.9 17 7M7 17l-2.1 2.1" />
</>);
const IcChevronRight = mkIcon(<path d="m9 18 6-6-6-6" />);
const IcFileSearch = mkIcon(<>
  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h5" /><path d="M14 2v6h6" />
  <circle cx="16.5" cy="16.5" r="2.5" /><path d="m21 21-1.6-1.6" />
</>);
const IcLogout = mkIcon(<><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /><path d="m16 17 5-5-5-5M21 12H9" /></>);
const IcLayers = mkIcon(<><path d="m12 2 9 5-9 5-9-5 9-5z" /><path d="m3 12 9 5 9-5M3 17l9 5 9-5" /></>);
const IcArchive = mkIcon(<><rect x="3" y="4" width="18" height="4" rx="1" /><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8" /><path d="M10 12h4" /></>);
const IcRestore = mkIcon(<><path d="M3 12a9 9 0 1 0 3-6.7L3 8" /><path d="M3 3v5h5" /></>);

const TYPE_ICONS = {
  processus: IcClipboardCheck,
  procedure: IcGitBranch,
  enregistrement: IcDatabase,
  formulaire: IcListChecks,
  logigramme: IcSitemap,
  instruction: IcBookOpen,
  modeoperatoire: IcGear,
};

const TYPE_DESCRIPTIONS = {
  processus: "Décrit une activité maîtrisée transformant des entrées en sorties (approche processus ISO).",
  procedure: "Définit la manière de réaliser une activité : qui fait quoi, quand et comment.",
  enregistrement: "Preuve documentée des résultats obtenus ou des activités réalisées (registres, listes).",
  formulaire: "Support vierge servant à collecter et tracer des informations (fiches à remplir).",
  logigramme: "Représentation graphique séquentielle d'un processus ou d'une procédure.",
  instruction: "Consigne détaillée pour exécuter une tâche précise à un poste donné.",
  modeoperatoire: "Description pas-à-pas d'une opération technique spécifique.",
};

/* --- Liste déroulante AVEC RECHERCHE (combobox) --- */
const IcChevronDown = mkIcon(<path d="m6 9 6 6 6-6" />);

function Combobox({ value, onChange, placeholder, groups, searchPlaceholder = "Rechercher…" }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const ref = useRef(null);

  useEffect(() => {
    const onDoc = (e) => { if (ref.current && !ref.current.contains(e.target)) { setOpen(false); setQ(""); } };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  const norm = (s) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  const nq = norm(q.trim());
  const fgroups = groups
    .map((g) => ({ ...g, items: g.items.filter((it) => !nq || norm(it.label).includes(nq)) }))
    .filter((g) => g.items.length > 0);

  let current = placeholder;
  for (const g of groups) for (const it of g.items) if (it.value === value) current = it.label;

  const pick = (v) => { onChange(v); setOpen(false); setQ(""); };

  return (
    <div className="cbx" ref={ref}>
      <button type="button" className={"cbx-btn" + (value ? " has-val" : "")} onClick={() => setOpen((o) => !o)}>
        <span className="cbx-val">{current}</span>
        <span className="cbx-chev"><IcChevronDown size={16} /></span>
      </button>
      {open && (
        <div className="cbx-pop">
          <div className="cbx-search">
            <IcSearch size={14} />
            <input autoFocus placeholder={searchPlaceholder} value={q} onChange={(e) => setQ(e.target.value)} />
          </div>
          <div className="cbx-list">
            {fgroups.length === 0 ? (
              <div className="cbx-empty">Aucun résultat</div>
            ) : fgroups.map((g, gi) => (
              <div key={gi}>
                {g.label && <div className="cbx-group">{g.label}</div>}
                {g.items.map((it, ii) => (
                  <div key={ii} className={"cbx-opt" + (it.value === value ? " sel" : "")} onClick={() => pick(it.value)}>
                    {it.label}
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   APPLICATION PRINCIPALE — tableau de bord
   ========================================================================= */
function DocCodeApp({ user, onLogout, onUpdateUser, onDeleteAccount }) {
  const [type, setType]             = useState("");
  const [customType, setCustomType] = useState("");
  const [customCode, setCustomCode] = useState("");
  const [docModel, setDocModel]     = useState("");
  const [intitule, setIntitule]     = useState("");
  const [query, setQuery]           = useState("");

  const initial = loadUserState(user.email);
  const [codes, setCodes]       = useState(initial?.codes || []);
  const [counters, setCounters] = useState(initial?.counters || { ...EMPTY_COUNTERS });
  const [history, setHistory]   = useState(() => loadUserHistory(user.email));
  const [snapshots, setSnapshots] = useState(() => loadUserSnapshots(user.email));

  const [toast, setToast]               = useState(null);
  const [showHistory, setShowHistory]   = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showTypes, setShowTypes]       = useState(false);
  const [showExports, setShowExports]   = useState(false);

  useEffect(() => {
    try { localStorage.setItem(stateKey(user.email), JSON.stringify({ codes, counters })); }
    catch (e) {}
  }, [codes, counters, user.email]);

  useEffect(() => {
    try { localStorage.setItem(histKey(user.email), JSON.stringify(history)); }
    catch (e) {}
  }, [history, user.email]);

  useEffect(() => {
    try { localStorage.setItem(snapKey(user.email), JSON.stringify(snapshots)); }
    catch (e) {}
  }, [snapshots, user.email]);

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 3000); };

  const handleLogoUpload = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      const { dataUrl, width, height } = await readImageFile(file, 256);
      onUpdateUser({ logo: dataUrl, logoW: width, logoH: height });
      showToast("✅ Logo mis à jour");
    } catch (err) { showToast("❌ Image illisible"); }
    e.target.value = "";
  };
  const handleLogoRemove = () => {
    onUpdateUser({ logo: null, logoW: null, logoH: null });
    showToast("🗑️ Logo supprimé");
  };

  const handleDeleteAccount = () => {
    const ok = window.confirm(
      `Supprimer définitivement le compte « ${user.company} » (${user.email}) ?\n\n` +
      `Tous ses codes et son historique seront effacés. Cette action est IRRÉVERSIBLE.`
    );
    if (!ok) return;
    onDeleteAccount();
  };

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
      code, intitule: intitule.trim(), type: typeCode, typeLabel, seq,
      version: 0, motif: "Création", createdAt: new Date().toISOString(),
    };

    setCodes((prev) => [entry, ...prev]);
    setCounters((prev) => ({ ...prev, [typeCode]: seq }));
    setHistory((prev) => [entry, ...prev]);
    showToast(`✅ Code généré : ${code}`);

    setIntitule(""); setType(""); setCustomType(""); setCustomCode(""); setDocModel("");
  };

  const handleDelete = (id) => {
    setCodes((prev) => prev.filter((c) => c.id !== id));
    showToast("🗑️ Code supprimé du tableau (conservé dans l'historique)");
  };

  const handleClearAll = () => {
    if (codes.length === 0) { showToast("⚠️ Le tableau est déjà vide"); return; }
    if (!window.confirm(`Vider le tableau (${codes.length} codes) et réinitialiser les compteurs ?\n\nL'historique sera conservé.`)) return;
    setCodes([]);
    setCounters({ ...EMPTY_COUNTERS });
    showToast("🗑️ Tableau vidé — compteurs réinitialisés");
  };

  const handleClearHistory = () => {
    if (history.length === 0) { showToast("⚠️ L'historique est déjà vide"); return; }
    if (!window.confirm(`Effacer DÉFINITIVEMENT les ${history.length} entrées de l'historique ?`)) return;
    setHistory([]);
    showToast("🗑️ Historique effacé");
  };

  const branding = { company: user.company, logo: user.logo || null, logoW: user.logoW, logoH: user.logoH };

  // Sauvegarde une copie complète du tableau courant (instantané d'export)
  const saveSnapshot = (format) => {
    const snap = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      createdAt: new Date().toISOString(),
      format,                                   // "word" | "excel"
      count: codes.length,
      codes: codes.map((c) => ({ ...c })),      // copie figée
      counters: { ...counters },
    };
    setSnapshots((prev) => [snap, ...prev]);
  };

  // Remet un instantané dans le tableau de travail (pour le modifier puis réexporter)
  const handleRestoreSnapshot = (snap) => {
    if (codes.length > 0 &&
      !window.confirm("Remplacer le tableau actuel par ce tableau exporté ?\n\nLe tableau courant sera écrasé. L'historique et les autres tableaux exportés sont conservés.")) return;
    setCodes(snap.codes.map((c) => ({ ...c })));
    setCounters({ ...EMPTY_COUNTERS, ...snap.counters });
    setShowExports(false);
    showToast("✅ Tableau restauré — modifiez-le puis réexportez");
  };

  // Réexporte un instantané tel quel
  const handleReexport = async (snap, fmt) => {
    try {
      if (fmt === "word") await exportWord(snap.codes, branding);
      else await exportExcel(snap.codes, branding);
      showToast(fmt === "word" ? "📄 Word téléchargé" : "📊 Excel téléchargé");
    } catch (e) { console.error(e); showToast("❌ Erreur export"); }
  };

  const handleDeleteSnapshot = (id) => {
    setSnapshots((prev) => prev.filter((s) => s.id !== id));
    showToast("🗑️ Tableau exporté supprimé");
  };

  const handleClearSnapshots = () => {
    if (snapshots.length === 0) { showToast("⚠️ Aucun tableau exporté"); return; }
    if (!window.confirm(`Supprimer les ${snapshots.length} tableaux exportés sauvegardés ?`)) return;
    setSnapshots([]);
    showToast("🗑️ Tableaux exportés supprimés");
  };

  const handleExportWord = async () => {
    if (codes.length === 0) { showToast("⚠️ Aucun code à exporter"); return; }
    try { await exportWord(codes, branding); saveSnapshot("word"); showToast("📄 Word téléchargé et sauvegardé"); }
    catch (e) { console.error(e); showToast("❌ Erreur export Word"); }
  };
  const handleExportExcel = async () => {
    if (codes.length === 0) { showToast("⚠️ Aucun code à exporter"); return; }
    try { await exportExcel(codes, branding); saveSnapshot("excel"); showToast("📊 Excel téléchargé et sauvegardé"); }
    catch (e) { console.error(e); showToast("❌ Erreur export Excel"); }
  };

  const preview = previewCode(type, customCode, intitule, counters);
  const selectedColor = type && type !== "__autre__" ? TYPE_LABELS[type]?.color : "#3B6D11";
  const initials = (user.company || user.name || "?")
    .trim().split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase();
  const isFirstTime = codes.length === 0 && history.length === 0;

  const q = query.trim().toLowerCase();
  const filtered = q
    ? codes.filter((c) => c.code.toLowerCase().includes(q) || c.intitule.toLowerCase().includes(q))
    : codes;

  const ARROW = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%239AA0A6' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E";

  const statCard = (key) => {
    const val = TYPE_LABELS[key];
    const Ic = TYPE_ICONS[key];
    return (
      <div className="stat-card" key={key}
        style={{ background: `${val.color}0D`, borderColor: `${val.color}26` }}>
        <div className="stat-top">
          <div className="stat-ic" style={{ background: `${val.color}1F`, color: val.color }}>
            <Ic size={16} />
          </div>
          <div className="stat-name">{TYPE_LABELS_PLURAL[val.code]}</div>
        </div>
        <div className="stat-val">{String(counters[val.code] || 0).padStart(2, "0")}</div>
        <div className="stat-foot">
          <span className="stat-total">Total</span>
          <span className="stat-trend"><IcTrendUp size={13} /> 0%</span>
        </div>
      </div>
    );
  };

  const badge = (color) => ({
    display: "inline-flex", padding: "4px 10px", borderRadius: "8px",
    fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", fontWeight: 800,
    background: `${color}18`, color,
  });

  return (
    <div className="dcg">
      <style>{`
        .dcg, .dcg * { box-sizing: border-box; }
        .dcg { font-family: 'Syne', sans-serif; color: #1F2421; background: #F4F6F5; min-height: 100vh; }
        .dcg button { cursor: pointer; }
        @keyframes slideUp { from { transform: translateY(16px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        .dcg ::-webkit-scrollbar { width: 6px; height: 6px; }
        .dcg ::-webkit-scrollbar-thumb { background: #D3D7D4; border-radius: 3px; }

        .dcg-shell { display: grid; grid-template-columns: 340px 1fr; min-height: 100vh; }

        /* ---- Sidebar ---- */
        .dcg-side { background: #FFFFFF; border-right: 1px solid #E7E9E7; padding: 26px 26px 40px; overflow-y: auto; }
        .side-logo { display: flex; align-items: center; gap: 14px; margin-bottom: 22px; }
        .side-logo-mark {
          width: 60px; height: 60px; border-radius: 16px; position: relative; flex-shrink: 0;
          background: linear-gradient(150deg,#1AAE86,#0F6E56); color: #fff;
          display: flex; align-items: center; justify-content: center;
          font-size: 30px; font-weight: 800; box-shadow: 0 8px 20px rgba(15,110,86,.28);
        }
        .side-logo-mark::after { content: ""; position: absolute; inset: 9px; border: 2px solid rgba(255,255,255,.5); border-radius: 9px; }
        .side-logo-mark.has-img { width: 96px; height: 96px; background: #fff; border: 1px solid #E2E5E2; box-shadow: 0 6px 16px rgba(0,0,0,.10); padding: 6px; }
        .side-logo-mark.has-img::after { display: none; }
        .side-logo-mark img { width: 100%; height: 100%; object-fit: contain; }
        .side-logo-txt b { display: block; font-size: 21px; font-weight: 800; line-height: 1.05; letter-spacing: -.5px; color: #15201B; }
        .side-logo-txt b span { color: #0F6E56; }
        .side-logo-txt small { font-size: 11px; letter-spacing: 3px; color: #9AA0A6; font-weight: 700; }

        .side-head {
          display: flex; align-items: center; gap: 8px; margin: 24px 0 14px;
          font-size: 12px; font-weight: 800; letter-spacing: 1px; text-transform: uppercase; color: #0F6E56;
        }
        .fld { margin-bottom: 15px; }
        .fld-label { display: block; font-size: 12.5px; font-weight: 700; color: #5C635E; margin-bottom: 7px; }
        .fld-input, .fld-select {
          width: 100%; background: #F6F7F6; border: 1px solid #E2E5E2; border-radius: 11px;
          padding: 13px 14px; font-size: 14px; color: #1F2421; font-family: 'Syne', sans-serif;
          outline: none; transition: border-color .2s, background .2s;
        }
        .fld-input:focus, .fld-select:focus { border-color: #1AAE86; background: #fff; }
        .fld-select { appearance: none; cursor: pointer;
          background-image: url("${ARROW}"); background-repeat: no-repeat; background-position: right 14px center; }

        /* Combobox (liste cherchable) */
        .cbx { position: relative; }
        .cbx-btn {
          width: 100%; display: flex; align-items: center; justify-content: space-between; gap: 10px;
          background: #F6F7F6; border: 1px solid #E2E5E2; border-radius: 11px;
          padding: 13px 14px; font-size: 14px; color: #1F2421; font-family: 'Syne', sans-serif; text-align: left;
          transition: border-color .2s, background .2s;
        }
        .cbx-btn:hover { border-color: #CBD3CD; }
        .cbx-val { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .cbx-btn:not(.has-val) .cbx-val { color: #9AA0A6; }
        .cbx-chev { color: #9AA0A6; flex-shrink: 0; display: flex; }
        .cbx-pop {
          position: absolute; z-index: 50; top: calc(100% + 6px); left: 0; right: 0;
          background: #fff; border: 1px solid #E2E5E2; border-radius: 12px; overflow: hidden;
          box-shadow: 0 14px 34px rgba(0,0,0,.14); animation: fadeIn .12s ease;
        }
        .cbx-search { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-bottom: 1px solid #EEF0EE; }
        .cbx-search svg { color: #A2A8A4; flex-shrink: 0; }
        .cbx-search input { flex: 1; border: none; outline: none; background: transparent; font-size: 13.5px; font-family: 'Syne', sans-serif; color: #1F2421; }
        .cbx-list { max-height: 260px; overflow-y: auto; padding: 6px; }
        .cbx-group { font-size: 10.5px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: #9AA0A6; padding: 10px 10px 6px; }
        .cbx-opt { padding: 9px 10px; border-radius: 8px; font-size: 13.5px; color: #39403B; cursor: pointer; }
        .cbx-opt:hover { background: #F1F6F3; }
        .cbx-opt.sel { background: #E7F4ED; color: #0F6E56; font-weight: 700; }
        .cbx-empty { padding: 18px; text-align: center; font-size: 12.5px; color: #A2A8A4; }

        .autre-box { background: #F3F8F5; border: 1px dashed #B9D9CC; border-radius: 11px; padding: 13px; margin-top: 10px; }
        .autre-cap { font-size: 11px; font-weight: 700; color: #6FA593; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; }
        .autre-code-input {
          width: 100%; text-align: center; text-transform: uppercase; letter-spacing: 4px;
          font-family: 'IBM Plex Mono', monospace; font-size: 20px; font-weight: 700; color: #0F6E56;
          background: #EAF6F1; border: 2px solid #1AAE86; border-radius: 9px; padding: 10px; outline: none;
        }
        .autre-tag { font-size: 10px; font-family: 'IBM Plex Mono', monospace; color: #3B6D11; text-align: center; margin-top: 6px; }

        .preview-box { background: #EAF6F1; border: 1px dashed #9DD3C2; border-radius: 12px; padding: 14px 16px; margin: 4px 0 16px; }
        .preview-cap { font-size: 11px; font-weight: 700; color: #6FA593; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px; }
        .preview-code { font-family: 'IBM Plex Mono', monospace; font-size: 24px; font-weight: 700; letter-spacing: 3px; }

        .btn-generate {
          width: 100%; border: none; border-radius: 12px;
          background: linear-gradient(135deg,#16996F,#0E5E48); color: #fff;
          padding: 14px; font-size: 15px; font-weight: 800; font-family: 'Syne', sans-serif;
          display: flex; align-items: center; justify-content: center; gap: 9px;
          box-shadow: 0 10px 24px rgba(15,110,86,.28); transition: transform .1s, box-shadow .2s;
        }
        .btn-generate:hover { box-shadow: 0 12px 30px rgba(15,110,86,.36); }
        .btn-generate:active { transform: scale(.98); }
        .gen-hint { display: flex; align-items: center; gap: 7px; font-size: 12px; color: #9AA0A6; margin: 12px 2px 0; }

        .ref-list { margin-top: 6px; display: flex; flex-direction: column; gap: 8px; }
        .ref-row {
          display: flex; align-items: center; gap: 12px; padding: 9px 12px;
          border: 1px solid #EEF0EE; border-radius: 11px; background: #FCFDFC; transition: background .15s, border-color .15s;
        }
        .ref-row:hover { background: #F4F8F6; border-color: #DDE6E1; }
        .ref-ic { width: 30px; height: 30px; border-radius: 9px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
        .ref-name { flex: 1; font-size: 13.5px; font-weight: 600; color: #39403B; }
        .ref-badge { font-family: 'IBM Plex Mono', monospace; font-size: 11px; font-weight: 800; padding: 3px 9px; border-radius: 7px; letter-spacing: .5px; }

        .btn-types {
          width: 100%; margin-top: 18px; border: none; border-radius: 11px;
          background: #1C2B25; color: #EAF1EE; padding: 13px 16px; font-size: 13px; font-weight: 700;
          font-family: 'Syne', sans-serif; display: flex; align-items: center; justify-content: space-between; transition: background .2s;
        }
        .btn-types:hover { background: #24382F; }

        /* ---- Main ---- */
        .dcg-main { padding: 22px 30px 44px; overflow: auto; }
        .topbar { display: flex; align-items: center; justify-content: space-between; gap: 16px; margin-bottom: 22px; }
        .topbar-title { font-size: 21px; font-weight: 800; color: #15201B; letter-spacing: -.4px; }
        .topbar-sub { font-size: 12px; color: #9AA0A6; font-family: 'IBM Plex Mono', monospace; margin-top: 3px; }
        .user-wrap { display: flex; align-items: center; gap: 14px; }
        .user-meta { text-align: right; }
        .user-name { font-size: 13px; font-weight: 700; color: #15201B; }
        .user-email { font-size: 11px; color: #9AA0A6; font-family: 'IBM Plex Mono', monospace; }
        .user-avatar {
          width: 46px; height: 46px; border-radius: 50%; overflow: hidden; flex-shrink: 0;
          display: flex; align-items: center; justify-content: center;
          border: 1px solid #E2E5E2; color: #fff; font-weight: 800;
        }
        .btn-logout {
          display: flex; align-items: center; gap: 7px; padding: 11px 16px; border-radius: 10px;
          border: 1px solid #F1C7C7; background: #FCEEEE; color: #C0392B; font-size: 13px; font-weight: 700;
          font-family: 'Syne', sans-serif; white-space: nowrap;
        }
        .btn-logout:hover { background: #FAE1E1; }

        .welcome {
          background: linear-gradient(120deg,#11815F 0%,#0B5240 100%); border-radius: 18px;
          padding: 22px 26px; margin-bottom: 26px; color: #fff;
          display: flex; align-items: center; justify-content: space-between; gap: 18px; flex-wrap: wrap;
          box-shadow: 0 16px 36px rgba(11,82,64,.22);
        }
        .welcome-left { display: flex; align-items: center; gap: 18px; }
        .welcome-logo {
          width: 96px; height: 96px; border-radius: 18px; background: #fff; flex-shrink: 0; padding: 10px;
          display: flex; align-items: center; justify-content: center; box-shadow: 0 6px 18px rgba(0,0,0,.16);
        }
        .welcome-logo img { max-width: 100%; max-height: 100%; object-fit: contain; }
        .welcome-h { font-size: 25px; font-weight: 800; letter-spacing: -.5px; }
        .welcome-p { font-size: 13px; opacity: .9; margin-top: 5px; }
        .welcome-pill {
          display: flex; align-items: center; gap: 8px; padding: 10px 16px; border-radius: 12px;
          background: rgba(255,255,255,.15); border: 1px solid rgba(255,255,255,.25); font-size: 13px; font-weight: 700; white-space: nowrap;
        }

        .sec-head-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
        .sec-head { display: flex; align-items: center; gap: 9px; font-size: 16px; font-weight: 800; color: #15201B; }
        .sec-head .sec-ic { color: #0F6E56; display: flex; }
        .year-chip {
          display: flex; align-items: center; gap: 8px; padding: 9px 14px; border-radius: 10px;
          background: #fff; border: 1px solid #E2E5E2; font-size: 13px; font-weight: 600; color: #5C635E;
        }

        .stats-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 10px; margin-bottom: 10px; }
        .stats-grid-3 { display: grid; grid-template-columns: repeat(3,1fr); gap: 10px; margin-bottom: 20px; }
        .stat-card { border-radius: 12px; padding: 9px 11px; border: 1px solid; }
        .stat-top { display: flex; align-items: center; gap: 8px; }
        .stat-ic { width: 28px; height: 28px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
        .stat-name { font-size: 11.5px; font-weight: 700; color: #39403B; }
        .stat-val { font-size: 19px; font-weight: 800; letter-spacing: -.5px; color: #15201B; margin: 6px 0 0; line-height: 1; }
        .stat-foot { display: flex; align-items: center; justify-content: space-between; margin-top: 5px; }
        .stat-total { font-size: 10.5px; color: #9AA0A6; }
        .stat-trend { display: flex; align-items: center; gap: 3px; font-size: 10.5px; font-weight: 700; color: #2F8F46; }

        .codes-head { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
        .codes-title { font-size: 16px; font-weight: 800; color: #15201B; }
        .codes-sub { font-size: 12px; color: #9AA0A6; font-family: 'IBM Plex Mono', monospace; margin-top: 3px; }
        .toolbar { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
        .search-box { display: flex; align-items: center; gap: 8px; background: #fff; border: 1px solid #E2E5E2; border-radius: 10px; padding: 0 12px; }
        .search-box:focus-within { border-color: #1AAE86; }
        .search-box input { border: none; outline: none; background: transparent; padding: 10px 0; font-size: 13px; width: 160px; font-family: 'Syne', sans-serif; color: #1F2421; }
        .search-box svg { color: #A2A8A4; }
        .tbtn {
          display: flex; align-items: center; gap: 7px; padding: 10px 15px; border-radius: 10px;
          font-size: 13px; font-weight: 700; font-family: 'Syne', sans-serif; border: 1px solid transparent; transition: filter .15s;
        }
        .tbtn:hover { filter: brightness(.97); }
        .tbtn-hist { background: #EFEDF9; border-color: #D9D4F0; color: #5B4FC4; }
        .tbtn-word { background: #E7F0FB; border-color: #C3DBF4; color: #1F66B0; }
        .tbtn-excel { background: #E7F4EA; border-color: #BFE2C6; color: #2F8F46; }
        .tbtn-clear { background: #FCEEEE; border-color: #F1C7C7; color: #C0392B; }
        .tbtn-snap { background: #FBF3E3; border-color: #ECD9AE; color: #9A6B12; }
        .tbtn-restore { background: #E6F3F1; border-color: #BDE0DA; color: #0F6E56; }

        .ctable { background: #fff; border: 1px solid #E7E9E7; border-radius: 16px; overflow: hidden; }
        .ctable-head, .ctable-row { display: grid; grid-template-columns: 200px 1fr 120px 160px 90px; align-items: center; }
        .ctable-head {
          padding: 13px 22px; background: #FAFBFA; border-bottom: 1px solid #EEF0EE;
          font-size: 11px; letter-spacing: 1px; text-transform: uppercase; color: #9AA0A6; font-weight: 700;
        }
        .ctable-row { padding: 14px 22px; border-bottom: 1px solid #F2F4F2; transition: background .15s; }
        .ctable-row:last-child { border-bottom: none; }
        .ctable-row:hover { background: #FAFBFA; }
        .ccode { font-family: 'IBM Plex Mono', monospace; font-weight: 700; font-size: 14px; }
        .cintitule { font-size: 13.5px; color: #39403B; }
        .cbadge { display: inline-flex; padding: 3px 9px; border-radius: 7px; font-family: 'IBM Plex Mono', monospace; font-size: 11px; font-weight: 700; }
        .cdate { font-family: 'IBM Plex Mono', monospace; font-size: 11.5px; color: #9AA0A6; }
        .cdel { width: 30px; height: 30px; border-radius: 8px; border: 1px solid #E2E5E2; background: #fff; color: #9AA0A6; display: flex; align-items: center; justify-content: center; }
        .cdel:hover { background: #FCEEEE; border-color: #F1C7C7; color: #C0392B; }

        .empty-state { padding: 54px 20px; text-align: center; }
        .empty-ic { width: 64px; height: 64px; border-radius: 50%; background: #F2F4F2; color: #B9BFBA; display: flex; align-items: center; justify-content: center; margin: 0 auto 16px; }
        .empty-t { font-size: 14px; font-weight: 700; color: #5C635E; }
        .empty-s { font-size: 12.5px; color: #A2A8A4; margin-top: 5px; }

        @media (max-width: 920px) {
          .dcg-shell { grid-template-columns: 1fr; }
          .dcg-side { border-right: none; border-bottom: 1px solid #E7E9E7; }
          .stats-grid, .stats-grid-3 { grid-template-columns: repeat(2,1fr); }
          .ctable-head, .ctable-row { grid-template-columns: 130px 1fr 80px 70px; }
          .ctable-head span:nth-child(4), .ctable-row span:nth-child(4) { display: none; }
        }
      `}</style>

      <div className="dcg-shell">
        {/* ===================== SIDEBAR ===================== */}
        <aside className="dcg-side">
          <div className="side-logo">
            <div className={"side-logo-mark" + (BRAND_LOGO ? " has-img" : "")}>
              {BRAND_LOGO ? <img src={BRAND_LOGO} alt="DocCode" /> : "D"}
            </div>
            <div className="side-logo-txt">
              <b>CoDoc <span>Generator</span></b>
              <small>QMS / ISO</small>
            </div>
          </div>

          <div className="side-head"><IcFile size={15} /> Nouveau document</div>

          <div className="fld">
            <label className="fld-label">Type de document</label>
            <Combobox
              placeholder="-- Sélectionner --"
              searchPlaceholder="Tapez une lettre…"
              value={type}
              onChange={(v) => { setType(v); setCustomType(""); setCustomCode(""); setDocModel(""); }}
              groups={[{ label: null, items: [
                { value: "processus", label: "Processus (PS)" },
                { value: "procedure", label: "Procédure (PR)" },
                { value: "enregistrement", label: "Enregistrement (ENG)" },
                { value: "formulaire", label: "Formulaire (F)" },
                { value: "logigramme", label: "Logigramme (L)" },
                { value: "instruction", label: "Instruction (INS)" },
                { value: "modeoperatoire", label: "Mode opératoire (MO)" },
                { value: "__autre__", label: "✏️ Autre type (personnalisé)..." },
              ] }]}
            />

            {type === "__autre__" && (
              <div className="autre-box">
                <div className="autre-cap">Définir un nouveau type</div>
                <div style={{ marginBottom: "12px" }}>
                  <label className="fld-label">Nom du type</label>
                  <input className="fld-input" placeholder="ex: Rapport d'audit"
                    value={customType}
                    onChange={e => { const v = e.target.value; setCustomType(v); setCustomCode(autoGenerateCode(v)); }} />
                </div>
                <label className="fld-label">Code (max 3 lettres)</label>
                <input className="autre-code-input" maxLength={3} value={customCode}
                  onChange={e => setCustomCode(e.target.value.toUpperCase().replace(/[^A-Z]/g, ""))} />
                <div className="autre-tag">↑ généré automatiquement — modifiable</div>
              </div>
            )}
          </div>

          {type && type !== "__autre__" && DOC_CATALOG[type] && (
            <div className="fld">
              <label className="fld-label">Document (référentiel)</label>
              <Combobox
                placeholder="-- Choisir un document --"
                searchPlaceholder="Rechercher un document…"
                value={docModel}
                onChange={(v) => { setDocModel(v); if (v === "__autre__") setIntitule(""); else if (v !== "") setIntitule(v); }}
                groups={[
                  ...DOC_CATALOG[type].map((grp) => ({
                    label: grp.group,
                    items: grp.items.map((it) => ({ value: it, label: it })),
                  })),
                  { label: null, items: [{ value: "__autre__", label: "✏️ Autre (saisie libre)…" }] },
                ]}
              />
            </div>
          )}

          <div className="fld">
            <label className="fld-label">Intitulé du document</label>
            <input className="fld-input" placeholder="ex: Gestion des non conformités"
              value={intitule} onChange={e => setIntitule(e.target.value)}
              onKeyDown={e => e.key === "Enter" && handleSubmit()} />
          </div>

          <div className="preview-box">
            <div className="preview-cap">Aperçu du code</div>
            <div className="preview-code" style={{ color: selectedColor }}>{preview}</div>
          </div>

          <button className="btn-generate" onClick={handleSubmit}>
            <IcSparkles size={18} /> Générer le code
          </button>
          <div className="gen-hint"><IcInfo size={14} /> Remplissez les champs pour générer un code.</div>

          <div className="side-head"><IcBookOpen size={15} /> Référence</div>
          <div className="ref-list">
            {Object.entries(TYPE_LABELS).map(([key, val]) => {
              const Ic = TYPE_ICONS[key];
              return (
                <div className="ref-row" key={key}>
                  <div className="ref-ic" style={{ background: `${val.color}1A`, color: val.color }}><Ic size={16} /></div>
                  <span className="ref-name">{val.label}</span>
                  <span className="ref-badge" style={{ background: `${val.color}18`, color: val.color }}>{val.code}</span>
                </div>
              );
            })}
          </div>

          <button className="btn-types" onClick={() => setShowTypes(true)}>
            Voir la description des types <IcChevronRight size={16} />
          </button>
        </aside>

        {/* ===================== MAIN ===================== */}
        <main className="dcg-main">
          <div className="topbar">
            <div>
              <div className="topbar-title">CoDoc Generator</div>
              <div className="topbar-sub">Système de codification documentaire — QMS/ISO</div>
            </div>
            <div className="user-wrap">
              <div className="user-meta">
                <div className="user-name">{user.name}</div>
                <div className="user-email">{user.email}</div>
              </div>
              <div className="user-avatar" style={{ background: user.logo ? "#fff" : "#0F6E56" }}
                onClick={() => setShowSettings(true)} title="Logo de l'entreprise">
                {user.logo
                  ? <img src={user.logo} alt="logo" style={{ width: "100%", height: "100%", objectFit: "contain" }} />
                  : initials}
              </div>
              <button className="btn-logout" onClick={onLogout}><IcLogout size={16} /> Déconnexion</button>
            </div>
          </div>

          {/* Bandeau de bienvenue */}
          <div className="welcome">
            <div className="welcome-left">
              {user.logo && (
                <div className="welcome-logo"><img src={user.logo} alt="logo" /></div>
              )}
              <div>
                <div className="welcome-h">Bonjour, {user.company} 👋</div>
                <div className="welcome-p">
                  {isFirstTime
                    ? "Bienvenue ! Votre espace est vierge — créez votre premier document."
                    : `Heureux de vous revoir, ${user.name}.`}
                </div>
              </div>
            </div>
            <div className="welcome-pill">
              <IcLayers size={16} /> {history.length} code{history.length !== 1 ? "s" : ""} au total
            </div>
          </div>

          {/* Aperçu par type */}
          <div className="sec-head-row">
            <div className="sec-head"><span className="sec-ic"><IcTrendUp size={18} /></span> Aperçu par type de document</div>
            <div className="year-chip"><IcCalendar size={15} /> Cette année</div>
          </div>

          <div className="stats-grid">
            {["processus", "procedure", "enregistrement", "formulaire"].map(statCard)}
          </div>
          <div className="stats-grid-3">
            {["logigramme", "instruction", "modeoperatoire"].map(statCard)}
          </div>

          {/* Codes enregistrés */}
          <div className="codes-head">
            <div>
              <div className="codes-title">Codes enregistrés</div>
              <div className="codes-sub">{codes.length} code{codes.length !== 1 ? "s" : ""} enregistré{codes.length !== 1 ? "s" : ""}</div>
            </div>
            <div className="toolbar">
              <div className="search-box">
                <IcSearch size={15} />
                <input placeholder="Rechercher..." value={query} onChange={e => setQuery(e.target.value)} />
              </div>
              <button className="tbtn tbtn-hist" onClick={() => setShowHistory(true)}>
                <IcHistory size={15} /> Historique{history.length ? ` (${history.length})` : ""}
              </button>
              <button className="tbtn tbtn-snap" onClick={() => setShowExports(true)}>
                <IcArchive size={15} /> Tableaux exportés{snapshots.length ? ` (${snapshots.length})` : ""}
              </button>
              <button className="tbtn tbtn-word" onClick={handleExportWord}><IcFileText size={15} /> Exporter Word</button>
              <button className="tbtn tbtn-excel" onClick={handleExportExcel}><IcFileText size={15} /> Exporter Excel</button>
              <button className="tbtn tbtn-clear" onClick={handleClearAll}><IcTrash size={15} /> Tout effacer</button>
            </div>
          </div>

          <div className="ctable">
            <div className="ctable-head">
              <span>Code</span><span>Intitulé</span><span>Type</span><span>Date</span><span>Actions</span>
            </div>

            {codes.length === 0 ? (
              <div className="empty-state">
                <div className="empty-ic"><IcFileSearch size={28} /></div>
                <div className="empty-t">Aucun code enregistré pour le moment.</div>
                <div className="empty-s">Générez votre premier code pour le voir apparaître ici.</div>
              </div>
            ) : filtered.length === 0 ? (
              <div className="empty-state">
                <div className="empty-ic"><IcSearch size={26} /></div>
                <div className="empty-t">Aucun résultat</div>
                <div className="empty-s">Aucun code ne correspond à « {query} ».</div>
              </div>
            ) : filtered.map((c) => {
              const typeInfo = Object.values(TYPE_LABELS).find(t => t.code === c.type);
              const color = typeInfo?.color || "#888780";
              return (
                <div className="ctable-row" key={c.id}>
                  <span className="ccode" style={{ color }}>{c.code}</span>
                  <span className="cintitule">{c.intitule}</span>
                  <span><span className="cbadge" style={{ background: `${color}18`, color }}>{c.type}</span></span>
                  <span className="cdate">{new Date(c.createdAt).toLocaleString("fr-FR", { dateStyle: "short", timeStyle: "short" })}</span>
                  <span><button className="cdel" onClick={() => handleDelete(c.id)} title="Supprimer"><IcTrash size={14} /></button></span>
                </div>
              );
            })}
          </div>
        </main>
      </div>

      {/* ===== Fenêtre Logo / Paramètres ===== */}
      {showSettings && (
        <div style={{ ...styles.modalOverlay, animation: "fadeIn 0.2s ease" }} onClick={() => setShowSettings(false)}>
          <div style={{ ...styles.modal, maxWidth: "460px" }} onClick={(e) => e.stopPropagation()}>
            <div style={styles.modalHead}>
              <div>
                <div style={styles.modalTitle}>🏢 Logo de l'entreprise</div>
                <div style={styles.modalSub}>{user.company}</div>
              </div>
              <button style={styles.closeX} onClick={() => setShowSettings(false)} title="Fermer">×</button>
            </div>
            <div style={{ padding: "28px", textAlign: "center" }}>
              <div style={styles.logoPreviewBox}>
                {user.logo
                  ? <img src={user.logo} alt="logo" style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain" }} />
                  : <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: "#B4B2A9" }}>Aucun logo</span>}
              </div>
              <div>
                <label style={styles.logoUploadBtn}>
                  {user.logo ? "Changer le logo" : "Importer un logo"}
                  <input type="file" accept="image/*" style={{ display: "none" }} onChange={handleLogoUpload} />
                </label>
                {user.logo && <button style={styles.logoRemoveBtn} onClick={handleLogoRemove}>Supprimer</button>}
              </div>
              <p style={{ fontSize: "11px", color: "#888780", marginTop: "18px", lineHeight: 1.6 }}>
                PNG ou JPG. L'image est redimensionnée automatiquement.
                Le logo apparaît dans l'en-tête, l'accueil et les exports Word / Excel.
              </p>

              <div style={{ borderTop: "1px solid #EEE9E0", margin: "22px 0 16px" }} />
              <div style={{ fontSize: "11px", fontFamily: "'IBM Plex Mono', monospace", color: "#A53030", letterSpacing: "1px", textTransform: "uppercase", marginBottom: "10px" }}>
                Zone dangereuse
              </div>
              <button
                onClick={handleDeleteAccount}
                style={{
                  display: "inline-flex", alignItems: "center", justifyContent: "center", gap: "8px",
                  padding: "12px 22px", borderRadius: "9px", cursor: "pointer",
                  border: "1px solid #E8B4B4", background: "#FBEAEA", color: "#A53030",
                  fontFamily: "'Syne', sans-serif", fontSize: "14px", fontWeight: "800",
                }}>
                🗑️ Supprimer ce compte
              </button>
              <p style={{ fontSize: "10.5px", color: "#B4B2A9", marginTop: "10px", lineHeight: 1.6 }}>
                Efface définitivement le compte, ses codes et son historique sur ce navigateur. Irréversible.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* ===== Fenêtre Description des types ===== */}
      {showTypes && (
        <div style={{ ...styles.modalOverlay, animation: "fadeIn 0.2s ease" }} onClick={() => setShowTypes(false)}>
          <div style={{ ...styles.modal, maxWidth: "560px" }} onClick={(e) => e.stopPropagation()}>
            <div style={styles.modalHead}>
              <div>
                <div style={styles.modalTitle}>📚 Description des types de documents</div>
                <div style={styles.modalSub}>Codification &amp; rôle de chaque type</div>
              </div>
              <button style={styles.closeX} onClick={() => setShowTypes(false)} title="Fermer">×</button>
            </div>
            <div style={styles.modalBody}>
              {Object.entries(TYPE_LABELS).map(([key, val]) => (
                <div key={key} style={{ display: "flex", gap: "14px", padding: "16px 24px", borderBottom: "1px solid #F1EFE8", alignItems: "flex-start" }}>
                  <span style={badge(val.color)}>{val.code}</span>
                  <div>
                    <div style={{ fontWeight: 800, fontSize: "14px", color: "#2C2C2A" }}>{val.label}</div>
                    <div style={{ fontSize: "12.5px", color: "#6C6F6A", marginTop: "3px", lineHeight: 1.5 }}>{TYPE_DESCRIPTIONS[key]}</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ===== Fenêtre Historique ===== */}
      {showHistory && (
        <div style={{ ...styles.modalOverlay, animation: "fadeIn 0.2s ease" }} onClick={() => setShowHistory(false)}>
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <div style={styles.modalHead}>
              <div>
                <div style={styles.modalTitle}>🕑 Historique des codes générés</div>
                <div style={styles.modalSub}>Archive permanente de {user.company} — conservée même après « Tout effacer »</div>
              </div>
              <button style={styles.closeX} onClick={() => setShowHistory(false)} title="Fermer">×</button>
            </div>
            <div style={styles.modalBody}>
              {history.length === 0 ? (
                <div style={styles.empty}>Aucun code dans l'historique pour le moment</div>
              ) : (
                <>
                  <div style={styles.histHeaderRow}>
                    <span>Code</span><span>Intitulé</span><span>Type</span><span>Date</span>
                  </div>
                  {history.map((c) => {
                    const typeInfo = Object.values(TYPE_LABELS).find(t => t.code === c.type);
                    const color = typeInfo?.color || "#888780";
                    return (
                      <div key={c.id} style={styles.histRow}>
                        <span style={{ ...styles.codeBadge, color }}>{c.code}</span>
                        <span style={{ fontSize: "13px", color: "#444441" }}>{c.intitule}</span>
                        <span><span style={{ ...styles.typeBadge, background: `${color}18`, color }}>{c.type}</span></span>
                        <span style={styles.dateText}>{new Date(c.createdAt).toLocaleString("fr-FR", { dateStyle: "short", timeStyle: "short" })}</span>
                      </div>
                    );
                  })}
                </>
              )}
            </div>
            <div style={styles.modalFoot}>
              <div style={{ fontSize: "12px", fontFamily: "'IBM Plex Mono', monospace", color: "#888780" }}>
                {history.length} entrée{history.length !== 1 ? "s" : ""} au total
              </div>
              <button style={styles.clearBtn} onClick={handleClearHistory}>🗑️ Vider l'historique</button>
            </div>
          </div>
        </div>
      )}

      {/* ===== Fenêtre Tableaux exportés (instantanés) ===== */}
      {showExports && (
        <div style={{ ...styles.modalOverlay, animation: "fadeIn 0.2s ease" }} onClick={() => setShowExports(false)}>
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <div style={styles.modalHead}>
              <div>
                <div style={styles.modalTitle}>📦 Tableaux exportés</div>
                <div style={styles.modalSub}>Chaque export Word/Excel sauvegarde une copie du tableau. Restaurez-la pour la modifier puis réexporter.</div>
              </div>
              <button style={styles.closeX} onClick={() => setShowExports(false)} title="Fermer">×</button>
            </div>
            <div style={styles.modalBody}>
              {snapshots.length === 0 ? (
                <div style={styles.empty}>
                  Aucun tableau exporté pour le moment.<br />
                  Exportez en Word ou Excel : une sauvegarde sera créée ici automatiquement.
                </div>
              ) : snapshots.map((s) => (
                <div key={s.id} style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  gap: "12px", padding: "14px 24px", borderBottom: "1px solid #F1EFE8", flexWrap: "wrap",
                }}>
                  <div style={{ minWidth: "180px" }}>
                    <div style={{ fontWeight: 800, fontSize: "13.5px", color: "#2C2C2A" }}>
                      {new Date(s.createdAt).toLocaleString("fr-FR", { dateStyle: "short", timeStyle: "short" })}
                    </div>
                    <div style={{ fontSize: "11.5px", color: "#888780", fontFamily: "'IBM Plex Mono', monospace", marginTop: "2px" }}>
                      {s.count} code{s.count !== 1 ? "s" : ""} · {s.format === "word" ? "Word" : "Excel"}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
                    <button className="tbtn tbtn-restore" onClick={() => handleRestoreSnapshot(s)}>
                      <IcRestore size={15} /> Restaurer
                    </button>
                    <button className="tbtn tbtn-word" onClick={() => handleReexport(s, "word")}><IcFileText size={15} /> Word</button>
                    <button className="tbtn tbtn-excel" onClick={() => handleReexport(s, "excel")}><IcFileText size={15} /> Excel</button>
                    <button className="cdel" onClick={() => handleDeleteSnapshot(s.id)} title="Supprimer"><IcTrash size={14} /></button>
                  </div>
                </div>
              ))}
            </div>
            <div style={styles.modalFoot}>
              <div style={{ fontSize: "12px", fontFamily: "'IBM Plex Mono', monospace", color: "#888780" }}>
                {snapshots.length} tableau{snapshots.length !== 1 ? "x" : ""} sauvegardé{snapshots.length !== 1 ? "s" : ""}
              </div>
              <button style={styles.clearBtn} onClick={handleClearSnapshots}>🗑️ Tout vider</button>
            </div>
          </div>
        </div>
      )}

      {toast && <div style={styles.toast}>{toast}</div>}
    </div>
  );
}

/* =========================================================================
   ROUTEUR : connexion ↔ application
   ========================================================================= */
export default function App() {
  const [session, setSession]   = useState(() => loadSession());
  const [accounts, setAccounts] = useState(() => loadAccounts());

  // Au démarrage, vérifier que la session pointe vers un compte existant
  useEffect(() => {
    if (session && !accounts[session]) { clearSession(); setSession(null); }
  }, [session, accounts]);

  const user = session && accounts[session] ? accounts[session] : null;

  const handleAuthenticated = (email) => {
    setAccounts(loadAccounts()); // recharge (le compte vient peut-être d'être créé)
    saveSession(email);
    setSession(email);
  };
  const handleLogout = () => { clearSession(); setSession(null); };

  // Suppression définitive du compte courant + toutes ses données
  const deleteAccount = () => {
    if (!session) return;
    const email = session;
    setAccounts((prev) => {
      const next = { ...prev };
      delete next[email];
      saveAccounts(next);
      return next;
    });
    try { localStorage.removeItem(stateKey(email)); } catch (e) {}
    try { localStorage.removeItem(histKey(email)); } catch (e) {}
    try { localStorage.removeItem(snapKey(email)); } catch (e) {}
    clearSession();
    setSession(null);
  };

  // Mise à jour du compte courant (ex: logo) — persistée + reflétée à l'écran
  const updateUser = (patch) => {
    setAccounts((prev) => {
      if (!session || !prev[session]) return prev;
      const next = { ...prev, [session]: { ...prev[session], ...patch } };
      saveAccounts(next);
      return next;
    });
  };

  if (!user) return <AuthScreen onAuthenticated={handleAuthenticated} />;
  return <DocCodeApp key={user.email} user={user} onLogout={handleLogout} onUpdateUser={updateUser} onDeleteAccount={deleteAccount} />;
}