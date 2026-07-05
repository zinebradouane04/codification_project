import React, { useState, useEffect, useRef } from "react";
import ExcelJS from "exceljs";
import brandLogo from "./logo_codoc.png";
import logoOcg from "./logo_ocg.png";
import {
  Document, Packer, Table, TableRow, TableCell, Paragraph, TextRun, ImageRun,
  WidthType, AlignmentType, BorderStyle, VerticalAlign, ShadingType,
  PageOrientation,
} from "docx";

const BRAND_LOGO = brandLogo;

/* ----------------------- AUTHENTIFICATION (locale) ----------------------- */
const ACCOUNTS_KEY = "doccode_accounts_v1";
const SESSION_KEY  = "doccode_session_v1";

function loadAccounts() {
  try { return JSON.parse(localStorage.getItem(ACCOUNTS_KEY)) || {}; } catch (e) { return {}; }
}
function saveAccounts(acc) { try { localStorage.setItem(ACCOUNTS_KEY, JSON.stringify(acc)); } catch (e) {} }
function loadSession() { try { return localStorage.getItem(SESSION_KEY) || null; } catch (e) { return null; } }
function saveSession(email) { try { localStorage.setItem(SESSION_KEY, email); } catch (e) {} }
function clearSession() { try { localStorage.removeItem(SESSION_KEY); } catch (e) {} }

function hashPassword(pw) {
  let h = 0;
  for (let i = 0; i < pw.length; i++) h = (Math.imul(31, h) + pw.charCodeAt(i)) | 0;
  return "h" + (h >>> 0).toString(36);
}

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
        canvas.getContext("2d").drawImage(img, 0, 0, w, h);
        try { resolve({ dataUrl: canvas.toDataURL("image/png"), width: w, height: h }); } catch (e) { reject(e); }
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function dataUrlToBase64(dataUrl) {
  if (!dataUrl || typeof dataUrl !== "string") return null;
  const i = dataUrl.indexOf("base64,");
  return i === -1 ? null : dataUrl.slice(i + 7);
}
function base64ToUint8(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
function scaleDims(w, h, maxW, maxH) {
  if (!w || !h) return { width: maxW, height: maxH };
  const r = Math.min(maxW / w, maxH / h, 1);
  return { width: Math.max(1, Math.round(w * r)), height: Math.max(1, Math.round(h * r)) };
}

/* ------------------ DONNÉES PAR UTILISATEUR (localStorage) ---------------- */
const stateKey = (email) => `doccode_state_v1__${email}`;
const histKey  = (email) => `doccode_history_v1__${email}`;
const snapKey  = (email) => `doccode_snapshots_v1__${email}`;
const cartoKey = (email) => `doccode_carto_v1__${email}`;

function loadUserState(email) {
  try { const raw = localStorage.getItem(stateKey(email)); if (raw) return JSON.parse(raw); } catch (e) {}
  return null;
}
function loadUserHistory(email) {
  try { const raw = localStorage.getItem(histKey(email)); if (raw) return JSON.parse(raw); } catch (e) {}
  return [];
}
function loadUserSnapshots(email) {
  try { const raw = localStorage.getItem(snapKey(email)); if (raw) return JSON.parse(raw); } catch (e) {}
  return [];
}
function loadUserCarto() {
  // La cartographie repart vierge à chaque session (pas de persistence)
  return {
    management: [],
    realisation: [],
    support: [],
    inputs: "Exigences Clients\nRéglementaires et légales\nAnalyse des risques",
    outputs: "Satisfaction client\nPerformance SMI\nSécurité de l'information\nConformité réglementaire",
  };
}

/* =========================================================================
   CATALOGUE DES PROCESSUS (pour la Cartographie)
   ========================================================================= */
const PROCESSUS_CATALOG = {
  management: [
    "Pilotage stratégique",
    "Pilotage SMI",
    "Management de la Sécurité de l'Information",
    "Management du système intégré",
    "Gestion des risques et opportunités",
    "Revue de direction",
    "Veille réglementaire",
    "Amélioration continue",
  ],
  realisation: [
    "Développement Commercial et Gestion des Contrats",
    "Étude et Conception des Solutions",
    "Achats et Gestion des Fournisseurs",
    "Intégration et Déploiement des Solutions",
    "Maintenance, Support Technique et SAV",
    "Gestion commerciale",
    "Gestion des appels d'offres",
    "Conception et développement",
    "Production / Réalisation du service",
    "Contrôle qualité",
    "Logistique",
    "Livraison",
    "Service après-vente",
    "Gestion des réclamations",
  ],
  support: [
    "RH & compétences",
    "Gestion des Infrastructures IT",
    "Gestion des Actifs Informationnels",
    "Moyens généraux & logistique",
    "Gestion Financière et Comptable",
    "Ressources humaines",
    "Formation",
    "Maintenance",
    "Système d'information",
    "Finance et comptabilité",
    "Gestion documentaire",
    "Métrologie",
    "Gestion des infrastructures",
  ],
};


/* =========================================================================
   EXPORT WORD CARTOGRAPHIE v5 — toutes les flèches + logo du compte connecté
   ========================================================================= */

// Couleurs exactes Koutoubia
const CK_BLUE_DARK   = "1F3864";
const CK_BLUE_MED    = "2E6FB5";
const CK_BLUE_SIDE   = "D6E8F5";
const CK_GREEN_DARK  = "2E7D32";
const CK_GREEN_LIGHT = "E7F4E8";
const CK_ORANGE_DARK = "E2671C";
const CK_ORANGE_LIGHT= "FBE7D7";
const CK_BLUE_SUPP   = "1565C0";
const CK_BLUE_SUPP_L = "E3F0FA";
const CK_W           = "FFFFFF";

// Dimensions exactes Koutoubia (DXA)
const CK_TOTAL_W   = 15704;
const CK_SIDE_L_W  = 1977;
const CK_ARROW_L_W = 977;
const CK_INNER_W   = 9804;
const CK_ARROW_R_W = 984;
const CK_SIDE_R_W  = 1962;
const CK_ARROW_INN = 320;

// Flèches PNG encodées en base64 (grandes flèches bleues + petites flèches transition)
const PNG_ARROW_R = "iVBORw0KGgoAAAANSUhEUgAAAFAAAAA8CAYAAADxJz2MAAABXElEQVR4nO3by63CMBBG4ZMr6qERiqEWiqERGoIFspRFHn4Pd+Y/+yjRJ0exxgoopZRd1/vzbXn/P8ub98oS0QUg2CG6AQQbRFeA8EWcCekOMDUL0S0gzEF0DQjjEd0DwljEEIAwDjEMIIxBDAUI/bc54QBTvRDDAkIfxNCA0I4YHhDaEBfredov9XrcltJrtAJX1XyhBbhRCaIAd8pFFOBBOYgCPOkMUYAZHSEKMLM9RAEWtLXNEWBFa0QBVpYQBdjQ9f58C7Ch1+O2CLCyNHgQYEXrqc3F8kH+W1vjLq3AzPZmhQLM6GjQKsCTzqbUAjwoZ8QvwJ1yz0cEuFHJ4ZK2MatqTuWKL/i1eh3L1uCBXmGgHg8E2IQHwQFb8SAwYA88CPgV7gWXCrUCe+NBIMAReBAEcBQeBAAciQfOAUfjgWPAGXjgcBszCy7lagXOxgNHgBZ4LtJfBip2Hx7HfKhiQK/HAAAAAElFTkSuQmCC";
const PNG_ARROW_L = "iVBORw0KGgoAAAANSUhEUgAAAFAAAAA8CAYAAADxJz2MAAABYklEQVR4nO3b2Q3CMBBF0ReLemgkxVBLiqGRNAQfUUCRs3gZj5WZd/9R4ODIWWyAMcbKe77en97foabQ8+B3xwM6AlrAA4CH9gGtwK2pjkBreIAioEU8QAnQKh6gAGgZD2gMaB0PaDQLe4BbEx+BnvAAYUBveIAgoEc8QAjQKx5QOYl4hlsrHoHEWxpKPkS8f9kjkHjbsgCJF5d0ChPuuMsRSLzzTgGJd90hIPHS2gUkXnoRIPHy+s3ChCsrAMSrKRCvrjBPY9H9MFsKAEDE8iI4ntJ5RZcxHI157V5IEzG9w1s5IqZ1+jCBiNddPs4i4nlZOJyh47Ie6XM0xmW/VCLitmIMydP5zn9K8Yv1O/9oyarWxszTOHiHFFlc5BlRbHmbV0TRBZYeEcWX+HpDbLJK39Pk0nSbgwfE5httrCOqbPWyjKi22dAqoup2V4uI6jvWrc3Q3bb8W0LsGp9yM1bTF8QdeAAGxLADAAAAAElFTkSuQmCC";
const PNG_ARROW_D = "iVBORw0KGgoAAAANSUhEUgAAADAAAAA8CAYAAAAgwDn8AAABFElEQVR4nO2aQQ6DIBBFRw7kovfryvu58ELtqomxIAzMzMfkv12bCP+lCjNYEYJl6b1wfb8+lkFERI5tV+dJ1iGioQAaCqChABoKoKEAGgqgoQAaCqChABoKoKEAGgqgoQAaCqChAJrHCywiPi8rokgifW9GZuDY9iWdPyDDaPnlTbkvZ+ec8+8hnl3imi+7Cs0qkctVXEZnkyjlud0HZpG4y1HdyNAStfmbdmKURMu8zaVEtETrfKpaKEpCM4+6mPOW0I7fVY16SYT+W8Vaone8oX7ASmJknOGGZlRi9HqTjqw3hMUvaNZSasNY3X6mPXFrKMsFwLypr4WzXr1cTiVKIT32D7djlWtYdFXbzZOPbEL4Ai56ZleQ/qHwAAAAAElFTkSuQmCC";
const PNG_ARROW_U = "iVBORw0KGgoAAAANSUhEUgAAADAAAAA8CAYAAAAgwDn8AAAA/ElEQVR4nO2Y0Q2DMAwFITt0ny7TqbpM9+kQ7ZclRJMQO3ZeUr37AxTnTkKA2Dby59wer0/k/BQ5XOQjI8ICztJRESEBJdmICPeAK0nvCNeAVjnPCLcArZRXhEuAVcYjojugV6J3fVeA123QM8cc4P00sc4zBUS9lCxz1QHR3zba+aqAaHnLPs0Bo+S1+zUFjJbX7HsZgJJv3b8agJYXah7FgFnkhZJPNmA2eSHn9RMwq7xw9ku1i7Ny9Ey5kysgvvvxYEV268KI6PfzrvYJ/S80AgagYQAaBqBhABoGoGEAGgagYQAaBqBhABoGoGEAGgagYQAaBqBZPmB5vgX/aq2OTr0YAAAAAElFTkSuQmCC";

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function ck5_nil()        { return { style: BorderStyle.NIL, size: 0, color: CK_W }; }
function ck5_solid(c,s=4) { return { style: BorderStyle.SINGLE, size: s, color: c }; }
function ck5_b64Uint8(b64){ const bin=atob(b64);const b=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)b[i]=bin.charCodeAt(i);return b;}

function ck5_emptyC(w, fill=CK_W) {
  return new TableCell({width:{size:w,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill},
    borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil()},
    margins:{top:0,bottom:0,left:0,right:0},children:[new Paragraph({children:[]})]});
}

function ck5_titleBand(text, fill, icon="") {
  return new TableCell({width:{size:CK_INNER_W,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill},
    borders:{top:ck5_solid(fill,8),bottom:ck5_solid(fill,8),left:ck5_solid(fill,8),right:ck5_solid(fill,8)},
    verticalAlign:VerticalAlign.CENTER,margins:{top:80,bottom:80,left:120,right:120},
    children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
      ...(icon?[new TextRun({text:icon+"  ",font:"Segoe UI Emoji",size:20,color:CK_W})]:[]),
      new TextRun({text,font:"Arial",size:20,bold:true,color:CK_W}),
    ]})]});
}

function ck5_procBox(num, name, fill, border, width) {
  return new TableCell({width:{size:width,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill},
    borders:{top:ck5_solid(border,8),bottom:ck5_solid(border,8),left:ck5_solid(border,8),right:ck5_solid(border,8)},
    verticalAlign:VerticalAlign.CENTER,margins:{top:100,bottom:100,left:80,right:80},
    children:[
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:30,after:10},children:[new TextRun({text:num,font:"Arial",size:18,bold:true,color:border})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:30},children:[new TextRun({text:name,font:"Arial",size:16,color:"1F1F1F"})]}),
    ]});
}

function ck5_arrowBetween(bgFill,w=CK_ARROW_INN){
  return new TableCell({width:{size:w,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:bgFill},
    borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil()},
    verticalAlign:VerticalAlign.CENTER,margins:{top:0,bottom:0,left:0,right:0},
    children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({text:"\u279C",font:"Arial",size:22,bold:true,color:CK_BLUE_MED})]})]});
}

function ck5_sideCol(lines,nRows,fill,borderC,textColor,width){
  return new TableCell({width:{size:width,type:WidthType.DXA},rowSpan:nRows,
    shading:{type:ShadingType.CLEAR,fill},
    borders:{top:ck5_solid(borderC,8),bottom:ck5_solid(borderC,8),left:ck5_solid(borderC,8),right:ck5_solid(borderC,8)},
    verticalAlign:VerticalAlign.CENTER,margins:{top:120,bottom:120,left:100,right:100},
    children:lines.map(l=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:30,after:30},
      children:[new TextRun({text:l,font:"Arial",size:16,bold:true,color:textColor})]}))});
}

function ck5_bigArrowCell(b64png,nRows,cellW,imgW,imgH){
  return new TableCell({width:{size:cellW,type:WidthType.DXA},rowSpan:nRows,
    shading:{type:ShadingType.CLEAR,fill:CK_W},
    borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil()},
    verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new ImageRun({type:"png",data:ck5_b64Uint8(b64png),transformation:{width:imgW,height:imgH}})]})]}); 
}

function ck5_transitionRow(b64png,bgFill,nArrows,imgW,imgH){
  const arrowW=imgW+20;
  const spacerW=Math.floor((CK_INNER_W-nArrows*arrowW)/(nArrows+1));
  const colW=[];const cells=[];
  for(let i=0;i<nArrows;i++){
    colW.push(spacerW);cells.push(ck5_emptyC(spacerW,bgFill));
    colW.push(arrowW);
    cells.push(new TableCell({width:{size:arrowW,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:bgFill},
      borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil()},
      verticalAlign:VerticalAlign.CENTER,
      children:[new Paragraph({alignment:AlignmentType.CENTER,
        children:[new ImageRun({type:"png",data:ck5_b64Uint8(b64png),transformation:{width:imgW,height:imgH}})]})]})
    );
  }
  const lastW=Math.max(1,CK_INNER_W-colW.reduce((a,b)=>a+b,0));
  colW.push(lastW);cells.push(ck5_emptyC(lastW,bgFill));
  return new Table({width:{size:CK_INNER_W,type:WidthType.DXA},columnWidths:colW,
    borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil(),insideHorizontal:ck5_nil(),insideVertical:ck5_nil()},
    rows:[new TableRow({height:{value:260,rule:"atLeast"},children:cells})]});
}

function ck5_emptyT(){
  return new Table({width:{size:CK_INNER_W,type:WidthType.DXA},columnWidths:[CK_INNER_W],
    borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil(),insideHorizontal:ck5_nil(),insideVertical:ck5_nil()},
    rows:[new TableRow({children:[ck5_emptyC(CK_INNER_W)]})]});
}

function ck5_procSubTable(items,bgFill,borderColor){
  if(!items.length)return ck5_emptyT();
  const n=items.length;const aW=CK_ARROW_INN;
  const bW=Math.floor((CK_INNER_W-(n-1)*aW)/n);
  const extra=CK_INNER_W-n*bW-(n-1)*aW;
  const colW=[];const cells=[];
  items.forEach((it,i)=>{const w=i===n-1?bW+extra:bW;colW.push(w);
    cells.push(ck5_procBox(it.num,it.name,bgFill,borderColor,w));
    if(i<n-1){colW.push(aW);cells.push(ck5_arrowBetween(bgFill,aW));}
  });
  return new Table({width:{size:CK_INNER_W,type:WidthType.DXA},columnWidths:colW,
    borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil(),insideHorizontal:ck5_nil(),insideVertical:ck5_nil()},
    rows:[new TableRow({height:{value:800,rule:"atLeast"},children:cells})]});
}

function ck5_innerCell(sub,bgFill=CK_W){
  return new TableCell({width:{size:CK_INNER_W,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:bgFill},
    borders:{top:ck5_nil(),bottom:ck5_nil(),left:ck5_nil(),right:ck5_nil()},
    margins:{top:60,bottom:60,left:0,right:0},children:[sub]});
}

async function exportCartoWord(carto, branding={}) {
  const company=(branding.company||"OCG CONSULTING").trim();
  const inputLines =(carto.inputs ||"").split("\n").filter(Boolean);
  const outputLines=(carto.outputs||"").split("\n").filter(Boolean);
  let counter=1;
  const toItems=names=>names.map(n=>({num:`Ps ${String(counter++).padStart(2,"0")}.`,name:n}));
  const mgmtItems=toItems(carto.management||[]);
  const realItems=toItems(carto.realisation||[]);
  const suppItems=toItems(carto.support||[]);
  const TOTAL_ROWS=8;
  const nDown=Math.min(Math.max(mgmtItems.length,2),6);
  const nUp  =Math.min(Math.max(suppItems.length,2),6);
  const rows=[];

  // R1 : Entrées(rs8) | BigArrow→(rs8) | Titre Management | BigArrow←(rs8) | Sorties(rs8)
  rows.push(new TableRow({height:{value:420,rule:"atLeast"},children:[
    ck5_sideCol(inputLines,TOTAL_ROWS,CK_BLUE_SIDE,CK_BLUE_MED,CK_BLUE_DARK,CK_SIDE_L_W),
    ck5_bigArrowCell(PNG_ARROW_R,TOTAL_ROWS,CK_ARROW_L_W,60,46),
    ck5_titleBand("PROCESSUS MANAGEMENT",CK_GREEN_DARK,"\u2699\uFE0F"),
    ck5_bigArrowCell(PNG_ARROW_L,TOTAL_ROWS,CK_ARROW_R_W,60,46),
    ck5_sideCol(outputLines,TOTAL_ROWS,CK_BLUE_SIDE,CK_BLUE_MED,CK_BLUE_DARK,CK_SIDE_R_W),
  ]}));
  // R2 : Boîtes Management
  rows.push(new TableRow({height:{value:780,rule:"atLeast"},children:[ck5_innerCell(ck5_procSubTable(mgmtItems,CK_GREEN_LIGHT,CK_GREEN_DARK),CK_GREEN_LIGHT)]}));
  // R3 : Flèches ↓ Management→Réalisation
  rows.push(new TableRow({height:{value:360,rule:"atLeast"},children:[ck5_innerCell(ck5_transitionRow(PNG_ARROW_D,CK_GREEN_LIGHT,nDown,40,52),CK_GREEN_LIGHT)]}));
  // R4 : Titre Réalisation
  rows.push(new TableRow({height:{value:420,rule:"atLeast"},children:[ck5_titleBand("PROCESSUS R\u00C9ALISATION",CK_ORANGE_DARK,"\uD83C\uDFED")]}));
  // R5 : Boîtes Réalisation
  rows.push(new TableRow({height:{value:780,rule:"atLeast"},children:[ck5_innerCell(ck5_procSubTable(realItems,CK_ORANGE_LIGHT,CK_ORANGE_DARK),CK_ORANGE_LIGHT)]}));
  // R6 : Flèches ↑ Support→Réalisation
  rows.push(new TableRow({height:{value:360,rule:"atLeast"},children:[ck5_innerCell(ck5_transitionRow(PNG_ARROW_U,CK_ORANGE_LIGHT,nUp,40,52),CK_ORANGE_LIGHT)]}));
  // R7 : Titre Support
  rows.push(new TableRow({height:{value:420,rule:"atLeast"},children:[ck5_titleBand("PROCESSUS SUPPORT",CK_BLUE_SUPP,"\u2699\uFE0F")]}));
  // R8 : Boîtes Support
  rows.push(new TableRow({height:{value:780,rule:"atLeast"},children:[ck5_innerCell(ck5_procSubTable(suppItems,CK_BLUE_SUPP_L,CK_BLUE_SUPP),CK_BLUE_SUPP_L)]}));

  const mainTable=new Table({
    width:{size:CK_TOTAL_W,type:WidthType.DXA},
    columnWidths:[CK_SIDE_L_W,CK_ARROW_L_W,CK_INNER_W,CK_ARROW_R_W,CK_SIDE_R_W],
    borders:{top:ck5_solid(CK_BLUE_MED,10),bottom:ck5_solid(CK_BLUE_MED,10),left:ck5_solid(CK_BLUE_MED,10),right:ck5_solid(CK_BLUE_MED,10),insideHorizontal:ck5_nil(),insideVertical:ck5_nil()},
    rows});

  // En-tête — logo du COMPTE CONNECTÉ (logo importé lors de l'inscription ou des paramètres)
  const b64=dataUrlToBase64(branding.logo);
  let logoChildren;
  if(b64){
    const{width,height}=scaleDims(branding.logoW,branding.logoH,200,70);
    logoChildren=[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new ImageRun({type:"png",data:base64ToUint8(b64),transformation:{width,height}})]})];
  } else {
    logoChildren=[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({text:company,font:"Arial",size:22,bold:true,color:CK_BLUE_DARK})]})];
  }

  const headerTable=new Table({
    width:{size:CK_TOTAL_W,type:WidthType.DXA},columnWidths:[2900,9700,3104],
    borders:{top:ck5_solid(CK_BLUE_DARK,12),bottom:ck5_solid(CK_BLUE_DARK,12),left:ck5_solid(CK_BLUE_DARK,12),right:ck5_solid(CK_BLUE_DARK,12),insideHorizontal:ck5_solid(CK_BLUE_DARK,4),insideVertical:ck5_solid(CK_BLUE_DARK,4)},
    rows:[new TableRow({height:{value:700,rule:"atLeast"},children:[
      new TableCell({width:{size:2900,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,shading:{type:ShadingType.CLEAR,fill:CK_W},
        borders:{top:ck5_solid(CK_BLUE_DARK,12),bottom:ck5_solid(CK_BLUE_DARK,12),left:ck5_solid(CK_BLUE_DARK,12),right:ck5_solid(CK_BLUE_DARK,12)},
        margins:{top:60,bottom:60,left:80,right:80},children:logoChildren}),
      new TableCell({width:{size:9700,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,shading:{type:ShadingType.CLEAR,fill:CK_W},
        borders:{top:ck5_solid(CK_BLUE_DARK,12),bottom:ck5_solid(CK_BLUE_DARK,12),left:ck5_solid(CK_BLUE_DARK,12),right:ck5_solid(CK_BLUE_DARK,12)},
        margins:{top:60,bottom:60,left:100,right:100},
        children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"CARTOGRAPHIE DES PROCESSUS",font:"Arial",size:36,bold:true,color:CK_BLUE_DARK})]})]}),
      new TableCell({width:{size:3104,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,shading:{type:ShadingType.CLEAR,fill:CK_W},
        borders:{top:ck5_solid(CK_BLUE_DARK,12),bottom:ck5_solid(CK_BLUE_DARK,12),left:ck5_solid(CK_BLUE_DARK,12),right:ck5_solid(CK_BLUE_DARK,12)},
        margins:{top:60,bottom:60,left:120,right:120},
        children:[
          new Paragraph({spacing:{after:60},children:[new TextRun({text:"Code : ",font:"Arial",size:18,bold:true,color:CK_BLUE_DARK}),new TextRun({text:"____________",font:"Arial",size:18,color:CK_BLUE_DARK})]}),
          new Paragraph({spacing:{after:40},children:[new TextRun({text:"Version : ",font:"Arial",size:18,bold:true,color:CK_BLUE_DARK}),new TextRun({text:"00",font:"Arial",size:18,color:CK_BLUE_DARK})]}),
          new Paragraph({children:[new TextRun({text:"Date : ",font:"Arial",size:18,bold:true,color:CK_BLUE_DARK}),new TextRun({text:new Date().toLocaleDateString("fr-FR"),font:"Arial",size:18,color:CK_BLUE_DARK})]}),
        ]}),
    ]})]});

  const doc=new Document({sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:480,right:480,bottom:480,left:480}}},
    children:[headerTable,new Paragraph({spacing:{before:120,after:0},children:[]}),mainTable]}]});
  downloadBlob(await Packer.toBlob(doc),"cartographie-des-processus.docx");
}



/* =========================================================================
   TYPE_LABELS & utilitaires (inchangés)
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
  modalOverlay: { position:"fixed",inset:0,background:"rgba(44,44,42,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:"24px" },
  modal: { background:"#FFFFFF",borderRadius:"14px",width:"100%",maxWidth:"820px",maxHeight:"82vh",display:"flex",flexDirection:"column",boxShadow:"0 12px 48px rgba(0,0,0,0.22)",overflow:"hidden" },
  modalHead: { display:"flex",justifyContent:"space-between",alignItems:"center",padding:"20px 24px",borderBottom:"1px solid #D3D1C7",background:"#F8F8F6" },
  modalTitle: { fontSize:"16px",fontWeight:"800",color:"#2C2C2A",letterSpacing:"-0.3px" },
  modalSub: { fontSize:"11px",fontFamily:"'IBM Plex Mono', monospace",color:"#888780",marginTop:"3px" },
  closeX: { background:"none",border:"1px solid #D3D1C7",borderRadius:"6px",width:"32px",height:"32px",cursor:"pointer",fontSize:"16px",color:"#888780",display:"flex",alignItems:"center",justifyContent:"center" },
  modalBody: { overflowY:"auto",flex:1 },
  modalFoot: { display:"flex",justifyContent:"space-between",alignItems:"center",padding:"14px 24px",borderTop:"1px solid #D3D1C7",background:"#F8F8F6" },
  histHeaderRow: { display:"grid",gridTemplateColumns:"190px 1fr 90px 150px",padding:"10px 24px",borderBottom:"1px solid #D3D1C7",fontSize:"10px",fontFamily:"'IBM Plex Mono', monospace",color:"#888780",letterSpacing:"2px",textTransform:"uppercase",background:"#FCFCFB",position:"sticky",top:0 },
  histRow: { display:"grid",gridTemplateColumns:"190px 1fr 90px 150px",padding:"12px 24px",borderBottom:"1px solid #F1EFE8",alignItems:"center" },
  codeBadge: { fontFamily:"'IBM Plex Mono', monospace",fontSize:"14px",fontWeight:"600",letterSpacing:"1px" },
  typeBadge: { display:"inline-flex",alignItems:"center",padding:"3px 8px",borderRadius:"5px",fontSize:"10px",fontFamily:"'IBM Plex Mono', monospace",fontWeight:"600",letterSpacing:"1px" },
  dateText: { fontFamily:"'IBM Plex Mono', monospace",fontSize:"11px",color:"#888780" },
  clearBtn: { display:"flex",alignItems:"center",gap:"8px",padding:"9px 18px",borderRadius:"7px",border:"1px solid #E8B4B4",background:"#FBEAEA",color:"#A53030",fontFamily:"'Syne', sans-serif",fontSize:"13px",fontWeight:"700",cursor:"pointer" },
  toast: { position:"fixed",bottom:"28px",right:"28px",background:"#FFFFFF",border:"1px solid #D3D1C7",borderRadius:"10px",padding:"14px 20px",fontFamily:"'IBM Plex Mono', monospace",fontSize:"12px",color:"#0F6E56",zIndex:999,boxShadow:"0 4px 20px rgba(0,0,0,0.1)" },
  empty: { textAlign:"center",padding:"56px",color:"#B4B2A9",fontFamily:"'IBM Plex Mono', monospace",fontSize:"12px" },
  logoPreviewBox: { width:"180px",height:"180px",margin:"0 auto 20px",borderRadius:"14px",border:"1px solid #D3D1C7",background:"#F8F8F6",display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",padding:"14px",boxSizing:"border-box" },
  logoUploadBtn: { display:"inline-flex",alignItems:"center",justifyContent:"center",gap:"8px",padding:"12px 22px",borderRadius:"9px",cursor:"pointer",border:"1px solid #0F6E56",background:"#0F6E56",color:"#FFFFFF",fontFamily:"'Syne', sans-serif",fontSize:"14px",fontWeight:"800" },
  logoRemoveBtn: { display:"inline-flex",alignItems:"center",justifyContent:"center",gap:"8px",padding:"12px 22px",borderRadius:"9px",cursor:"pointer",border:"1px solid #E8B4B4",background:"#FBEAEA",color:"#A53030",fontFamily:"'Syne', sans-serif",fontSize:"14px",fontWeight:"700",marginLeft:"10px" },
};

const TYPE_MAP_BACKEND = {
  processus:"PS", procedure:"PR", enregistrement:"ENG",
  formulaire:"F", logigramme:"L", instruction:"INS", modeoperatoire:"MO",
};

const TYPE_LABELS_PLURAL = {
  PS:"Processus", PR:"Procédures", ENG:"Enregistrements",
  F:"Formulaires", L:"Logigrammes", INS:"Instructions", MO:"Modes opératoires",
};
const TYPE_ORDER = ["PS","PR","ENG","F","L","INS","MO"];

const DOC_CATALOG = {
  processus: [
    { group: "Catalogue OCG", items: ["Management stratégique","Pilotage stratégique","Pilotage SMQ","Pilotage SMI","Pilotage SMSDA","Management de la qualité","Achats","Achats et approvisionnement","Production","Fabrication","Vente","Comptabilité","Comptabilité et finance","Logistique et stockage","Maintenance","Ressources Humaines","Système d'information","Marketing","Management de l'environnement","Santé et sécurité au travail","Commercial","Commercial et gestion des commandes","Contrôle qualité","Contrôle qualité et analyse laboratoire","Amélioration du SMQ","Amélioration SMI","Amélioration SMSDA","Conditionnement et gestion des stocks","Gestion de l'entreprise et développement","Gestion IT"] },
    { group: "Processus de Management", items: ["Pilotage stratégique","Management du système intégré","Gestion des risques et opportunités","Revue de direction","Veille réglementaire","Amélioration continue"] },
    { group: "Processus Réalisation", items: ["Gestion commerciale","Gestion des appels d'offres","Conception et développement","Achats","Production / Réalisation du service","Contrôle qualité","Logistique","Livraison","Service après-vente","Gestion des réclamations"] },
    { group: "Processus Support", items: ["Ressources humaines","Formation","Maintenance","Système d'information","Finance et comptabilité","Gestion documentaire","Métrologie","Gestion des infrastructures"] },
  ],
  procedure: [
    { group: "Catalogue OCG", items: ["Revue de direction","Audit interne","Non-conformité et action corrective","Amélioration continue","Communication interne et externe","Gestion de changement","Gestion des risques et opportunités","Gestion des risques SST","Gestion des aspects et impacts environnementaux","Gestion des objectifs et KPI","Commercial et gestion des commandes","Achats","Sélection et évaluation des fournisseurs","Vente","Comptabilité","Logistique et stockage","Conditionnement","Gestion IT","Retrait / rappel","Gestion des situations d'urgence","Maîtrise des informations documentées","Gestion des compétences et de la formation","Maintenance des équipements","Gestion des ressources humaines","Achats et évaluation des fournisseurs","Contrôle qualité","Gestion des réclamations clients","Mesure de la satisfaction client","Gestion des produits non conformes","Nettoyage et désinfection","Lutte contre les nuisibles","Hygiène du personnel","Gestion des déchets","Stockage et transport","Prévention des contaminations croisées","Traçabilité"] },
    { group: "Management", items: ["Maîtrise des informations documentées","Gestion des risques et opportunités","Audit interne","Revue de direction","Actions correctives","Traitement des non-conformités","Veille réglementaire","Communication interne et externe","Gestion des changements","Gestion des objectifs"] },
    { group: "Qualité — ISO 9001", items: ["Gestion des commandes","Gestion des réclamations","Satisfaction client","Évaluation fournisseurs","Conception et développement","Contrôle qualité","Maîtrise des équipements de mesure"] },
    { group: "Environnement — ISO 14001", items: ["Identification des aspects environnementaux","Gestion des déchets","Gestion des produits chimiques","Gestion des situations d'urgence environnementale","Contrôle des consommations"] },
    { group: "SST — ISO 45001", items: ["Identification des dangers","Évaluation des risques SST","Gestion des accidents","Gestion des EPI","Gestion des entreprises extérieures","Gestion des permis de travail","Gestion des situations d'urgence"] },
    { group: "Sécurité alimentaire — ISO 22000", items: ["HACCP","Gestion des CCP","Gestion des PRP","Gestion des PRPo","Traçabilité","Gestion des rappels","Gestion des allergènes"] },
    { group: "Dispositifs médicaux — ISO 13485", items: ["Gestion des risques ISO 14971","Validation des procédés","Validation des logiciels","Matériovigilance","Traçabilité dispositifs médicaux"] },
    { group: "Sécurité de l'information — ISO 27001", items: ["Gestion des actifs informationnels","Gestion des accès","Gestion des mots de passe","Gestion des sauvegardes","Gestion des incidents de sécurité","Gestion des vulnérabilités","Gestion des fournisseurs IT","Continuité d'activité","Gestion de la cryptographie"] },
  ],
  instruction: [
    { group: "Instructions de travail", items: ["Traitement des commandes","Contrôle qualité","Gestion des déchets","Évacuation d'urgence","Utilisation des EPI","Nettoyage et désinfection","Sauvegarde informatique","Gestion des mots de passe","Gestion des visiteurs","Produits non conformes","Surveillance CCP","Gestion des incidents SSI"] },
  ],
  formulaire: [
    { group: "Catalogue OCG", items: ["Fiche de non-conformité","Fiche de réclamation client","Fiche d'évaluation des fournisseurs","Fiche de contrôle à la réception","Fiche de contrôle de l'efficacité du nettoyage","Fiche d'étalonnage","Fiche de présence à la formation","Fiche d'évaluation de la formation","Bon de commande","Bon de livraison","Bon de réception","Ordre de fabrication","Fiche de gestion de changement","Questionnaire de satisfaction client","Bon de sortie","Test de traçabilité"] },
    { group: "Formulaires", items: ["Fiche de non-conformité","Fiche d'action corrective","Fiche d'audit","Fiche de réclamation client","Fiche d'évaluation fournisseur","Fiche d'évaluation formation","Fiche d'analyse des risques","Fiche d'accident de travail","Fiche de contrôle qualité","Fiche de contrôle environnemental"] },
  ],
  enregistrement: [
    { group: "Catalogue OCG", items: ["Liste des informations documentées","Cartographie des processus","Politique qualité","Politique SMI","Politique SMSDA","Tableau des parties intéressées","Registre des risques et opportunités","Registre des réclamations client","Rapport d'audit interne","Compte rendu de revue de direction","Registre des non-conformités et actions correctives","Registre des retraits / rappels de produits","Plan de communication","Programme d'audit interne","Planning audit interne","Plan d'action","Tableau de bord KPI","Plan HACCP","Check-list des PRP","Plan de formation","Plan de maintenance"] },
    { group: "Enregistrements", items: ["Registre des risques","Registre des opportunités","Registre des non-conformités","Registre des actions correctives","Registre des audits internes","Registre des réclamations","Registre des accidents","Registre des déchets","Registre des formations","Registre des compétences","Registre des fournisseurs","Registre des équipements","Registre des CCP","Registre de traçabilité","Registre des incidents SSI","Inventaire des actifs informationnels","Registre des sauvegardes","Registre des accès utilisateurs"] },
  ],
};

const IGNORE = new Set(["DE","DES","ET","LA","LE","LES","DU","D","AU","AUX","UN","UNE","EN","PAR","SUR","POUR","AVEC","DANS","L","A","SE","SA","SES","OU","QUE","QUI","SOUS","VERS","ENTRE","DONT","SON","CAS","PROCESSUS","PROCEDURE","PROCEDURES","ENREGISTREMENT","ENREGISTREMENTS","FORMULAIRE","FORMULAIRES","FICHE","FICHES","LOGIGRAMME","LOGIGRAMMES","INSTRUCTION","INSTRUCTIONS","MODE","OPERATOIRE","OPERATOIRES"]);

function autoGenerateCode(name) {
  if (!name.trim()) return "";
  const cleaned = name.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^A-Z\s]/g,"").trim();
  const words = cleaned.split(/\s+/).filter(w=>w.length>0);
  if (words.length===0) return "";
  if (words.length>1) return words.map(w=>w[0]).join("").slice(0,3);
  const w=words[0];
  if(w.length<=4)return w.slice(0,1);
  if(w.length<=7)return w.slice(0,2);
  return w.slice(0,3);
}

function makeAbbrev(intitule) {
  const words = intitule.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^A-Z\s]/g," ").split(/\s+/).filter(w=>w&&!IGNORE.has(w));
  let abbrev = words.map(w=>w[0]).join("").slice(0,5);
  if(abbrev.length===1&&words[0])abbrev=words[0].slice(0,3);
  return abbrev||"DOC";
}

function previewCode(typeKey, customCode, intitule, counters) {
  if(!typeKey||!intitule.trim())return "_ _ _";
  let typeCode;
  if(typeKey==="__autre__"){typeCode=customCode.trim().toUpperCase().slice(0,3)||"??";}
  else{typeCode=TYPE_MAP_BACKEND[typeKey];}
  if(!typeCode)return "???";
  const abbrev=makeAbbrev(intitule);
  const next=String((counters[typeCode]||0)+1).padStart(2,"0");
  return `${typeCode}-${abbrev}-${next}`;
}

function orderedTypeCodes(codes) {
  const extras=[];
  for(const c of codes){if(!TYPE_ORDER.includes(c.type)&&!extras.includes(c.type))extras.push(c.type);}
  return [...TYPE_ORDER,...extras];
}
function pluralLabel(typeCode,group){return TYPE_LABELS_PLURAL[typeCode]||(group[0]&&group[0].typeLabel)||typeCode;}

/* =========================================================================
   EXPORT EXCEL — style aligné sur le modèle MID (ENG-LID-01)
   ========================================================================= */
const XFONT = "Times New Roman";
const HEADER_FILL = "FFED7D31";   // orange (en-tête de tableau)
const HEADER_FONT = "FFFFFFFF";   // texte blanc sur l'en-tête

// Une couleur claire par famille de type, pour repérer visuellement les groupes
const GROUP_FILLS = {
  PS:  "FFFCE4D6", // orange clair
  PR:  "FFDDEBF7", // bleu clair
  ENG: "FFE2EFDA", // vert clair
  F:   "FFFFF2CC", // jaune clair
  L:   "FFEAD1DC", // rose clair
  INS: "FFD9D2E9", // violet clair
  MO:  "FFD0E0E3", // bleu-vert clair
};
const groupFill = (t) => GROUP_FILLS[t] || "FFF2F2F2";

async function exportExcel(codes, branding={}) {
  const company=(branding.company||"OCG CONSULTING").trim();
  const wb=new ExcelJS.Workbook();wb.creator=`${company} — DocCode Generator`;
  const ws=wb.addWorksheet("Informations documentées",{views:[{state:"frozen",ySplit:8}],pageSetup:{orientation:"portrait",fitToPage:true,fitToWidth:1,fitToHeight:0}});

  // Largeurs alignées sur le modèle MID
  ws.columns=[{width:20.7},{width:77.9},{width:18.7},{width:11.6},{width:11.3},{width:20.4}];
  const thin={style:"thin",color:{argb:"FFBFBFBF"}};const border={top:thin,left:thin,bottom:thin,right:thin};

  // Bloc logo (A1:A3), sans fond coloré — sobre comme le modèle
  ws.mergeCells("A1:A3");const logo=ws.getCell("A1");
  ws.getRow(1).height=15;ws.getRow(2).height=15;ws.getRow(3).height=15.75;
  if(branding.logo&&dataUrlToBase64(branding.logo)){
    const imgId=wb.addImage({base64:branding.logo,extension:"png"});
    const{width,height}=scaleDims(branding.logoW,branding.logoH,140,50);
    ws.addImage(imgId,{tl:{col:0.15,row:0.15},ext:{width,height},editAs:"oneCell"});
  }else{
    logo.value=company.toUpperCase();logo.font={name:XFONT,bold:true,size:11};
    logo.alignment={horizontal:"center",vertical:"middle",wrapText:true};
  }

  // Titre (B1:C3)
  ws.mergeCells("B1:C3");const titre=ws.getCell("B1");
  titre.value="Liste des Informations documentées";
  titre.font={name:XFONT,bold:true,size:18};
  titre.alignment={horizontal:"center",vertical:"middle",wrapText:true};

  // Bloc Code / Version / Date en haut à droite
  ws.mergeCells("D1:F2");
  ws.getCell("D1").value=`Code : ENG-LID-01`;
  ws.getCell("D1").font={name:XFONT,size:12};
  ws.getCell("D1").alignment={horizontal:"left",vertical:"center",wrapText:true};
  ws.mergeCells("D3:F3");
  ws.getCell("D3").value=`Version : 00   |   Date : ${new Date().toLocaleDateString("fr-FR")}`;
  ws.getCell("D3").font={name:XFONT,size:12};
  ws.getCell("D3").alignment={horizontal:"left",vertical:"center"};

  for(let r0=1;r0<=3;r0++)for(let c0=1;c0<=6;c0++)ws.getCell(r0,c0).border=border;

  // En-tête de tableau à la ligne 8 (comme le modèle)
  const HR=8;
  const headers=["Type de document","Intitulé du document","Code","Version","Date","Motif de mise à jour"];
  const headerRow=ws.getRow(HR);
  headers.forEach((h,i)=>{
    const cell=headerRow.getCell(i+1);
    cell.value=h;
    cell.font={name:XFONT,bold:true,size:11,color:{argb:HEADER_FONT}};
    cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:HEADER_FILL}};
    cell.alignment={horizontal:"center",vertical:"middle",wrapText:true};
    cell.border=border;
  });
  headerRow.height=22;

  let r=HR+1;
  orderedTypeCodes(codes).forEach(t=>{
    const group=codes.filter(c=>c.type===t).slice().sort((a,b)=>(a.seq||0)-(b.seq||0));
    if(group.length===0)return;const start=r;
    const fill=groupFill(t);
    group.forEach(item=>{
      const row=ws.getRow(r);
      row.getCell(1).value=pluralLabel(t,group);
      row.getCell(2).value=item.intitule||"";
      row.getCell(3).value=item.code||"";
      row.getCell(4).value=item.version??0;
      row.getCell(5).value=item.createdAt?new Date(item.createdAt):new Date();
      row.getCell(6).value=item.motif||"Création";
      for(let col=1;col<=6;col++){
        const cell=row.getCell(col);
        cell.border=border;
        cell.font={name:XFONT,size:11};
        cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:fill}};
        cell.alignment={horizontal:"left",vertical:"middle",wrapText:true};
        if(col===1){cell.font={name:XFONT,size:12,bold:true};cell.alignment={horizontal:"center",vertical:"middle",wrapText:true};}
        else if(col===3){cell.font={name:XFONT,size:11,bold:true};cell.alignment={horizontal:"left",vertical:"middle"};}
        else if(col===4||col===6){cell.alignment={horizontal:"center",vertical:"middle"};}
        else if(col===5){cell.numFmt="dd/mm/yyyy";cell.alignment={horizontal:"center",vertical:"middle"};}
      }
      r++;
    });
    if(r-1>start)ws.mergeCells(`A${start}:A${r-1}`);
  });

  const buffer=await wb.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),"liste-informations-documentees.xlsx");
}

/* =========================================================================
   EXPORT WORD CODIFICATION (inchangé)
   ========================================================================= */
const NAVY_HEX="1F3864";const LIGHT_HEX="D9E1F2";

function wcell(text,opts={}){
  const{bold=false,color="000000",fill=null,align=AlignmentType.LEFT,size=20,width=null,mono=false,rowSpan=null}=opts;
  return new TableCell({width:width?{size:width,type:WidthType.PERCENTAGE}:undefined,verticalAlign:VerticalAlign.CENTER,rowSpan:rowSpan||undefined,shading:fill?{type:ShadingType.SOLID,color:fill,fill}:undefined,children:[new Paragraph({alignment:align,children:[new TextRun({text:String(text??""),bold,size,color,font:mono?"Consolas":"Times New Roman"})]})]});
}

async function exportWord(codes, branding={}) {
  const company=(branding.company||"OCG CONSULTING").trim();
  const rows=[];
  rows.push(new TableRow({tableHeader:true,children:["Type de document","Intitulé du document","Code","Version","Date","Motif de mise à jour"].map((h,i)=>wcell(h,{bold:true,color:"FFFFFF",fill:NAVY_HEX,align:AlignmentType.CENTER,size:20,width:[16,34,16,8,12,14][i]}))}));
  orderedTypeCodes(codes).forEach(t=>{
    const group=codes.filter(c=>c.type===t).slice().sort((a,b)=>(a.seq||0)-(b.seq||0));
    if(group.length===0)return;
    group.forEach((item,idx)=>{
      const cells=[];
      if(idx===0)cells.push(wcell(pluralLabel(t,group),{bold:true,color:NAVY_HEX,fill:LIGHT_HEX,align:AlignmentType.CENTER,rowSpan:group.length,width:16}));
      cells.push(wcell(item.intitule,{width:34}));cells.push(wcell(item.code,{bold:true,color:NAVY_HEX,mono:true,align:AlignmentType.CENTER,width:16}));
      cells.push(wcell(item.version??0,{align:AlignmentType.CENTER,width:8}));cells.push(wcell(new Date(item.createdAt||Date.now()).toLocaleDateString("fr-FR"),{align:AlignmentType.CENTER,width:12}));
      cells.push(wcell(item.motif||"Création",{align:AlignmentType.CENTER,width:14}));
      rows.push(new TableRow({children:cells}));
    });
  });
  const tableBorder={style:BorderStyle.SINGLE,size:4,color:"BFBFBF"};
  const headerChildren=[];
  const b64=dataUrlToBase64(branding.logo);
  if(b64){const{width,height}=scaleDims(branding.logoW,branding.logoH,150,75);headerChildren.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:100},children:[new ImageRun({type:"png",data:base64ToUint8(b64),transformation:{width,height}})]}));}
  headerChildren.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:company,bold:true,size:26,color:NAVY_HEX,font:"Times New Roman"})]}));
  headerChildren.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Liste des Informations Documentées",bold:true,size:32,color:NAVY_HEX,font:"Times New Roman"})]}));
  headerChildren.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:"Code : ENG-LID-01   |   Version : 00   |   Date : "+new Date().toLocaleDateString("fr-FR"),size:18,color:"555555",font:"Times New Roman"})]}));
  headerChildren.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},borders:{top:tableBorder,bottom:tableBorder,left:tableBorder,right:tableBorder,insideHorizontal:tableBorder,insideVertical:tableBorder},rows}));
  const doc=new Document({sections:[{properties:{page:{size:{orientation:"landscape"}}},children:headerChildren}]});
  downloadBlob(await Packer.toBlob(doc),"liste-informations-documentees.docx");
}

/* =========================================================================
   ÉCRAN DE CONNEXION / CRÉATION DE COMPTE (inchangé)
   ========================================================================= */
const IcUser=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>;
const IcMail=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 6-10 7L2 6"/></svg>;
const IcLock=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>;
const IcBuild=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 21h18"/><path d="M5 21V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v16"/><path d="M9 9h.01M9 13h.01M9 17h.01M15 9h.01M15 13h.01M15 17h.01"/></svg>;

function AuthScreen({ onAuthenticated }) {
  const[mode,setMode]=useState("login");
  const[name,setName]=useState("");const[company,setCompany]=useState("");const[email,setEmail]=useState("");const[pw,setPw]=useState("");const[pw2,setPw2]=useState("");const[error,setError]=useState("");const[logo,setLogo]=useState(null);const[logoW,setLogoW]=useState(null);const[logoH,setLogoH]=useState(null);
  const reset=()=>{setError("");setPw("");setPw2("");};
  const goLogin=()=>{setMode("login");reset();};const goSignup=()=>{setMode("signup");reset();};
  const handleLogoPick=async(e)=>{const file=e.target.files&&e.target.files[0];if(!file)return;try{const{dataUrl,width,height}=await readImageFile(file,256);setLogo(dataUrl);setLogoW(width);setLogoH(height);setError("");}catch(err){setError("Image illisible.");}e.target.value="";};
  const clearLogo=()=>{setLogo(null);setLogoW(null);setLogoH(null);};
  const doLogin=()=>{setError("");const key=email.trim().toLowerCase();if(!key||!pw){setError("Veuillez saisir votre email et mot de passe.");return;}const accounts=loadAccounts();const acc=accounts[key];if(!acc){setError("Aucun compte avec cet email.");return;}if(acc.passHash!==hashPassword(pw)){setError("Mot de passe incorrect.");return;}onAuthenticated(acc.email);};
  const doSignup=()=>{setError("");if(!name.trim()||!company.trim()||!email.trim()||!pw){setError("Tous les champs sont obligatoires.");return;}if(!/^\S+@\S+\.\S+$/.test(email.trim())){setError("Email invalide.");return;}if(pw.length<4){setError("Mot de passe trop court (4 car. min).");return;}if(pw!==pw2){setError("Mots de passe différents.");return;}const key=email.trim().toLowerCase();const accounts=loadAccounts();if(accounts[key]){setError("Compte déjà existant.");return;}accounts[key]={email:key,name:name.trim(),company:company.trim(),passHash:hashPassword(pw),createdAt:new Date().toISOString(),logo:logo||null,logoW:logoW||null,logoH:logoH||null};saveAccounts(accounts);onAuthenticated(key);};

  return(
    <div className="auth-wrap">
      <style>{`
        .auth-wrap*{box-sizing:border-box;}
        .auth-wrap{position:relative;min-height:100vh;width:100%;display:flex;align-items:center;justify-content:center;background:#EDEFF0;font-family:'Syne',sans-serif;padding:24px;overflow:hidden;}
        .bg-shape{position:absolute;z-index:0;}
        .bg-red{top:-70px;right:-70px;width:280px;height:280px;background:linear-gradient(135deg,#F26D7D,#E8576B);border-radius:44px;transform:rotate(45deg);box-shadow:0 20px 50px rgba(232,87,107,.25);}
        .bg-yellow{bottom:-120px;left:-120px;width:330px;height:330px;background:linear-gradient(135deg,#FFD45E,#F6C544);border-radius:50%;}
        .auth-container{position:relative;z-index:1;width:880px;max-width:100%;min-height:600px;background:#FFFFFF;border-radius:20px;overflow:hidden;box-shadow:0 30px 80px rgba(15,90,70,.20);}
        .form-container{position:absolute;top:0;height:100%;width:50%;transition:all .6s ease-in-out;overflow-y:auto;}
        .sign-in-container{left:0;z-index:2;}.sign-up-container{left:0;opacity:0;z-index:1;}
        .auth-container.right-panel-active .sign-in-container{transform:translateX(100%);opacity:0;}
        .auth-container.right-panel-active .sign-up-container{transform:translateX(100%);opacity:1;z-index:5;animation:ac-show .6s;}
        @keyframes ac-show{0%,49.99%{opacity:0;z-index:1;}50%,100%{opacity:1;z-index:5;}}
        .auth-form{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;padding:34px 48px;text-align:center;}
        .brand-mini{font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#0F6E56;font-weight:800;margin-bottom:12px;display:flex;align-items:center;gap:6px;}.brand-glyph{color:#1AAE86;}
        .auth-title{font-size:27px;font-weight:800;color:#0F6E56;margin:0;letter-spacing:-.6px;}.auth-sub{font-size:12px;color:#9AA0A6;margin:6px 0 16px;}
        .logo-drop{width:92px;height:92px;border-radius:50%;border:2px dashed #CBD0CC;background:#F5F6F5;display:flex;align-items:center;justify-content:center;cursor:pointer;overflow:hidden;flex-shrink:0;transition:border-color .2s,transform .15s;}
        .logo-drop:hover{transform:scale(1.04);border-color:#1AAE86;}.logo-drop.has-logo{border-style:solid;border-color:#1AAE86;background:#FFFFFF;padding:6px;}
        .logo-drop img{width:100%;height:100%;object-fit:contain;}.logo-drop-empty{font-size:11px;font-weight:800;letter-spacing:1px;color:#B4BAB5;line-height:1.3;}
        .logo-cap{font-size:11px;color:#9AA0A6;margin:8px 0 6px;}.logo-link{background:none;border:none;color:#0F6E56;font-weight:700;font-size:12px;cursor:pointer;margin:8px 0 4px;font-family:'Syne',sans-serif;}
        .in-row{display:flex;align-items:center;gap:10px;width:100%;background:#EFF0F0;border:1px solid #EFF0F0;border-radius:10px;padding:0 14px;margin:6px 0;transition:border-color .2s,background .2s;}
        .in-row:focus-within{border-color:#1AAE86;background:#FFFFFF;}.in-row svg{color:#A2A8A4;flex-shrink:0;}
        .in-row input{flex:1;border:none;background:transparent;outline:none;padding:13px 0;font-size:14px;color:#2C2C2A;font-family:'Syne',sans-serif;}
        .auth-error{width:100%;background:#FCEBEB;border:1px solid #F1C7C7;color:#C0392B;border-radius:9px;padding:9px 12px;font-size:12px;margin:6px 0 4px;}
        .btn-solid{margin-top:18px;border:1px solid #0F6E56;background:linear-gradient(135deg,#1AAE86,#0F6E56);color:#fff;border-radius:24px;padding:13px 52px;cursor:pointer;font-size:12px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;font-family:'Syne',sans-serif;transition:transform .1s,box-shadow .2s;box-shadow:0 8px 20px rgba(15,110,86,.25);}
        .btn-solid:hover{box-shadow:0 10px 26px rgba(15,110,86,.32);}.btn-solid:active{transform:scale(.96);}
        .btn-ghost{margin-top:22px;border:1.5px solid rgba(255,255,255,.9);background:transparent;color:#fff;border-radius:24px;padding:12px 46px;cursor:pointer;font-size:12px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;font-family:'Syne',sans-serif;transition:background .2s,transform .1s;}
        .btn-ghost:hover{background:rgba(255,255,255,.14);}.btn-ghost:active{transform:scale(.96);}
        .overlay-container{position:absolute;top:0;left:50%;width:50%;height:100%;overflow:hidden;transition:transform .6s ease-in-out;z-index:100;}
        .auth-container.right-panel-active .overlay-container{transform:translateX(-100%);}
        .overlay{position:relative;left:-100%;height:100%;width:200%;background:linear-gradient(135deg,#1AAE86 0%,#0E6A52 100%);color:#fff;transform:translateX(0);transition:transform .6s ease-in-out;}
        .auth-container.right-panel-active .overlay{transform:translateX(50%);}
        .overlay-panel{position:absolute;top:0;height:100%;width:50%;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:0 44px;}
        .overlay-left{left:0;transform:translateX(-18%);}.auth-container.right-panel-active .overlay-left{transform:translateX(0);}
        .overlay-right{right:0;transform:translateX(0);}.auth-container.right-panel-active .overlay-right{transform:translateX(18%);}
        .overlay h1{font-size:30px;font-weight:800;margin:14px 0 0;letter-spacing:-.5px;}.overlay p{font-size:13px;line-height:1.6;margin:16px 0 0;opacity:.92;max-width:260px;}
        .ov-logo{width:90px;height:90px;border-radius:16px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.3);display:flex;align-items:center;justify-content:center;font-size:26px;font-weight:800;}
        .ov-foot{position:absolute;bottom:26px;left:0;right:0;font-size:10px;letter-spacing:1px;opacity:.7;font-family:'IBM Plex Mono',monospace;padding:0 30px;}
        .ov-diamond{position:absolute;background:rgba(255,255,255,.08);transform:rotate(45deg);border-radius:8px;z-index:0;}
        .ov-diamond.d1{width:66px;height:66px;top:14%;left:30%;}.ov-diamond.d2{width:32px;height:32px;bottom:20%;left:22%;}.ov-diamond.d3{width:46px;height:46px;top:24%;right:26%;}
        .mobile-switch{display:none;margin-top:18px;font-size:12px;color:#9AA0A6;}.mobile-switch span{color:#0F6E56;font-weight:800;cursor:pointer;}
        @media(max-width:780px){.auth-container{min-height:auto;}.overlay-container{display:none;}.form-container{position:relative;width:100%;height:auto;transform:none!important;}.sign-in-container{opacity:1;}.sign-up-container{display:none;}.auth-container.right-panel-active .sign-in-container{display:none;}.auth-container.right-panel-active .sign-up-container{display:block;opacity:1;transform:none!important;animation:none;}.mobile-switch{display:block;}.auth-form{padding:36px 26px;}}
      `}</style>
      <span className="bg-shape bg-red"/><span className="bg-shape bg-yellow"/>
      <div className={"auth-container"+(mode==="signup"?" right-panel-active":"")}>
        <div className="form-container sign-up-container">
          <div className="auth-form">
            <div className="brand-mini"><span className="brand-glyph">◆</span> CoDoc Generator</div>
            <h2 className="auth-title">Créer un compte</h2><p className="auth-sub">Votre espace de codification documentaire</p>
            <label className={"logo-drop"+(logo?" has-logo":"")} title="Logo entreprise">{logo?<img src={logo} alt="logo"/>:<span className="logo-drop-empty">＋<br/>LOGO</span>}<input type="file" accept="image/*" style={{display:"none"}} onChange={handleLogoPick}/></label>
            {logo?<button type="button" className="logo-link" onClick={clearLogo}>Retirer le logo</button>:<div className="logo-cap">Logo de l'entreprise (optionnel)</div>}
            {error&&mode==="signup"&&<div className="auth-error">{error}</div>}
            <div className="in-row"><IcUser/><input placeholder="Nom complet" value={name} onChange={e=>setName(e.target.value)}/></div>
            <div className="in-row"><IcBuild/><input placeholder="Entreprise" value={company} onChange={e=>setCompany(e.target.value)}/></div>
            <div className="in-row"><IcMail/><input type="email" placeholder="Email" value={email} onChange={e=>setEmail(e.target.value)}/></div>
            <div className="in-row"><IcLock/><input type="password" placeholder="Mot de passe" value={pw} onChange={e=>setPw(e.target.value)}/></div>
            <div className="in-row"><IcLock/><input type="password" placeholder="Confirmer le mot de passe" value={pw2} onChange={e=>setPw2(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doSignup()}/></div>
            <button className="btn-solid" onClick={doSignup}>S'inscrire</button>
            <div className="mobile-switch">Déjà inscrit ? <span onClick={goLogin}>Se connecter</span></div>
          </div>
        </div>
        <div className="form-container sign-in-container">
          <div className="auth-form">
            <div className="brand-mini"><span className="brand-glyph">◆</span> CoDoc Generator</div>
            <h2 className="auth-title">Connexion</h2><p className="auth-sub">Heureux de vous revoir</p>
            {error&&mode==="login"&&<div className="auth-error">{error}</div>}
            <div className="in-row"><IcMail/><input type="email" placeholder="Email" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doLogin()}/></div>
            <div className="in-row"><IcLock/><input type="password" placeholder="Mot de passe" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doLogin()}/></div>
            <button className="btn-solid" onClick={doLogin}>Se connecter</button>
            <div className="mobile-switch">Pas de compte ? <span onClick={goSignup}>Créer un compte</span></div>
          </div>
        </div>
        <div className="overlay-container"><div className="overlay">
          <span className="ov-diamond d1"/><span className="ov-diamond d2"/><span className="ov-diamond d3"/>
          <div className="overlay-panel overlay-left">
            <div className="ov-logo" style={{background:"white",padding:"6px"}}><img src={logoOcg} alt="logo" style={{width:"100%",height:"100%",objectFit:"contain"}}/></div>
            <h1>Déjà membre ?</h1><p>Connectez-vous pour retrouver votre tableau de codes et votre historique.</p>
            <button className="btn-ghost" onClick={goLogin}>Se connecter</button>
            <div className="ov-foot">OCG CONSULTING — ISO 9001 · 14001 · 45001 · 27001</div>
          </div>
          <div className="overlay-panel overlay-right">
            <div className="ov-logo" style={{background:"white",padding:"6px"}}><img src={logoOcg} alt="logo" style={{width:"100%",height:"100%",objectFit:"contain"}}/></div>
            <h1>Bonjour</h1><p>Créez votre espace pour générer et archiver vos codes documentaires QMS / ISO.</p>
            <button className="btn-ghost" onClick={goSignup}>Créer un compte</button>
            <div className="ov-foot">OCG CONSULTING — ISO 9001 · 14001 · 45001 · 27001</div>
          </div>
        </div></div>
      </div>
    </div>
  );
}

/* =========================================================================
   ÉCRAN DE SÉLECTION DU MODULE — style "Select User Type"
   Cartes sobres · icônes SVG en trait · sans description · couleurs OCG
   ========================================================================= */
function ModuleSelector({ user, onSelect, onLogout }) {
  const [selected, setSelected] = useState(null);
  const initials = (user.company || user.name || "?")
    .trim().split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase();

  const MODULES = [
    {
      id: "codification",
      label: "Codification",
      color: "#0F6E56",
      grad: "linear-gradient(135deg,#1AAE86,#0F6E56)",
      glow: "rgba(15,110,86,.42)",
      icon: (active) => (
        <svg width="36" height="36" viewBox="0 0 24 24" fill="none"
          stroke={active ? "#fff" : "#0F6E56"} strokeWidth="2"
          strokeLinecap="round" strokeLinejoin="round">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
          <path d="M14 2v6h6" /><path d="M9 13h6M9 17h4" />
        </svg>
      ),
    },
    {
      id: "cartographie",
      label: "Cartographie",
      color: "#E8703A",
      grad: "linear-gradient(135deg,#F07040,#C0501A)",
      glow: "rgba(200,80,26,.42)",
      icon: (active) => (
        <svg width="36" height="36" viewBox="0 0 24 24" fill="none"
          stroke={active ? "#fff" : "#E8703A"} strokeWidth="2"
          strokeLinecap="round" strokeLinejoin="round">
          <rect x="9" y="3" width="6" height="5" rx="1" />
          <rect x="3" y="16" width="6" height="5" rx="1" />
          <rect x="15" y="16" width="6" height="5" rx="1" />
          <path d="M12 8v4M6 16v-2h12v2" />
        </svg>
      ),
    },
  ];

  const proceed = () => { if (selected) onSelect(selected); };
  const arrowColor = selected
    ? MODULES.find(m => m.id === selected).color
    : "#C7CDC9";

  return (
    <div className="ms-wrap">
      <style>{`
        .ms-wrap, .ms-wrap *{ box-sizing:border-box; }
        .ms-wrap{
          min-height:100vh; width:100%; display:flex; align-items:center;
          justify-content:center; padding:28px;
          font-family:'Syne',sans-serif;
          background:#E7ECEA;
        }
        .ms-panel{
          position:relative; width:760px; max-width:100%;
          background:#FFFFFF; border-radius:22px; padding:38px 46px 34px;
          box-shadow:0 30px 70px rgba(20,40,34,.16);
        }
        /* En-tête compte (discret, en haut à droite) */
        .ms-user{
          position:absolute; top:26px; right:30px;
          display:flex; align-items:center; gap:10px;
        }
        .ms-user-meta{ text-align:right; line-height:1.2; }
        .ms-user-co{ font-size:13px; font-weight:800; color:#15201B; }
        .ms-user-mail{ font-size:10.5px; color:#9AA0A6; font-family:'IBM Plex Mono',monospace; }
        .ms-avatar{
          width:38px; height:38px; border-radius:50%; overflow:hidden; flex-shrink:0;
          display:flex; align-items:center; justify-content:center;
          background:#0F6E56; color:#fff; font-weight:800; font-size:13px;
          border:1px solid #E2E5E2;
        }
        .ms-avatar img{ width:100%; height:100%; object-fit:contain; background:#fff; }
        /* Titre + accent */
        .ms-title{
          font-size:27px; font-weight:800; letter-spacing:.5px;
          text-transform:uppercase; color:#1B2B25; margin:6px 0 0;
        }
        .ms-accent{ width:54px; height:5px; border-radius:3px; background:#0F6E56; margin:14px 0 30px; }
        /* Cartes */
        .ms-cards{ display:flex; gap:22px; flex-wrap:wrap; }
        .ms-card{
          flex:1; min-width:150px; background:#FCFDFC;
          border:1px solid #EDF0EE; border-radius:18px;
          padding:30px 18px 24px; cursor:pointer; text-align:center;
          transition:transform .18s ease, box-shadow .2s ease, background .2s ease;
          box-shadow:0 6px 18px rgba(20,40,34,.06);
        }
        .ms-card:hover{ transform:translateY(-5px); box-shadow:0 16px 34px rgba(20,40,34,.12); }
        .ms-card.active{ transform:translateY(-5px); border-color:transparent; }
        .ms-ic{
          width:74px; height:74px; border-radius:18px; margin:0 auto 18px;
          display:flex; align-items:center; justify-content:center;
          background:#F2F5F3; transition:background .2s ease;
        }
        .ms-card.active .ms-ic{ background:rgba(255,255,255,.18); }
        .ms-label{ font-size:15px; font-weight:800; letter-spacing:-.2px; }
        /* Barre du bas */
        .ms-foot{ display:flex; align-items:center; justify-content:center; gap:16px; margin-top:34px; }
        .ms-back{
          background:#F1F3F2; border:1px solid #E6E9E7; color:#6C746F;
          border-radius:26px; padding:13px 30px; font-family:'Syne',sans-serif;
          font-size:13px; font-weight:800; cursor:pointer; transition:background .15s;
        }
        .ms-back:hover{ background:#E8EBE9; }
        .ms-next{
          width:48px; height:48px; border-radius:50%; border:none;
          display:flex; align-items:center; justify-content:center;
          transition:transform .12s ease, box-shadow .2s ease, background .2s ease;
        }
        .ms-next:enabled{ cursor:pointer; }
        .ms-next:enabled:hover{ transform:scale(1.06); }
        .ms-next:disabled{ cursor:not-allowed; opacity:.85; }
        .ms-hint{
          text-align:center; margin-top:18px; font-size:11px; letter-spacing:1px;
          text-transform:uppercase; color:#A8AEAA; font-family:'IBM Plex Mono',monospace;
        }
        .ms-foot-brand{
          text-align:center; margin-top:26px; font-size:10px; letter-spacing:1px;
          color:#B7BDB9; font-family:'IBM Plex Mono',monospace;
        }
        @media(max-width:560px){
          .ms-panel{ padding:30px 24px 26px; }
          .ms-user{ position:static; justify-content:flex-end; margin-bottom:18px; }
          .ms-cards{ flex-direction:column; }
        }
      `}</style>

      <div className="ms-panel">
        {/* Compte connecté */}
        <div className="ms-user">
          <div className="ms-user-meta">
            <div className="ms-user-co">{user.company}</div>
            <div className="ms-user-mail">{user.email}</div>
          </div>
          <div className="ms-avatar">
            {user.logo
              ? <img src={user.logo} alt="logo" />
              : initials}
          </div>
        </div>

        {/* Titre */}
        <h2 className="ms-title">Choisir un programme</h2>
        <div className="ms-accent" />

        {/* Cartes */}
        <div className="ms-cards">
          {MODULES.map(m => {
            const active = selected === m.id;
            return (
              <div
                key={m.id}
                className={"ms-card" + (active ? " active" : "")}
                onClick={() => setSelected(m.id)}
                onDoubleClick={() => onSelect(m.id)}
                style={active ? {
                  background: m.grad,
                  color: "#fff",
                  boxShadow: `0 22px 44px ${m.glow}`,
                } : {}}
              >
                <div className="ms-ic">{m.icon(active)}</div>
                <div className="ms-label" style={{ color: active ? "#fff" : m.color }}>
                  {m.label}
                </div>
              </div>
            );
          })}
        </div>

        {/* Bas : Déconnexion + flèche valider */}
        <div className="ms-foot">
          <button className="ms-back" onClick={onLogout}>Déconnexion</button>
          <button
            className="ms-next"
            onClick={proceed}
            disabled={!selected}
            style={{
              background: arrowColor,
              boxShadow: selected ? `0 10px 22px ${MODULES.find(m => m.id === selected).glow}` : "none",
            }}
            title={selected ? "Continuer" : "Sélectionnez un programme"}
          >
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none"
              stroke="#fff" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
              <path d="M5 12h14M13 6l6 6-6 6" />
            </svg>
          </button>
        </div>

        <div className="ms-hint">
          {selected ? "Cliquez sur la flèche pour continuer" : "Sélectionnez un programme"}
        </div>

        <div className="ms-foot-brand">
          OCG CONSULTING — ISO 9001 · 14001 · 45001 · 27001
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   MODULE CARTOGRAPHIE
   ========================================================================= */
function SearchableList({ items, selected, onToggle, onAddCustom, placeholder = "Rechercher…", accent = "#E8703A" }) {
  const [q, setQ] = useState("");
  const [custom, setCustom] = useState("");
  const norm = s => s.normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();
  const filtered = items.filter(it => !q.trim() || norm(it).includes(norm(q.trim())));
  const addCustom = () => {
    const v = custom.trim();
    if (!v) return;
    if (onAddCustom) onAddCustom(v);
    setCustom("");
  };
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:"8px",background:"#F6F7F6",border:"1px solid #E2E5E2",borderRadius:"9px",padding:"0 12px",marginBottom:"8px"}}>
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#A2A8A4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg>
        <input style={{flex:1,border:"none",background:"transparent",outline:"none",padding:"10px 0",fontSize:"13px",fontFamily:"'Syne',sans-serif",color:"#1F2421"}} placeholder={placeholder} value={q} onChange={e=>setQ(e.target.value)}/>
      </div>
      <div style={{maxHeight:"200px",overflowY:"auto",border:"1px solid #E2E5E2",borderRadius:"9px",background:"#fff"}}>
        {filtered.length===0?(
          <div style={{padding:"16px",textAlign:"center",fontSize:"12px",color:"#A2A8A4",fontFamily:"'IBM Plex Mono',monospace"}}>Aucun résultat</div>
        ):filtered.map(item=>{
          const active=selected.includes(item);
          return(
            <div key={item} onClick={()=>onToggle(item)} style={{display:"flex",alignItems:"center",gap:"10px",padding:"9px 12px",cursor:"pointer",background:active?`${accent}14`:"transparent",borderBottom:"1px solid #F2F4F2",transition:"background .1s"}}>
              <div style={{width:"18px",height:"18px",borderRadius:"4px",border:active?`2px solid ${accent}`:"2px solid #D3D7D4",background:active?accent:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                {active&&<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><path d="M20 6 9 17l-5-5"/></svg>}
              </div>
              <span style={{fontSize:"13px",color:active?accent:"#39403B",fontWeight:active?700:400}}>{item}</span>
            </div>
          );
        })}
      </div>

      {/* Champ Autre — saisie d'un processus personnalisé */}
      {onAddCustom && (
        <div style={{display:"flex",alignItems:"center",gap:"8px",marginTop:"8px"}}>
          <div style={{flex:1,display:"flex",alignItems:"center",gap:"8px",background:"#fff",border:`1px dashed ${accent}66`,borderRadius:"9px",padding:"0 12px"}}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={accent} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 5v14M5 12h14"/></svg>
            <input
              style={{flex:1,border:"none",background:"transparent",outline:"none",padding:"10px 0",fontSize:"13px",fontFamily:"'Syne',sans-serif",color:"#1F2421"}}
              placeholder="Autre processus (saisie libre)…"
              value={custom}
              onChange={e=>setCustom(e.target.value)}
              onKeyDown={e=>{ if(e.key==="Enter"){ e.preventDefault(); addCustom(); } }}
            />
          </div>
          <button
            onClick={addCustom}
            disabled={!custom.trim()}
            style={{border:"none",borderRadius:"9px",background:custom.trim()?accent:"#D3D7D4",color:"#fff",padding:"10px 16px",fontSize:"13px",fontWeight:"800",fontFamily:"'Syne',sans-serif",cursor:custom.trim()?"pointer":"not-allowed",whiteSpace:"nowrap"}}
          >
            Ajouter
          </button>
        </div>
      )}
    </div>
  );
}

function SelectedTags({ items, onRemove, color="#E8703A" }) {
  if (items.length===0) return <div style={{fontSize:"12px",color:"#A2A8A4",fontStyle:"italic",padding:"8px 0"}}>Aucun processus sélectionné</div>;
  return(
    <div style={{display:"flex",flexWrap:"wrap",gap:"6px",marginTop:"8px"}}>
      {items.map((it,i)=>(
        <div key={it} style={{display:"inline-flex",alignItems:"center",gap:"6px",background:`${color}18`,border:`1px solid ${color}40`,borderRadius:"20px",padding:"4px 10px 4px 10px",fontSize:"12px",color,fontWeight:700}}>
          <span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:"10px",color:`${color}99`,marginRight:"2px"}}>#{i+1}</span>
          {it}
          <button onClick={()=>onRemove(it)} style={{background:"none",border:"none",color,cursor:"pointer",padding:"0",display:"flex",alignItems:"center",fontSize:"14px",lineHeight:1}}>×</button>
        </div>
      ))}
    </div>
  );
}

function CartoApp({ user, onBack, onLogout }) {
  const [carto, setCarto] = useState(() => loadUserCarto());
  const [toast, setToast] = useState(null);
  const showToast = msg => { setToast(msg); setTimeout(()=>setToast(null),3000); };

  // Pas de sauvegarde automatique — la cartographie repart vierge à chaque session

  const toggle = (family, item) => {
    setCarto(prev => {
      const arr = prev[family] || [];
      return { ...prev, [family]: arr.includes(item) ? arr.filter(x=>x!==item) : [...arr, item] };
    });
  };
  const remove = (family, item) => setCarto(prev=>({ ...prev, [family]: (prev[family]||[]).filter(x=>x!==item) }));

  const addCustom = (family, value) => {
    const v = (value || "").trim();
    if (!v) return;
    setCarto(prev => {
      const arr = prev[family] || [];
      if (arr.some(x => x.toLowerCase() === v.toLowerCase())) { showToast("⚠️ Ce processus existe déjà"); return prev; }
      return { ...prev, [family]: [...arr, v] };
    });
  };

  const moveUp = (family, i) => {
    if (i===0) return;
    const arr = [...(carto[family]||[])];
    [arr[i-1],arr[i]]=[arr[i],arr[i-1]];
    setCarto(prev=>({ ...prev, [family]: arr }));
  };
  const moveDown = (family, i) => {
    const arr = [...(carto[family]||[])];
    if (i>=arr.length-1) return;
    [arr[i],arr[i+1]]=[arr[i+1],arr[i]];
    setCarto(prev=>({ ...prev, [family]: arr }));
  };

  const handleExport = async () => {
    const total = (carto.management||[]).length + (carto.realisation||[]).length + (carto.support||[]).length;
    if (total===0) { showToast("⚠️ Ajoutez au moins un processus"); return; }
    try {
      await exportCartoWord(carto, { company: user.company, logo: user.logo, logoW: user.logoW, logoH: user.logoH });
      showToast("✅ Cartographie exportée en Word");
    } catch(e) { console.error(e); showToast("❌ Erreur export"); }
  };

  const initials=(user.company||user.name||"?").trim().split(/\s+/).map(w=>w[0]).slice(0,2).join("").toUpperCase();

  const ORANGE = "#E8703A";
  const sectionStyle = (color) => ({
    background:"#fff",borderRadius:"14px",border:"1px solid #E2E5E2",padding:"22px",marginBottom:"20px",
  });
  const sectionHead = (label, count, color) => (
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"14px"}}>
      <div style={{display:"flex",alignItems:"center",gap:"10px"}}>
        <div style={{width:"10px",height:"10px",borderRadius:"50%",background:color}}/>
        <span style={{fontWeight:"800",fontSize:"15px",color:"#15201B"}}>{label}</span>
      </div>
      <span style={{background:`${color}18`,color,border:`1px solid ${color}40`,borderRadius:"20px",padding:"3px 12px",fontSize:"11px",fontWeight:"800",fontFamily:"'IBM Plex Mono',monospace"}}>{count} processus</span>
    </div>
  );

  const OrderableList = ({ family, color }) => {
    const items = carto[family] || [];
    return items.length > 0 ? (
      <div style={{marginTop:"10px"}}>
        {items.map((it,i)=>(
          <div key={it} style={{display:"flex",alignItems:"center",gap:"8px",padding:"6px 10px",background:`${color}08`,border:`1px solid ${color}20`,borderRadius:"8px",marginBottom:"5px"}}>
            <span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:"10px",color:`${color}80`,width:"20px"}}>{i+1}.</span>
            <span style={{flex:1,fontSize:"13px",color:"#39403B"}}>{it}</span>
            <button onClick={()=>moveUp(family,i)} style={{background:"none",border:`1px solid ${color}40`,borderRadius:"5px",color,padding:"2px 6px",cursor:"pointer",fontSize:"11px"}} title="Monter">↑</button>
            <button onClick={()=>moveDown(family,i)} style={{background:"none",border:`1px solid ${color}40`,borderRadius:"5px",color,padding:"2px 6px",cursor:"pointer",fontSize:"11px"}} title="Descendre">↓</button>
            <button onClick={()=>remove(family,it)} style={{background:"none",border:`1px solid #F1C7C7`,borderRadius:"5px",color:"#C0392B",padding:"2px 6px",cursor:"pointer",fontSize:"11px"}}>×</button>
          </div>
        ))}
      </div>
    ) : <div style={{fontSize:"12px",color:"#A2A8A4",fontStyle:"italic",padding:"8px 0"}}>Aucun processus sélectionné</div>;
  };

  return (
    <div style={{minHeight:"100vh",background:"#F4F6F5",fontFamily:"'Syne',sans-serif",color:"#1F2421"}}>
      <style>{`*{box-sizing:border-box;}`}</style>

      {/* Topbar */}
      <div style={{background:"#fff",borderBottom:"1px solid #E7E9E7",padding:"14px 32px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"center",gap:"14px"}}>
          <button onClick={onBack} style={{background:"#F6F7F6",border:"1px solid #E2E5E2",borderRadius:"9px",padding:"9px 14px",cursor:"pointer",fontSize:"13px",fontWeight:"700",fontFamily:"'Syne',sans-serif",color:"#5C635E",display:"flex",alignItems:"center",gap:"6px"}}>
            ← Retour
          </button>
          <div>
            <div style={{fontWeight:"800",fontSize:"18px",color:"#15201B",letterSpacing:"-.3px"}}>Programme de Cartographie</div>
            <div style={{fontSize:"11px",color:"#9AA0A6",fontFamily:"'IBM Plex Mono',monospace",marginTop:"2px"}}>Composez et exportez votre cartographie des processus</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
          <div style={{textAlign:"right"}}>
            <div style={{fontWeight:"700",fontSize:"13px"}}>{user.name}</div>
            <div style={{fontSize:"11px",color:"#9AA0A6",fontFamily:"'IBM Plex Mono',monospace"}}>{user.company}</div>
          </div>
          <div style={{width:"40px",height:"40px",borderRadius:"50%",background:user.logo?"#fff":"#E8703A",border:"1px solid #E2E5E2",display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",color:"#fff",fontWeight:"800",fontSize:"13px"}}>
            {user.logo?<img src={user.logo} alt="logo" style={{width:"100%",height:"100%",objectFit:"contain"}}/>:initials}
          </div>
          <button onClick={onLogout} style={{background:"#FCEEEE",border:"1px solid #F1C7C7",borderRadius:"9px",color:"#C0392B",padding:"9px 14px",fontFamily:"'Syne',sans-serif",fontSize:"12px",fontWeight:"700",cursor:"pointer"}}>Déconnexion</button>
        </div>
      </div>

      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"28px 24px",display:"grid",gridTemplateColumns:"1fr 360px",gap:"24px",alignItems:"start"}}>

        {/* ---- Zone gauche : sélecteurs ---- */}
        <div>

          {/* MANAGEMENT */}
          <div style={sectionStyle(ORANGE)}>
            {sectionHead("Processus Management", (carto.management||[]).length, "#0F6E56")}
            <SearchableList items={PROCESSUS_CATALOG.management} selected={carto.management||[]} onToggle={it=>toggle("management",it)} onAddCustom={v=>addCustom("management",v)} accent="#0F6E56" placeholder="Rechercher un processus de management…"/>
            <div style={{marginTop:"12px",paddingTop:"12px",borderTop:"1px dashed #E2E5E2"}}>
              <div style={{fontSize:"11px",fontWeight:"700",color:"#0F6E56",textTransform:"uppercase",letterSpacing:"1px",marginBottom:"6px"}}>Ordre dans la cartographie</div>
              <OrderableList family="management" color="#0F6E56"/>
            </div>
          </div>

          {/* RÉALISATION */}
          <div style={sectionStyle(ORANGE)}>
            {sectionHead("Processus Réalisation", (carto.realisation||[]).length, ORANGE)}
            <SearchableList items={PROCESSUS_CATALOG.realisation} selected={carto.realisation||[]} onToggle={it=>toggle("realisation",it)} onAddCustom={v=>addCustom("realisation",v)} accent={ORANGE} placeholder="Rechercher un processus de réalisation…"/>
            <div style={{marginTop:"12px",paddingTop:"12px",borderTop:"1px dashed #E2E5E2"}}>
              <div style={{fontSize:"11px",fontWeight:"700",color:ORANGE,textTransform:"uppercase",letterSpacing:"1px",marginBottom:"6px"}}>Ordre dans la cartographie</div>
              <OrderableList family="realisation" color={ORANGE}/>
            </div>
          </div>

          {/* SUPPORT */}
          <div style={sectionStyle(ORANGE)}>
            {sectionHead("Processus Support", (carto.support||[]).length, "#185FA5")}
            <SearchableList items={PROCESSUS_CATALOG.support} selected={carto.support||[]} onToggle={it=>toggle("support",it)} onAddCustom={v=>addCustom("support",v)} accent="#185FA5" placeholder="Rechercher un processus support…"/>
            <div style={{marginTop:"12px",paddingTop:"12px",borderTop:"1px dashed #E2E5E2"}}>
              <div style={{fontSize:"11px",fontWeight:"700",color:"#185FA5",textTransform:"uppercase",letterSpacing:"1px",marginBottom:"6px"}}>Ordre dans la cartographie</div>
              <OrderableList family="support" color="#185FA5"/>
            </div>
          </div>
        </div>

        {/* ---- Zone droite : entrées/sorties + export ---- */}
        <div style={{position:"sticky",top:"24px"}}>

          {/* Aperçu */}
          <div style={{background:"#fff",borderRadius:"14px",border:"1px solid #E2E5E2",padding:"20px",marginBottom:"16px"}}>
            <div style={{fontWeight:"800",fontSize:"14px",color:"#15201B",marginBottom:"14px"}}>Aperçu de la structure</div>

            {/* Simulation visuelle */}
            <div style={{background:`${ORANGE}0A`,borderRadius:"10px",padding:"12px",marginBottom:"10px",border:`1px solid ${ORANGE}30`}}>
              <div style={{fontSize:"10px",fontWeight:"700",color:ORANGE,textTransform:"uppercase",letterSpacing:"1px",textAlign:"center",marginBottom:"8px"}}>Processus Management</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:"4px",justifyContent:"center"}}>
                {(carto.management||[]).map(it=><span key={it} style={{background:ORANGE+"20",borderRadius:"4px",padding:"2px 7px",fontSize:"10px",color:"#555"}}>{it}</span>)}
                {(carto.management||[]).length===0&&<span style={{fontSize:"11px",color:"#aaa",fontStyle:"italic"}}>Aucun</span>}
              </div>
            </div>
            <div style={{background:"#FAD5C218",borderRadius:"10px",padding:"12px",marginBottom:"10px",border:"1px solid #FAD5C260"}}>
              <div style={{fontSize:"10px",fontWeight:"700",color:"#C05A20",textTransform:"uppercase",letterSpacing:"1px",textAlign:"center",marginBottom:"8px"}}>Processus Réalisation</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:"4px",justifyContent:"center"}}>
                {(carto.realisation||[]).map(it=><span key={it} style={{background:"#F4B9A030",borderRadius:"4px",padding:"2px 7px",fontSize:"10px",color:"#555"}}>{it}</span>)}
                {(carto.realisation||[]).length===0&&<span style={{fontSize:"11px",color:"#aaa",fontStyle:"italic"}}>Aucun</span>}
              </div>
            </div>
            <div style={{background:"#185FA508",borderRadius:"10px",padding:"12px",border:"1px solid #185FA520"}}>
              <div style={{fontSize:"10px",fontWeight:"700",color:"#185FA5",textTransform:"uppercase",letterSpacing:"1px",textAlign:"center",marginBottom:"8px"}}>Processus Support</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:"4px",justifyContent:"center"}}>
                {(carto.support||[]).map(it=><span key={it} style={{background:"#185FA518",borderRadius:"4px",padding:"2px 7px",fontSize:"10px",color:"#555"}}>{it}</span>)}
                {(carto.support||[]).length===0&&<span style={{fontSize:"11px",color:"#aaa",fontStyle:"italic"}}>Aucun</span>}
              </div>
            </div>
          </div>

          {/* Entrées */}
          <div style={{background:"#fff",borderRadius:"14px",border:"1px solid #E2E5E2",padding:"20px",marginBottom:"16px"}}>
            <label style={{display:"block",fontWeight:"800",fontSize:"13px",color:"#15201B",marginBottom:"8px"}}>
              Entrées <span style={{fontWeight:400,color:"#9AA0A6",fontSize:"11px"}}>(une par ligne)</span>
            </label>
            <textarea
              style={{width:"100%",background:"#F6F7F6",border:"1px solid #E2E5E2",borderRadius:"9px",padding:"10px 12px",fontFamily:"'Syne',sans-serif",fontSize:"13px",color:"#1F2421",resize:"vertical",minHeight:"80px",outline:"none"}}
              value={carto.inputs||""}
              onChange={e=>setCarto(prev=>({...prev,inputs:e.target.value}))}
              placeholder={"Exigences Clients\nRéglementaires et légales\nAnalyse des risques"}
            />
          </div>

          {/* Sorties */}
          <div style={{background:"#fff",borderRadius:"14px",border:"1px solid #E2E5E2",padding:"20px",marginBottom:"20px"}}>
            <label style={{display:"block",fontWeight:"800",fontSize:"13px",color:"#15201B",marginBottom:"8px"}}>
              Sorties <span style={{fontWeight:400,color:"#9AA0A6",fontSize:"11px"}}>(une par ligne)</span>
            </label>
            <textarea
              style={{width:"100%",background:"#F6F7F6",border:"1px solid #E2E5E2",borderRadius:"9px",padding:"10px 12px",fontFamily:"'Syne',sans-serif",fontSize:"13px",color:"#1F2421",resize:"vertical",minHeight:"80px",outline:"none"}}
              value={carto.outputs||""}
              onChange={e=>setCarto(prev=>({...prev,outputs:e.target.value}))}
              placeholder={"Satisfaction client\nPerformance SMI\nSécurité de l'information"}
            />
          </div>

          {/* Bouton export */}
          <button
            onClick={handleExport}
            style={{width:"100%",background:"linear-gradient(135deg,#F07040,#C0501A)",border:"none",borderRadius:"12px",color:"#fff",padding:"16px",fontSize:"15px",fontWeight:"800",fontFamily:"'Syne',sans-serif",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:"10px",boxShadow:"0 10px 24px rgba(200,80,26,.28)",transition:"transform .1s"}}
          >
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M9 13h6M9 17h4"/></svg>
            Exporter en Word (.docx)
          </button>

          <div style={{marginTop:"12px",fontSize:"11px",color:"#9AA0A6",textAlign:"center",lineHeight:1.5}}>
            Le fichier sera au format A4 paysage avec votre logo et les couleurs OCG.
          </div>
        </div>
      </div>

      {toast && <div style={styles.toast}>{toast}</div>}
    </div>
  );
}

/* =========================================================================
   MODULE CODIFICATION (App principale — inchangée sauf minor refactor)
   ========================================================================= */
const EMPTY_COUNTERS={PS:0,PR:0,ENG:0,F:0,L:0,INS:0,MO:0};

const mkIcon=(children)=>({size=18,...rest})=><svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...rest}>{children}</svg>;
const IcSparkles=mkIcon(<><path d="M12 3l1.8 4.7L18.5 9.5 13.8 11.3 12 16l-1.8-4.7L5.5 9.5l4.7-1.8z"/><path d="M19 14l.6 1.6 1.6.6-1.6.6-.6 1.6-.6-1.6-1.6-.6 1.6-.6z"/></>);
const IcFile=mkIcon(<><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></>);
const IcBookOpen=mkIcon(<><path d="M2 4h7a3 3 0 0 1 3 3v13a2.5 2.5 0 0 0-2.5-2.5H2z"/><path d="M22 4h-7a3 3 0 0 0-3 3v13a2.5 2.5 0 0 1 2.5-2.5H22z"/></>);
const IcInfo=mkIcon(<><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></>);
const IcTrendUp=mkIcon(<><path d="M22 7l-8.5 8.5-5-5L2 17"/><path d="M16 7h6v6"/></>);
const IcCalendar=mkIcon(<><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></>);
const IcSearch=mkIcon(<><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></>);
const IcHistory=mkIcon(<><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l3 2"/></>);
const IcFileText=mkIcon(<><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M9 13h6M9 17h4"/></>);
const IcTrash=mkIcon(<><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></>);
const IcClipboardCheck=mkIcon(<><rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><path d="m9 14 2 2 4-4"/></>);
const IcGitBranch=mkIcon(<><rect x="3" y="3" width="6" height="6" rx="1"/><rect x="15" y="15" width="6" height="6" rx="1"/><path d="M6 9v6a3 3 0 0 0 3 3h6"/></>);
const IcDatabase=mkIcon(<><ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"/></>);
const IcListChecks=mkIcon(<><path d="m3 6 2 2 3-3"/><path d="m3 14 2 2 3-3"/><path d="M11 6h10M11 15h10"/></>);
const IcSitemap=mkIcon(<><rect x="9" y="3" width="6" height="5" rx="1"/><rect x="3" y="16" width="6" height="5" rx="1"/><rect x="15" y="16" width="6" height="5" rx="1"/><path d="M12 8v4M6 16v-2h12v2"/></>);
const IcGear=mkIcon(<><circle cx="12" cy="12" r="3.2"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M19.1 4.9 17 7M7 17l-2.1 2.1"/></>);
const IcChevronRight=mkIcon(<path d="m9 18 6-6-6-6"/>);
const IcFileSearch=mkIcon(<><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h5"/><path d="M14 2v6h6"/><circle cx="16.5" cy="16.5" r="2.5"/><path d="m21 21-1.6-1.6"/></>);
const IcLogout=mkIcon(<><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="m16 17 5-5-5-5M21 12H9"/></>);
const IcLayers=mkIcon(<><path d="m12 2 9 5-9 5-9-5 9-5z"/><path d="m3 12 9 5 9-5M3 17l9 5 9-5"/></>);
const IcArchive=mkIcon(<><rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8"/><path d="M10 12h4"/></>);
const IcRestore=mkIcon(<><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/></>);
const IcChevronDown=mkIcon(<path d="m6 9 6 6 6-6"/>);

const TYPE_ICONS={processus:IcClipboardCheck,procedure:IcGitBranch,enregistrement:IcDatabase,formulaire:IcListChecks,logigramme:IcSitemap,instruction:IcBookOpen,modeoperatoire:IcGear};
const TYPE_DESCRIPTIONS={processus:"Décrit une activité maîtrisée transformant des entrées en sorties (approche processus ISO).",procedure:"Définit la manière de réaliser une activité : qui fait quoi, quand et comment.",enregistrement:"Preuve documentée des résultats obtenus ou des activités réalisées (registres, listes).",formulaire:"Support vierge servant à collecter et tracer des informations (fiches à remplir).",logigramme:"Représentation graphique séquentielle d'un processus ou d'une procédure.",instruction:"Consigne détaillée pour exécuter une tâche précise à un poste donné.",modeoperatoire:"Description pas-à-pas d'une opération technique spécifique."};

function Combobox({value,onChange,placeholder,groups,searchPlaceholder="Rechercher…"}){
  const[open,setOpen]=useState(false);const[q,setQ]=useState("");const ref=useRef(null);
  useEffect(()=>{const onDoc=(e)=>{if(ref.current&&!ref.current.contains(e.target)){setOpen(false);setQ("");}};document.addEventListener("mousedown",onDoc);return()=>document.removeEventListener("mousedown",onDoc);},[]);
  const norm=(s)=>s.normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();const nq=norm(q.trim());
  const fgroups=groups.map(g=>({...g,items:g.items.filter(it=>!nq||norm(it.label).includes(nq))})).filter(g=>g.items.length>0);
  let current=placeholder;for(const g of groups)for(const it of g.items)if(it.value===value)current=it.label;
  const pick=(v)=>{onChange(v);setOpen(false);setQ("");};
  return(
    <div className="cbx" ref={ref}>
      <button type="button" className={"cbx-btn"+(value?" has-val":"")} onClick={()=>setOpen(o=>!o)}>
        <span className="cbx-val">{current}</span><span className="cbx-chev"><IcChevronDown size={16}/></span>
      </button>
      {open&&(<div className="cbx-pop">
        <div className="cbx-search"><IcSearch size={14}/><input autoFocus placeholder={searchPlaceholder} value={q} onChange={e=>setQ(e.target.value)}/></div>
        <div className="cbx-list">{fgroups.length===0?<div className="cbx-empty">Aucun résultat</div>:fgroups.map((g,gi)=><div key={gi}>{g.label&&<div className="cbx-group">{g.label}</div>}{g.items.map((it,ii)=><div key={ii} className={"cbx-opt"+(it.value===value?" sel":"")} onClick={()=>pick(it.value)}>{it.label}</div>)}</div>)}</div>
      </div>)}
    </div>
  );
}

function DocCodeApp({ user, onLogout, onUpdateUser, onDeleteAccount, onBack }) {
  const[type,setType]=useState("");const[customType,setCustomType]=useState("");const[customCode,setCustomCode]=useState("");const[docModel,setDocModel]=useState("");const[intitule,setIntitule]=useState("");const[query,setQuery]=useState("");
  const initial=loadUserState(user.email);
  const[codes,setCodes]=useState(initial?.codes||[]);const[counters,setCounters]=useState(initial?.counters||{...EMPTY_COUNTERS});const[history,setHistory]=useState(()=>loadUserHistory(user.email));const[snapshots,setSnapshots]=useState(()=>loadUserSnapshots(user.email));
  const[toast,setToast]=useState(null);const[showHistory,setShowHistory]=useState(false);const[showSettings,setShowSettings]=useState(false);const[showTypes,setShowTypes]=useState(false);const[showExports,setShowExports]=useState(false);

  useEffect(()=>{try{localStorage.setItem(stateKey(user.email),JSON.stringify({codes,counters}));}catch(e){}},[codes,counters,user.email]);
  useEffect(()=>{try{localStorage.setItem(histKey(user.email),JSON.stringify(history));}catch(e){}},[history,user.email]);
  useEffect(()=>{try{localStorage.setItem(snapKey(user.email),JSON.stringify(snapshots));}catch(e){}},[snapshots,user.email]);

  const showToast=(msg)=>{setToast(msg);setTimeout(()=>setToast(null),3000);};
  const handleLogoUpload=async(e)=>{const file=e.target.files&&e.target.files[0];if(!file)return;try{const{dataUrl,width,height}=await readImageFile(file,256);onUpdateUser({logo:dataUrl,logoW:width,logoH:height});showToast(" Logo mis à jour");}catch(err){showToast("❌ Image illisible");}e.target.value="";};
  const handleLogoRemove=()=>{onUpdateUser({logo:null,logoW:null,logoH:null});showToast(" Logo supprimé");};
  const handleDeleteAccount=()=>{const ok=window.confirm(`Supprimer définitivement le compte « ${user.company} » ?\n\nIrréversible.`);if(!ok)return;onDeleteAccount();};
  const handleSubmit=()=>{
    if(!type){showToast("⚠️ Veuillez sélectionner un type");return;}
    if(type==="__autre__"&&(!customType.trim()||!customCode.trim())){showToast("⚠️ Saisir le nom et code du type");return;}
    if(!intitule.trim()){showToast("⚠️ Veuillez saisir un intitulé");return;}
    const typeCode=type==="__autre__"?customCode.trim().toUpperCase().slice(0,3):TYPE_MAP_BACKEND[type];
    const typeLabel=type==="__autre__"?customType.trim():TYPE_LABELS_PLURAL[typeCode];
    const abbrev=makeAbbrev(intitule);const seq=(counters[typeCode]||0)+1;const code=`${typeCode}-${abbrev}-${String(seq).padStart(2,"0")}`;
    const entry={id:Date.now().toString(36)+Math.random().toString(36).slice(2,6),code,intitule:intitule.trim(),type:typeCode,typeLabel,seq,version:0,motif:"Création",createdAt:new Date().toISOString()};
    setCodes(prev=>[entry,...prev]);setCounters(prev=>({...prev,[typeCode]:seq}));setHistory(prev=>[entry,...prev]);
    showToast(`✅ Code généré : ${code}`);setIntitule("");setType("");setCustomType("");setCustomCode("");setDocModel("");
  };
  const handleDelete=(id)=>{setCodes(prev=>prev.filter(c=>c.id!==id));showToast("🗑️ Code supprimé du tableau");};
  const handleClearAll=()=>{if(codes.length===0){showToast("⚠️ Tableau déjà vide");return;}if(!window.confirm(`Vider le tableau (${codes.length} codes) et réinitialiser les compteurs ?`))return;setCodes([]);setCounters({...EMPTY_COUNTERS});showToast("🗑️ Tableau vidé");};
  const handleClearHistory=()=>{if(history.length===0){showToast("⚠️ Historique déjà vide");return;}if(!window.confirm(`Effacer les ${history.length} entrées ?`))return;setHistory([]);showToast("🗑️ Historique effacé");};
  const branding={company:user.company,logo:user.logo||null,logoW:user.logoW,logoH:user.logoH};
  const saveSnapshot=(format)=>{const snap={id:Date.now().toString(36)+Math.random().toString(36).slice(2,6),createdAt:new Date().toISOString(),format,count:codes.length,codes:codes.map(c=>({...c})),counters:{...counters}};setSnapshots(prev=>[snap,...prev]);};
  const handleRestoreSnapshot=(snap)=>{if(codes.length>0&&!window.confirm("Remplacer le tableau actuel par ce tableau exporté ?"))return;setCodes(snap.codes.map(c=>({...c})));setCounters({...EMPTY_COUNTERS,...snap.counters});setShowExports(false);showToast("✅ Tableau restauré");};
  const handleReexport=async(snap,fmt)=>{try{if(fmt==="word")await exportWord(snap.codes,branding);else await exportExcel(snap.codes,branding);showToast(fmt==="word"?"📄 Word téléchargé":"📊 Excel téléchargé");}catch(e){console.error(e);showToast("❌ Erreur export");}};
  const handleDeleteSnapshot=(id)=>{setSnapshots(prev=>prev.filter(s=>s.id!==id));showToast("🗑️ Tableau exporté supprimé");};
  const handleClearSnapshots=()=>{if(snapshots.length===0){showToast("⚠️ Aucun tableau exporté");return;}if(!window.confirm(`Supprimer les ${snapshots.length} tableaux exportés ?`))return;setSnapshots([]);showToast("🗑️ Tableaux exportés supprimés");};
  const handleExportWord=async()=>{if(codes.length===0){showToast("⚠️ Aucun code à exporter");return;}try{await exportWord(codes,branding);saveSnapshot("word");showToast("📄 Word téléchargé et sauvegardé");}catch(e){console.error(e);showToast("❌ Erreur export Word");}};
  const handleExportExcel=async()=>{if(codes.length===0){showToast("⚠️ Aucun code à exporter");return;}try{await exportExcel(codes,branding);saveSnapshot("excel");showToast("📊 Excel téléchargé et sauvegardé");}catch(e){console.error(e);showToast("❌ Erreur export Excel");}};

  const preview=previewCode(type,customCode,intitule,counters);
  const selectedColor=type&&type!=="__autre__"?TYPE_LABELS[type]?.color:"#3B6D11";
  const initials=(user.company||user.name||"?").trim().split(/\s+/).map(w=>w[0]).slice(0,2).join("").toUpperCase();
  const isFirstTime=codes.length===0&&history.length===0;
  const q=query.trim().toLowerCase();
  const filtered=q?codes.filter(c=>c.code.toLowerCase().includes(q)||c.intitule.toLowerCase().includes(q)):codes;
  const ARROW="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%239AA0A6' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E";
  const badge=(color)=>({display:"inline-flex",padding:"4px 10px",borderRadius:"8px",fontFamily:"'IBM Plex Mono', monospace",fontSize:"12px",fontWeight:800,background:`${color}18`,color});
  const statCard=(key)=>{const val=TYPE_LABELS[key];const Ic=TYPE_ICONS[key];return(<div className="stat-card" key={key} style={{background:`${val.color}0D`,borderColor:`${val.color}26`}}><div className="stat-top"><div className="stat-ic" style={{background:`${val.color}1F`,color:val.color}}><Ic size={16}/></div><div className="stat-name">{TYPE_LABELS_PLURAL[val.code]}</div></div><div className="stat-val">{String(counters[val.code]||0).padStart(2,"0")}</div><div className="stat-foot"><span className="stat-total">Total</span><span className="stat-trend"><IcTrendUp size={13}/> 0%</span></div></div>);};

  return(
    <div className="dcg">
      <style>{`
        .dcg,.dcg*{box-sizing:border-box;}.dcg{font-family:'Syne',sans-serif;color:#1F2421;background:#F4F6F5;min-height:100vh;}.dcg button{cursor:pointer;}
        @keyframes slideUp{from{transform:translateY(16px);opacity:0;}to{transform:translateY(0);opacity:1;}}@keyframes fadeIn{from{opacity:0;}to{opacity:1;}}
        .dcg ::-webkit-scrollbar{width:6px;height:6px;}.dcg ::-webkit-scrollbar-thumb{background:#D3D7D4;border-radius:3px;}
        .dcg-shell{display:grid;grid-template-columns:340px 1fr;min-height:100vh;}
        .dcg-side{background:#FFFFFF;border-right:1px solid #E7E9E7;padding:26px 26px 40px;overflow-y:auto;}
        .side-logo{display:flex;align-items:center;gap:14px;margin-bottom:22px;}
        .side-logo-mark{width:60px;height:60px;border-radius:16px;position:relative;flex-shrink:0;background:linear-gradient(150deg,#1AAE86,#0F6E56);color:#fff;display:flex;align-items:center;justify-content:center;font-size:30px;font-weight:800;box-shadow:0 8px 20px rgba(15,110,86,.28);}
        .side-logo-mark::after{content:"";position:absolute;inset:9px;border:2px solid rgba(255,255,255,.5);border-radius:9px;}
        .side-logo-mark.has-img{width:96px;height:96px;background:#fff;border:1px solid #E2E5E2;box-shadow:0 6px 16px rgba(0,0,0,.10);padding:6px;}.side-logo-mark.has-img::after{display:none;}
        .side-logo-mark img{width:100%;height:100%;object-fit:contain;}.side-logo-txt b{display:block;font-size:21px;font-weight:800;line-height:1.05;letter-spacing:-.5px;color:#15201B;}
        .side-logo-txt b span{color:#0F6E56;}.side-logo-txt small{font-size:11px;letter-spacing:3px;color:#9AA0A6;font-weight:700;}
        .side-head{display:flex;align-items:center;gap:8px;margin:24px 0 14px;font-size:12px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:#0F6E56;}
        .fld{margin-bottom:15px;}.fld-label{display:block;font-size:12.5px;font-weight:700;color:#5C635E;margin-bottom:7px;}
        .fld-input,.fld-select{width:100%;background:#F6F7F6;border:1px solid #E2E5E2;border-radius:11px;padding:13px 14px;font-size:14px;color:#1F2421;font-family:'Syne',sans-serif;outline:none;transition:border-color .2s,background .2s;}
        .fld-input:focus,.fld-select:focus{border-color:#1AAE86;background:#fff;}
        .fld-select{appearance:none;cursor:pointer;background-image:url("${ARROW}");background-repeat:no-repeat;background-position:right 14px center;}
        .cbx{position:relative;}.cbx-btn{width:100%;display:flex;align-items:center;justify-content:space-between;gap:10px;background:#F6F7F6;border:1px solid #E2E5E2;border-radius:11px;padding:13px 14px;font-size:14px;color:#1F2421;font-family:'Syne',sans-serif;text-align:left;transition:border-color .2s,background .2s;}
        .cbx-btn:hover{border-color:#CBD3CD;}.cbx-val{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}.cbx-btn:not(.has-val) .cbx-val{color:#9AA0A6;}.cbx-chev{color:#9AA0A6;flex-shrink:0;display:flex;}
        .cbx-pop{position:absolute;z-index:50;top:calc(100% + 6px);left:0;right:0;background:#fff;border:1px solid #E2E5E2;border-radius:12px;overflow:hidden;box-shadow:0 14px 34px rgba(0,0,0,.14);animation:fadeIn .12s ease;}
        .cbx-search{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid #EEF0EE;}.cbx-search svg{color:#A2A8A4;flex-shrink:0;}
        .cbx-search input{flex:1;border:none;outline:none;background:transparent;font-size:13.5px;font-family:'Syne',sans-serif;color:#1F2421;}
        .cbx-list{max-height:260px;overflow-y:auto;padding:6px;}.cbx-group{font-size:10.5px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#9AA0A6;padding:10px 10px 6px;}
        .cbx-opt{padding:9px 10px;border-radius:8px;font-size:13.5px;color:#39403B;cursor:pointer;}.cbx-opt:hover{background:#F1F6F3;}.cbx-opt.sel{background:#E7F4ED;color:#0F6E56;font-weight:700;}.cbx-empty{padding:18px;text-align:center;font-size:12.5px;color:#A2A8A4;}
        .autre-box{background:#F3F8F5;border:1px dashed #B9D9CC;border-radius:11px;padding:13px;margin-top:10px;}.autre-cap{font-size:11px;font-weight:700;color:#6FA593;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;}
        .autre-code-input{width:100%;text-align:center;text-transform:uppercase;letter-spacing:4px;font-family:'IBM Plex Mono',monospace;font-size:20px;font-weight:700;color:#0F6E56;background:#EAF6F1;border:2px solid #1AAE86;border-radius:9px;padding:10px;outline:none;}
        .autre-tag{font-size:10px;font-family:'IBM Plex Mono',monospace;color:#3B6D11;text-align:center;margin-top:6px;}
        .preview-box{background:#EAF6F1;border:1px dashed #9DD3C2;border-radius:12px;padding:14px 16px;margin:4px 0 16px;}.preview-cap{font-size:11px;font-weight:700;color:#6FA593;letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;}.preview-code{font-family:'IBM Plex Mono',monospace;font-size:24px;font-weight:700;letter-spacing:3px;}
        .btn-generate{width:100%;border:none;border-radius:12px;background:linear-gradient(135deg,#16996F,#0E5E48);color:#fff;padding:14px;font-size:15px;font-weight:800;font-family:'Syne',sans-serif;display:flex;align-items:center;justify-content:center;gap:9px;box-shadow:0 10px 24px rgba(15,110,86,.28);transition:transform .1s,box-shadow .2s;}
        .btn-generate:hover{box-shadow:0 12px 30px rgba(15,110,86,.36);}.btn-generate:active{transform:scale(.98);}
        .gen-hint{display:flex;align-items:center;gap:7px;font-size:12px;color:#9AA0A6;margin:12px 2px 0;}
        .ref-list{margin-top:6px;display:flex;flex-direction:column;gap:8px;}.ref-row{display:flex;align-items:center;gap:12px;padding:9px 12px;border:1px solid #EEF0EE;border-radius:11px;background:#FCFDFC;transition:background .15s,border-color .15s;}
        .ref-row:hover{background:#F4F8F6;border-color:#DDE6E1;}.ref-ic{width:30px;height:30px;border-radius:9px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}.ref-name{flex:1;font-size:13.5px;font-weight:600;color:#39403B;}.ref-badge{font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:800;padding:3px 9px;border-radius:7px;letter-spacing:.5px;}
        .btn-types{width:100%;margin-top:18px;border:none;border-radius:11px;background:#1C2B25;color:#EAF1EE;padding:13px 16px;font-size:13px;font-weight:700;font-family:'Syne',sans-serif;display:flex;align-items:center;justify-content:space-between;transition:background .2s;}.btn-types:hover{background:#24382F;}
        .btn-back-codi{width:100%;margin-top:10px;border:1px solid #E2E5E2;border-radius:11px;background:#F6F7F6;color:#5C635E;padding:12px 16px;font-size:13px;font-weight:700;font-family:'Syne',sans-serif;display:flex;align-items:center;justify-content:center;gap:6px;}.btn-back-codi:hover{background:#EDEEEC;}
        .dcg-main{padding:22px 30px 44px;overflow:auto;}.topbar{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:22px;}.topbar-title{font-size:21px;font-weight:800;color:#15201B;letter-spacing:-.4px;}.topbar-sub{font-size:12px;color:#9AA0A6;font-family:'IBM Plex Mono',monospace;margin-top:3px;}
        .user-wrap{display:flex;align-items:center;gap:14px;}.user-meta{text-align:right;}.user-name{font-size:13px;font-weight:700;color:#15201B;}.user-email{font-size:11px;color:#9AA0A6;font-family:'IBM Plex Mono',monospace;}
        .user-avatar{width:46px;height:46px;border-radius:50%;overflow:hidden;flex-shrink:0;display:flex;align-items:center;justify-content:center;border:1px solid #E2E5E2;color:#fff;font-weight:800;}
        .btn-logout{display:flex;align-items:center;gap:7px;padding:11px 16px;border-radius:10px;border:1px solid #F1C7C7;background:#FCEEEE;color:#C0392B;font-size:13px;font-weight:700;font-family:'Syne',sans-serif;white-space:nowrap;}.btn-logout:hover{background:#FAE1E1;}
        .welcome{background:linear-gradient(120deg,#11815F 0%,#0B5240 100%);border-radius:18px;padding:22px 26px;margin-bottom:26px;color:#fff;display:flex;align-items:center;justify-content:space-between;gap:18px;flex-wrap:wrap;box-shadow:0 16px 36px rgba(11,82,64,.22);}
        .welcome-left{display:flex;align-items:center;gap:18px;}.welcome-logo{width:96px;height:96px;border-radius:18px;background:#fff;flex-shrink:0;padding:10px;display:flex;align-items:center;justify-content:center;box-shadow:0 6px 18px rgba(0,0,0,.16);}.welcome-logo img{max-width:100%;max-height:100%;object-fit:contain;}
        .welcome-h{font-size:25px;font-weight:800;letter-spacing:-.5px;}.welcome-p{font-size:13px;opacity:.9;margin-top:5px;}
        .welcome-pill{display:flex;align-items:center;gap:8px;padding:10px 16px;border-radius:12px;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);font-size:13px;font-weight:700;white-space:nowrap;}
        .sec-head-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}.sec-head{display:flex;align-items:center;gap:9px;font-size:16px;font-weight:800;color:#15201B;}.sec-head .sec-ic{color:#0F6E56;display:flex;}
        .year-chip{display:flex;align-items:center;gap:8px;padding:9px 14px;border-radius:10px;background:#fff;border:1px solid #E2E5E2;font-size:13px;font-weight:600;color:#5C635E;}
        .stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px;}.stats-grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:20px;}
        .stat-card{border-radius:12px;padding:9px 11px;border:1px solid;}.stat-top{display:flex;align-items:center;gap:8px;}.stat-ic{width:28px;height:28px;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}.stat-name{font-size:11.5px;font-weight:700;color:#39403B;}
        .stat-val{font-size:19px;font-weight:800;letter-spacing:-.5px;color:#15201B;margin:6px 0 0;line-height:1;}.stat-foot{display:flex;align-items:center;justify-content:space-between;margin-top:5px;}.stat-total{font-size:10.5px;color:#9AA0A6;}.stat-trend{display:flex;align-items:center;gap:3px;font-size:10.5px;font-weight:700;color:#2F8F46;}
        .codes-head{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin-bottom:16px;flex-wrap:wrap;}.codes-title{font-size:16px;font-weight:800;color:#15201B;}.codes-sub{font-size:12px;color:#9AA0A6;font-family:'IBM Plex Mono',monospace;margin-top:3px;}
        .toolbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
        .search-box{display:flex;align-items:center;gap:8px;background:#fff;border:1px solid #E2E5E2;border-radius:10px;padding:0 12px;}.search-box:focus-within{border-color:#1AAE86;}.search-box input{border:none;outline:none;background:transparent;padding:10px 0;font-size:13px;width:160px;font-family:'Syne',sans-serif;color:#1F2421;}.search-box svg{color:#A2A8A4;}
        .tbtn{display:flex;align-items:center;gap:7px;padding:10px 15px;border-radius:10px;font-size:13px;font-weight:700;font-family:'Syne',sans-serif;border:1px solid transparent;transition:filter .15s;}.tbtn:hover{filter:brightness(.97);}
        .tbtn-hist{background:#EFEDF9;border-color:#D9D4F0;color:#5B4FC4;}.tbtn-word{background:#E7F0FB;border-color:#C3DBF4;color:#1F66B0;}.tbtn-excel{background:#E7F4EA;border-color:#BFE2C6;color:#2F8F46;}.tbtn-clear{background:#FCEEEE;border-color:#F1C7C7;color:#C0392B;}.tbtn-snap{background:#FBF3E3;border-color:#ECD9AE;color:#9A6B12;}.tbtn-restore{background:#E6F3F1;border-color:#BDE0DA;color:#0F6E56;}
        .ctable{background:#fff;border:1px solid #E7E9E7;border-radius:16px;overflow:hidden;}.ctable-head,.ctable-row{display:grid;grid-template-columns:200px 1fr 120px 160px 90px;align-items:center;}
        .ctable-head{padding:13px 22px;background:#FAFBFA;border-bottom:1px solid #EEF0EE;font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#9AA0A6;font-weight:700;}
        .ctable-row{padding:14px 22px;border-bottom:1px solid #F2F4F2;transition:background .15s;}.ctable-row:last-child{border-bottom:none;}.ctable-row:hover{background:#FAFBFA;}
        .ccode{font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:14px;}.cintitule{font-size:13.5px;color:#39403B;}.cbadge{display:inline-flex;padding:3px 9px;border-radius:7px;font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;}.cdate{font-family:'IBM Plex Mono',monospace;font-size:11.5px;color:#9AA0A6;}
        .cdel{width:30px;height:30px;border-radius:8px;border:1px solid #E2E5E2;background:#fff;color:#9AA0A6;display:flex;align-items:center;justify-content:center;}.cdel:hover{background:#FCEEEE;border-color:#F1C7C7;color:#C0392B;}
        .empty-state{padding:54px 20px;text-align:center;}.empty-ic{width:64px;height:64px;border-radius:50%;background:#F2F4F2;color:#B9BFBA;display:flex;align-items:center;justify-content:center;margin:0 auto 16px;}.empty-t{font-size:14px;font-weight:700;color:#5C635E;}.empty-s{font-size:12.5px;color:#A2A8A4;margin-top:5px;}
        @media(max-width:920px){.dcg-shell{grid-template-columns:1fr;}.dcg-side{border-right:none;border-bottom:1px solid #E7E9E7;}.stats-grid,.stats-grid-3{grid-template-columns:repeat(2,1fr);}.ctable-head,.ctable-row{grid-template-columns:130px 1fr 80px 70px;}.ctable-head span:nth-child(4),.ctable-row span:nth-child(4){display:none;}}
      `}</style>

      <div className="dcg-shell">
        <aside className="dcg-side">
          <div className="side-logo">
            <div className={"side-logo-mark"+(BRAND_LOGO?" has-img":"")}>{BRAND_LOGO?<img src={BRAND_LOGO} alt="CoDoc"/>:"D"}</div>
            <div className="side-logo-txt"><b>CoDoc <span>Generator</span></b><small>QMS / ISO</small></div>
          </div>
          <div className="side-head"><IcFile size={15}/> Nouveau document</div>
          <div className="fld">
            <label className="fld-label">Type de document</label>
            <Combobox placeholder="-- Sélectionner --" searchPlaceholder="Tapez une lettre…" value={type} onChange={v=>{setType(v);setCustomType("");setCustomCode("");setDocModel("");}}
              groups={[{label:null,items:[{value:"processus",label:"Processus (PS)"},{value:"procedure",label:"Procédure (PR)"},{value:"enregistrement",label:"Enregistrement (ENG)"},{value:"formulaire",label:"Formulaire (F)"},{value:"logigramme",label:"Logigramme (L)"},{value:"instruction",label:"Instruction (INS)"},{value:"modeoperatoire",label:"Mode opératoire (MO)"},{value:"__autre__",label:" Autre type (personnalisé)..."}]}]}/>
            {type==="__autre__"&&(<div className="autre-box"><div className="autre-cap">Définir un nouveau type</div><div style={{marginBottom:"12px"}}><label className="fld-label">Nom du type</label><input className="fld-input" placeholder="ex: Rapport d'audit" value={customType} onChange={e=>{const v=e.target.value;setCustomType(v);setCustomCode(autoGenerateCode(v));}}/></div><label className="fld-label">Code (max 3 lettres)</label><input className="autre-code-input" maxLength={3} value={customCode} onChange={e=>setCustomCode(e.target.value.toUpperCase().replace(/[^A-Z]/g,""))}/><div className="autre-tag">↑ généré automatiquement — modifiable</div></div>)}
          </div>
          {type&&type!=="__autre__"&&DOC_CATALOG[type]&&(<div className="fld"><label className="fld-label">Document (référentiel)</label><Combobox placeholder="-- Choisir un document --" searchPlaceholder="Rechercher un document…" value={docModel} onChange={v=>{setDocModel(v);if(v==="__autre__")setIntitule("");else if(v!=="")setIntitule(v);}} groups={[...DOC_CATALOG[type].map(grp=>({label:grp.group,items:grp.items.map(it=>({value:it,label:it}))})),{label:null,items:[{value:"__autre__",label:" Autre (saisie libre)…"}]}]}/></div>)}
          <div className="fld"><label className="fld-label">Intitulé du document</label><input className="fld-input" placeholder="ex: Gestion des non conformités" value={intitule} onChange={e=>setIntitule(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleSubmit()}/></div>
          <div className="preview-box"><div className="preview-cap">Aperçu du code</div><div className="preview-code" style={{color:selectedColor}}>{preview}</div></div>
          <button className="btn-generate" onClick={handleSubmit}><IcSparkles size={18}/> Générer le code</button>
          <div className="gen-hint"><IcInfo size={14}/> Remplissez les champs pour générer un code.</div>
          <div className="side-head"><IcBookOpen size={15}/> Référence</div>
          <div className="ref-list">{Object.entries(TYPE_LABELS).map(([key,val])=>{const Ic=TYPE_ICONS[key];return(<div className="ref-row" key={key}><div className="ref-ic" style={{background:`${val.color}1A`,color:val.color}}><Ic size={16}/></div><span className="ref-name">{val.label}</span><span className="ref-badge" style={{background:`${val.color}18`,color:val.color}}>{val.code}</span></div>);})}</div>
          <button className="btn-types" onClick={()=>setShowTypes(true)}>Voir la description des types <IcChevronRight size={16}/></button>
          <button className="btn-back-codi" onClick={onBack}>← Retour aux modules</button>
        </aside>

        <main className="dcg-main">
          <div className="topbar">
            <div><div className="topbar-title">Codification Documentaire</div><div className="topbar-sub">Système de codification QMS/ISO</div></div>
            <div className="user-wrap">
              <div className="user-meta"><div className="user-name">{user.name}</div><div className="user-email">{user.email}</div></div>
              <div className="user-avatar" style={{background:user.logo?"#fff":"#0F6E56"}} onClick={()=>setShowSettings(true)} title="Logo">
                {user.logo?<img src={user.logo} alt="logo" style={{width:"100%",height:"100%",objectFit:"contain"}}/>:initials}
              </div>
              <button className="btn-logout" onClick={onLogout}><IcLogout size={16}/> Déconnexion</button>
            </div>
          </div>
          <div className="welcome">
            <div className="welcome-left">{user.logo&&<div className="welcome-logo"><img src={user.logo} alt="logo"/></div>}<div><div className="welcome-h">Bonjour, {user.company} </div><div className="welcome-p">{isFirstTime?"Bienvenue ! Votre espace est vierge — créez votre premier document.":`Heureux de vous revoir, ${user.name}.`}</div></div></div>
            <div className="welcome-pill"><IcLayers size={16}/> {history.length} code{history.length!==1?"s":""} au total</div>
          </div>
          <div className="sec-head-row"><div className="sec-head"><span className="sec-ic"><IcTrendUp size={18}/></span> Aperçu par type</div><div className="year-chip"><IcCalendar size={15}/> Cette année</div></div>
          <div className="stats-grid">{["processus","procedure","enregistrement","formulaire"].map(statCard)}</div>
          <div className="stats-grid-3">{["logigramme","instruction","modeoperatoire"].map(statCard)}</div>
          <div className="codes-head">
            <div><div className="codes-title">Codes enregistrés</div><div className="codes-sub">{codes.length} code{codes.length!==1?"s":""} enregistré{codes.length!==1?"s":""}</div></div>
            <div className="toolbar">
              <div className="search-box"><IcSearch size={15}/><input placeholder="Rechercher..." value={query} onChange={e=>setQuery(e.target.value)}/></div>
              <button className="tbtn tbtn-hist" onClick={()=>setShowHistory(true)}><IcHistory size={15}/> Historique{history.length?` (${history.length})`:""}</button>
              <button className="tbtn tbtn-snap" onClick={()=>setShowExports(true)}><IcArchive size={15}/> Exportés{snapshots.length?` (${snapshots.length})`:""}</button>
              <button className="tbtn tbtn-word" onClick={handleExportWord}><IcFileText size={15}/> Word</button>
              <button className="tbtn tbtn-excel" onClick={handleExportExcel}><IcFileText size={15}/> Excel</button>
              <button className="tbtn tbtn-clear" onClick={handleClearAll}><IcTrash size={15}/> Effacer</button>
            </div>
          </div>
          <div className="ctable">
            <div className="ctable-head"><span>Code</span><span>Intitulé</span><span>Type</span><span>Date</span><span>Actions</span></div>
            {codes.length===0?(<div className="empty-state"><div className="empty-ic"><IcFileSearch size={28}/></div><div className="empty-t">Aucun code enregistré.</div><div className="empty-s">Générez votre premier code pour le voir ici.</div></div>):filtered.length===0?(<div className="empty-state"><div className="empty-ic"><IcSearch size={26}/></div><div className="empty-t">Aucun résultat</div><div className="empty-s">Aucun code ne correspond à « {query} ».</div></div>):filtered.map(c=>{const typeInfo=Object.values(TYPE_LABELS).find(t=>t.code===c.type);const color=typeInfo?.color||"#888780";return(<div className="ctable-row" key={c.id}><span className="ccode" style={{color}}>{c.code}</span><span className="cintitule">{c.intitule}</span><span><span className="cbadge" style={{background:`${color}18`,color}}>{c.type}</span></span><span className="cdate">{new Date(c.createdAt).toLocaleString("fr-FR",{dateStyle:"short",timeStyle:"short"})}</span><span><button className="cdel" onClick={()=>handleDelete(c.id)} title="Supprimer"><IcTrash size={14}/></button></span></div>);})}
          </div>
        </main>
      </div>

      {showSettings&&(<div style={{...styles.modalOverlay,animation:"fadeIn 0.2s ease"}} onClick={()=>setShowSettings(false)}><div style={{...styles.modal,maxWidth:"460px"}} onClick={e=>e.stopPropagation()}><div style={styles.modalHead}><div><div style={styles.modalTitle}> Logo de l'entreprise</div><div style={styles.modalSub}>{user.company}</div></div><button style={styles.closeX} onClick={()=>setShowSettings(false)} title="Fermer">×</button></div><div style={{padding:"28px",textAlign:"center"}}><div style={styles.logoPreviewBox}>{user.logo?<img src={user.logo} alt="logo" style={{maxWidth:"100%",maxHeight:"100%",objectFit:"contain"}}/>:<span style={{fontFamily:"'IBM Plex Mono', monospace",fontSize:"12px",color:"#B4B2A9"}}>Aucun logo</span>}</div><div><label style={styles.logoUploadBtn}>{user.logo?"Changer le logo":"Importer un logo"}<input type="file" accept="image/*" style={{display:"none"}} onChange={handleLogoUpload}/></label>{user.logo&&<button style={styles.logoRemoveBtn} onClick={handleLogoRemove}>Supprimer</button>}</div><p style={{fontSize:"11px",color:"#888780",marginTop:"18px",lineHeight:1.6}}>PNG ou JPG. Le logo apparaît dans les exports.</p><div style={{borderTop:"1px solid #EEE9E0",margin:"22px 0 16px"}}/><div style={{fontSize:"11px",fontFamily:"'IBM Plex Mono', monospace",color:"#A53030",letterSpacing:"1px",textTransform:"uppercase",marginBottom:"10px"}}>Zone dangereuse</div><button onClick={handleDeleteAccount} style={{display:"inline-flex",alignItems:"center",justifyContent:"center",gap:"8px",padding:"12px 22px",borderRadius:"9px",cursor:"pointer",border:"1px solid #E8B4B4",background:"#FBEAEA",color:"#A53030",fontFamily:"'Syne', sans-serif",fontSize:"14px",fontWeight:"800"}}>🗑️ Supprimer ce compte</button><p style={{fontSize:"10.5px",color:"#B4B2A9",marginTop:"10px",lineHeight:1.6}}>Efface définitivement le compte. Irréversible.</p></div></div></div>)}

      {showTypes&&(<div style={{...styles.modalOverlay,animation:"fadeIn 0.2s ease"}} onClick={()=>setShowTypes(false)}><div style={{...styles.modal,maxWidth:"560px"}} onClick={e=>e.stopPropagation()}><div style={styles.modalHead}><div><div style={styles.modalTitle}>📚 Description des types</div><div style={styles.modalSub}>Codification & rôle de chaque type</div></div><button style={styles.closeX} onClick={()=>setShowTypes(false)}>×</button></div><div style={styles.modalBody}>{Object.entries(TYPE_LABELS).map(([key,val])=>(<div key={key} style={{display:"flex",gap:"14px",padding:"16px 24px",borderBottom:"1px solid #F1EFE8",alignItems:"flex-start"}}><span style={badge(val.color)}>{val.code}</span><div><div style={{fontWeight:800,fontSize:"14px",color:"#2C2C2A"}}>{val.label}</div><div style={{fontSize:"12.5px",color:"#6C6F6A",marginTop:"3px",lineHeight:1.5}}>{TYPE_DESCRIPTIONS[key]}</div></div></div>))}</div></div></div>)}

      {showHistory&&(<div style={{...styles.modalOverlay,animation:"fadeIn 0.2s ease"}} onClick={()=>setShowHistory(false)}><div style={styles.modal} onClick={e=>e.stopPropagation()}><div style={styles.modalHead}><div><div style={styles.modalTitle}> Historique</div><div style={styles.modalSub}>Archive de {user.company}</div></div><button style={styles.closeX} onClick={()=>setShowHistory(false)}>×</button></div><div style={styles.modalBody}>{history.length===0?<div style={styles.empty}>Aucun code dans l'historique</div>:<>{<div style={styles.histHeaderRow}><span>Code</span><span>Intitulé</span><span>Type</span><span>Date</span></div>}{history.map(c=>{const typeInfo=Object.values(TYPE_LABELS).find(t=>t.code===c.type);const color=typeInfo?.color||"#888780";return(<div key={c.id} style={styles.histRow}><span style={{...styles.codeBadge,color}}>{c.code}</span><span style={{fontSize:"13px",color:"#444441"}}>{c.intitule}</span><span><span style={{...styles.typeBadge,background:`${color}18`,color}}>{c.type}</span></span><span style={styles.dateText}>{new Date(c.createdAt).toLocaleString("fr-FR",{dateStyle:"short",timeStyle:"short"})}</span></div>);})}</>}</div><div style={styles.modalFoot}><div style={{fontSize:"12px",fontFamily:"'IBM Plex Mono', monospace",color:"#888780"}}>{history.length} entrée{history.length!==1?"s":""}</div><button style={styles.clearBtn} onClick={handleClearHistory}>🗑️ Vider l'historique</button></div></div></div>)}

      {showExports&&(<div style={{...styles.modalOverlay,animation:"fadeIn 0.2s ease"}} onClick={()=>setShowExports(false)}><div style={styles.modal} onClick={e=>e.stopPropagation()}><div style={styles.modalHead}><div><div style={styles.modalTitle}>Tableaux exportés</div><div style={styles.modalSub}>Restaurez ou réexportez un tableau sauvegardé</div></div><button style={styles.closeX} onClick={()=>setShowExports(false)}>×</button></div><div style={styles.modalBody}>{snapshots.length===0?<div style={styles.empty}>Aucun tableau exporté.<br/>Exportez en Word ou Excel pour créer une sauvegarde.</div>:snapshots.map(s=>(<div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:"12px",padding:"14px 24px",borderBottom:"1px solid #F1EFE8",flexWrap:"wrap"}}><div style={{minWidth:"180px"}}><div style={{fontWeight:800,fontSize:"13.5px",color:"#2C2C2A"}}>{new Date(s.createdAt).toLocaleString("fr-FR",{dateStyle:"short",timeStyle:"short"})}</div><div style={{fontSize:"11.5px",color:"#888780",fontFamily:"'IBM Plex Mono', monospace",marginTop:"2px"}}>{s.count} code{s.count!==1?"s":""} · {s.format==="word"?"Word":"Excel"}</div></div><div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}><button className="tbtn tbtn-restore" onClick={()=>handleRestoreSnapshot(s)}><IcRestore size={15}/> Restaurer</button><button className="tbtn tbtn-word" onClick={()=>handleReexport(s,"word")}><IcFileText size={15}/> Word</button><button className="tbtn tbtn-excel" onClick={()=>handleReexport(s,"excel")}><IcFileText size={15}/> Excel</button><button className="cdel" onClick={()=>handleDeleteSnapshot(s.id)} title="Supprimer"><IcTrash size={14}/></button></div></div>))}</div><div style={styles.modalFoot}><div style={{fontSize:"12px",fontFamily:"'IBM Plex Mono', monospace",color:"#888780"}}>{snapshots.length} tableau{snapshots.length!==1?"x":""} sauvegardé{snapshots.length!==1?"s":""}</div><button style={styles.clearBtn} onClick={handleClearSnapshots}> Tout vider</button></div></div></div>)}

      {toast&&<div style={styles.toast}>{toast}</div>}
    </div>
  );
}

/* =========================================================================
   ROUTEUR PRINCIPAL
   ========================================================================= */
export default function App() {
  const[session,setSession]=useState(()=>loadSession());
  const[accounts,setAccounts]=useState(()=>loadAccounts());
  const[activeModule,setActiveModule]=useState(null); // null | "codification" | "cartographie"

  useEffect(()=>{if(session&&!accounts[session]){clearSession();setSession(null);setActiveModule(null);}},[session,accounts]);

  const user=session&&accounts[session]?accounts[session]:null;

  const handleAuthenticated=(email)=>{
    // Effacer la cartographie au login → page vierge à chaque connexion
    try{localStorage.removeItem(cartoKey(email));}catch(e){}
    setAccounts(loadAccounts());saveSession(email);setSession(email);setActiveModule(null);
  };
  const handleLogout=()=>{
    // Effacer la cartographie au logout → page vierge à la prochaine connexion
    if(session){try{localStorage.removeItem(cartoKey(session));}catch(e){}}
    clearSession();setSession(null);setActiveModule(null);
  };

  const deleteAccount=()=>{
    if(!session)return;const email=session;
    setAccounts(prev=>{const next={...prev};delete next[email];saveAccounts(next);return next;});
    try{localStorage.removeItem(stateKey(email));}catch(e){}
    try{localStorage.removeItem(histKey(email));}catch(e){}
    try{localStorage.removeItem(snapKey(email));}catch(e){}
    try{localStorage.removeItem(cartoKey(email));}catch(e){}
    clearSession();setSession(null);setActiveModule(null);
  };

  const updateUser=(patch)=>{
    setAccounts(prev=>{if(!session||!prev[session])return prev;const next={...prev,[session]:{...prev[session],...patch}};saveAccounts(next);return next;});
  };

  if(!user)return <AuthScreen onAuthenticated={handleAuthenticated}/>;
  if(!activeModule)return <ModuleSelector user={user} onSelect={setActiveModule} onLogout={handleLogout}/>;
  if(activeModule==="cartographie")return <CartoApp key={user.email} user={user} onBack={()=>{try{localStorage.removeItem(cartoKey(user.email));}catch(e){}setActiveModule(null);}} onLogout={handleLogout}/>;
  return <DocCodeApp key={user.email} user={user} onLogout={handleLogout} onUpdateUser={updateUser} onDeleteAccount={deleteAccount} onBack={()=>setActiveModule(null)}/>;
}