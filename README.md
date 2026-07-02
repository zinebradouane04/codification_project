# CoDoc Generator

**Générateur de codes documentaires pour systèmes de management ISO / QMS.**

Application web 100 % autonome (sans backend) qui codifie automatiquement les documents d'un Système de Management Intégré (SMI) selon une convention `TYPE-ABRÉVIATION-NN`, gère un référentiel documentaire ISO multi-normes, et exporte la **Liste des Informations Documentées** en Word et Excel — le tout depuis le navigateur.

Développée par **OCG Consulting** (cabinet de conseil et certification ISO, Casablanca / Settat).

---

## Sommaire

1. [Présentation](#présentation)
2. [Fonctionnalités](#fonctionnalités)
3. [Logique de codification](#logique-de-codification)
4. [Référentiel documentaire](#référentiel-documentaire)
5. [Exports Word & Excel](#exports-word--excel)
6. [Authentification & données](#authentification--données)
7. [Stack technique](#stack-technique)
8. [Installation & déploiement](#installation--déploiement)
9. [Personnalisation](#personnalisation)
10. [Sécurité & limitations](#sécurité--limitations)
11. [Pistes d'évolution](#pistes-dévolution)

---

## Présentation

CoDoc Generator répond à un besoin concret d'accompagnement ISO : produire rapidement des **codes de documents normalisés et cohérents** (procédures, processus, enregistrements, formulaires…), tenir à jour la **liste maîtresse des informations documentées** (exigence des normes ISO clause 7.5), et livrer un fichier prêt à intégrer dans le SMI du client.

L'application est **mono-fichier React** (`DocCodeApp.jsx` / `App.jsx`), sans serveur ni base de données. Toutes les données vivent dans le `localStorage` du navigateur, isolées par compte utilisateur. Elle se déploie sur n'importe quel hébergement statique (GitHub Pages, cPanel…).

Elle couvre nativement les normes : **ISO 9001, ISO 14001, ISO 45001, ISO 22000, ISO 13485, ISO 27001**.

---

## Fonctionnalités

### Codification automatique
- Génération d'un code à partir du **type de document** et de son **intitulé**.
- Compteur incrémental par type (numéro de séquence sur 2 chiffres).
- 7 types prédéfinis + possibilité de créer un **type personnalisé** avec code auto-suggéré et modifiable.
- Aperçu du code en temps réel avant validation.

### Référentiel documentaire intégré
- Catalogue de documents pré-listés (le « Catalogue OCG » + des groupes par norme ISO).
- **Combobox cherchable** insensible aux accents pour retrouver un document rapidement.
- Sélection d'un document du référentiel → l'intitulé est pré-rempli, ou saisie libre.

### Tableau de bord
- Statistiques par type de document (compteurs).
- Tableau des codes enregistrés avec **recherche** par code ou intitulé.
- Suppression unitaire (le code reste dans l'historique) ou remise à zéro complète.

### Historique permanent
- Chaque code généré est archivé dans un historique **conservé même après « Tout effacer »**.
- Vidage de l'historique possible mais explicite (confirmation).

### Tableaux exportés (instantanés / snapshots)
- À chaque export Word/Excel, une **copie figée** du tableau est sauvegardée.
- Un instantané peut être **restauré** (pour le modifier et le réexporter), **réexporté** tel quel, ou supprimé.

### Exports professionnels
- **Word** (`.docx`) et **Excel** (`.xlsx`) générés côté navigateur.
- Mise en page brandée : logo de l'entreprise, charte navy `#1F3864`, police Times New Roman, regroupement par type avec cellules fusionnées.

### Multi-comptes & image de marque
- Écran de connexion / inscription animé (panneau coulissant).
- Chaque compte porte son **nom, entreprise, email** et son **logo**.
- Le logo s'affiche dans l'en-tête, le bandeau d'accueil et les exports.
- Données **strictement séparées par compte** (clé = email).

---

## Logique de codification

### Format du code

```
TYPE - ABRÉVIATION - NN
 │         │          │
 │         │          └── numéro de séquence (compteur par type, ex : 01, 02…)
 │         └── abréviation dérivée de l'intitulé (max 5 lettres)
 └── code du type de document
```

**Exemple** : un *Enregistrement* intitulé « Compte rendu de revue de direction » →
`ENG-CRRD-01`

### Codes de type

| Type             | Code  | Couleur     |
|------------------|-------|-------------|
| Processus        | `PS`  | `#0F6E56`   |
| Procédure        | `PR`  | `#993C1D`   |
| Enregistrement   | `ENG` | `#534AB7`   |
| Formulaire       | `F`   | `#185FA5`   |
| Logigramme       | `L`   | `#3B6D11`   |
| Instruction      | `INS` | `#854F0B`   |
| Mode opératoire  | `MO`  | `#993556`   |

### Génération de l'abréviation (`makeAbbrev`)

1. L'intitulé est mis en majuscules et **désaccentué** (`NFD` + suppression des diacritiques).
2. Les **mots vides** sont ignorés (`de`, `des`, `et`, `la`, `le`, `du`, `pour`, `avec`… ainsi que les mots génériques `processus`, `procédure`, `fiche`, etc. listés dans `IGNORE`).
3. On prend la **première lettre de chaque mot restant**, limité à 5 caractères.
4. Cas particulier : si un seul mot subsiste, on prend ses 3 premières lettres.
5. Repli : `DOC` si rien ne reste.

### Type personnalisé (`autoGenerateCode`)

Pour un type « Autre », le code est proposé automatiquement à partir du nom saisi :
- plusieurs mots → initiales (max 3) ;
- un seul mot → 1 à 3 lettres selon sa longueur.

Le code reste **modifiable manuellement** (3 lettres max, majuscules).

### Compteurs

Les compteurs sont tenus **par code de type** (`PS`, `PR`, `ENG`…). À chaque génération, le numéro de séquence est `compteur + 1`, formaté sur 2 chiffres. Le compteur est persistant et se réinitialise via « Tout effacer ».

---

## Référentiel documentaire

Le référentiel (`DOC_CATALOG`) alimente la liste déroulante « Document » selon le type choisi. Il combine un **« Catalogue OCG »** transversal et des groupes spécialisés par norme.

| Type            | Groupes disponibles |
|-----------------|---------------------|
| Processus       | Catalogue OCG · Management · Réalisation · Support |
| Procédure       | Catalogue OCG · Management · ISO 9001 · ISO 14001 · ISO 45001 · ISO 22000 · ISO 13485 · ISO 27001 |
| Instruction     | Instructions de travail |
| Formulaire      | Catalogue OCG · Formulaires |
| Enregistrement  | Catalogue OCG · Enregistrements |

La recherche dans le `Combobox` est **insensible à la casse et aux accents** (normalisation `NFD`), ce qui permet de taper « securite » et trouver « sécurité ».

---

## Exports Word & Excel

Les deux exports produisent la même structure : **Liste des Informations Documentées**, en paysage, regroupée par type de document.

**Colonnes** : Type de document · Intitulé · Code · Version · Date · Motif de mise à jour.

### Excel (`exportExcel`, via ExcelJS)
- Feuille « Informations documentées », orientation paysage, ajustée à la largeur.
- Bloc logo (ou nom d'entreprise sur fond navy si pas de logo) + bandeau titre + bloc code/version/date.
- En-têtes sur fond navy `#1F3864`, police Times New Roman.
- Regroupement par type avec **fusion verticale** de la colonne « Type ».
- Volet figé sous l'en-tête ; format de date `dd/mm/yyyy` ; colonne « Code » en Consolas.
- Fichier : `liste-informations-documentees.xlsx`.

### Word (`exportWord`, via docx)
- En-tête : logo (si présent) + nom de l'entreprise + titre + ligne `Code | Version | Date`.
- Tableau pleine largeur, bordures grises, en-têtes navy, colonne « Type » fusionnée par groupe (`rowSpan`).
- Section en orientation paysage.
- Fichier : `liste-informations-documentees.docx`.

> L'ordre des types suit `TYPE_ORDER` = `PS, PR, ENG, F, L, INS, MO`, les types personnalisés étant ajoutés à la suite.

---

## Authentification & données

> ⚠️ **Sécurité locale uniquement.** L'authentification sert à séparer les espaces de travail sur un même navigateur. Ce **n'est pas** une protection cryptographique et il n'y a **aucune synchronisation** entre postes.

- Les comptes sont stockés dans `localStorage`.
- Le mot de passe n'est jamais stocké en clair : il passe par `hashPassword` (hachage **simple, non cryptographique** — type FNV/`Math.imul`), juste pour éviter le clair.
- Pas de récupération de mot de passe, pas de partage multi-postes.
- À la première connexion d'un compte : espace vierge + message « Bonjour, [entreprise] ».

### Clés `localStorage`

| Clé | Contenu |
|-----|---------|
| `doccode_accounts_v1` | Tous les comptes (email, nom, entreprise, hash, logo…) |
| `doccode_session_v1` | Email du compte connecté |
| `doccode_state_v1__{email}` | Codes courants + compteurs |
| `doccode_history_v1__{email}` | Historique permanent des codes |
| `doccode_snapshots_v1__{email}` | Tableaux exportés (instantanés) |

### Gestion du logo
- Image importée → redimensionnée (max 256 px côté long) → encodée en **base64 PNG** sur un `<canvas>`.
- Ce maintien d'une image légère garde le `localStorage` raisonnable.
- Conversion base64 → `Uint8Array` pour l'injection dans les exports docx (`ImageRun`).

---

## Stack technique

- **React** (hooks : `useState`, `useEffect`, `useRef`) — application mono-fichier.
- **[exceljs](https://github.com/exceljs/exceljs)** — génération des fichiers `.xlsx`.
- **[docx](https://github.com/dolanmiu/docx)** — génération des fichiers `.docx`.
- **Aucun backend, aucune API externe** : tout s'exécute dans le navigateur.
- Polices web attendues : `Syne`, `IBM Plex Mono` (interface), `Times New Roman` / `Consolas` (exports).
- Icônes : SVG inline (aucune dépendance d'icônes).

### Dépendances & assets

```jsx
import React, { useState, useEffect, useRef } from "react";
import ExcelJS from "exceljs";
import { Document, Packer, /* … */ } from "docx";
import brandLogo from "./logo_codoc.png"; // logo de l'app (sidebar)
import logoOcg   from "./logo_ocg.png";   // logo OCG (écran de connexion)
```

---

## Installation & déploiement

### Prérequis
- Node.js + un projet React (Vite ou Create React App).

### Mise en place
1. Placer le composant dans `src/` (ex. `App.jsx`).
2. Ajouter les deux images : `src/logo_codoc.png` et `src/logo_ocg.png`.
3. Installer les dépendances :
   ```bash
   npm install exceljs docx
   ```
4. (Recommandé) Charger les polices `Syne` et `IBM Plex Mono` via Google Fonts dans `index.html`.

### Lancement local
```bash
npm run dev      # Vite
# ou
npm start        # Create React App
```

### Déploiement (GitHub Pages)
1. Build : `npm run build`.
2. Publier le dossier de build (`dist/` ou `build/`) sur la branche `gh-pages`.
3. Aucune configuration serveur n'est nécessaire — l'app est 100 % statique.

> Compatible également avec un hébergement cPanel classique (dépôt des fichiers statiques).

---

## Personnalisation

- **Logo de l'app** : remplacer `logo_codoc.png` ou pointer `BRAND_LOGO` vers une data URI base64.
- **Charte exports** : constantes `NAVY` / `NAVY_HEX` (`1F3864`) et `LIGHT` / `LIGHT_HEX` (`D9E1F2`).
- **Types de documents** : tables `TYPE_LABELS`, `TYPE_MAP_BACKEND`, `TYPE_LABELS_PLURAL`, `TYPE_ORDER`.
- **Référentiel** : objet `DOC_CATALOG` (ajout de groupes/documents par norme).
- **Mots vides de l'abréviation** : ensemble `IGNORE`.
- **Référence du document de liste** : code `ENG-LID-01` (modifiable dans les fonctions d'export).

---

## Sécurité & limitations

- 🔒 Authentification **locale et non cryptographique** : adaptée à un usage interne, pas à un déploiement multi-utilisateurs sensible.
- 💾 Données dans le `localStorage` : **liées à un navigateur/poste**, sans synchronisation ni sauvegarde cloud. Effacer les données du navigateur supprime tout.
- 🔁 Pas de récupération de mot de passe.
- 👥 Pour un vrai multi-utilisateur (partage, audit centralisé, comptes sécurisés) → un **backend serait nécessaire**.
- 📦 La taille du `localStorage` est limitée (généralement ~5–10 Mo) : les logos sont volontairement compressés pour rester légers.

---

## Pistes d'évolution

- Backend optionnel (auth réelle + base partagée) pour le travail collaboratif.
- Export/import JSON des comptes et tableaux (sauvegarde manuelle, transfert entre postes).
- Versionnement des documents (au-delà de la version `00` initiale) avec motif de mise à jour éditable.
- Gestion des révisions et dates de validité par document.
- Export PDF de la liste documentaire.

---

*CoDoc Generator — OCG Consulting · Codification documentaire QMS / ISO (9001 · 14001 · 45001 · 27001 · 22000 · 13485).*
