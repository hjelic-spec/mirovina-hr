// Mirovina HR – React SPA
// Deps: react, recharts, lucide-react, xlsx (SheetJS)
// Storage: persistent (window.storage API) + in-memory fallback
import { useState, useEffect, useCallback, useRef } from "react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, AreaChart, Area, PieChart, Pie, Cell
} from "recharts";
import * as XLSX from "xlsx";
import {
  User, LogOut, Calculator, BookOpen, Download, Save, ChevronRight,
  Info, AlertTriangle, CheckCircle, TrendingUp, Shield, Clock,
  DollarSign, Users, Settings, Menu, X, Eye, EyeOff, Plus, Trash2,
  BarChart2, PiggyBank, FileText, Star, ChevronDown, ChevronUp
} from "lucide-react";

// ─── PALETTE & CONSTANTS ────────────────────────────────────────────────────
const C = {
  navy:   "#1a2e52", navyL: "#243a68",
  gold:   "#c9922b", goldL: "#e6ab3a",
  slate:  "#4a5568", slateL: "#718096",
  cream:  "#f7f5f0", white:  "#ffffff",
  green:  "#2d7a4f", greenL: "#e8f5ee",
  red:    "#c0392b", redL:   "#fde8e7",
  blue:   "#2563eb", blueL:  "#eff6ff",
  purple: "#7c3aed", purpleL:"#f3f0ff",
  amber:  "#d97706", amberL: "#fffbeb",
  border: "#d1d5db", bg: "#f0eff9",
};
const AVM = 14.45; // EUR, od 1.7.2025.
const MIN_MIROVINA_FAKTOR = 1.06; // najniža = 106% AVM * staž * PF * MF
const MAX_VB = 3.8; // maksimalni vrijednosni bodovi

// ─── STORAGE HELPERS ────────────────────────────────────────────────────────
async function stGet(key) {
  try { const r = await window.storage.get(key); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
async function stSet(key, val) {
  try { await window.storage.set(key, JSON.stringify(val)); return true; }
  catch { return false; }
}
async function stDel(key) {
  try { await window.storage.delete(key); } catch {}
}

// ─── PENSION MATH ────────────────────────────────────────────────────────────
function calcStup1({ pvb, staz, pf, mf, avm = AVM }) {
  const ob     = pvb * staz * pf;
  const bruto  = ob * mf * avm;
  const najniza = MIN_MIROVINA_FAKTOR * avm * staz * pf * mf;
  const najvisa = staz * MAX_VB * pf * mf * avm;
  const efektivna = Math.max(najniza, Math.min(najvisa, bruto));
  return { ob, bruto, najniza, najvisa, efektivna };
}

function calcStup2FV({ bruto_placa, god_stednje, neto_prinos }) {
  const pmt  = bruto_placa * 0.05;
  const r_mj = neto_prinos / 12;
  const n    = god_stednje * 12;
  if (r_mj === 0) return pmt * n;
  return pmt * ((Math.pow(1 + r_mj, n) - 1) / r_mj);
}

function rentaDozivotna(kapital, e65 = 18) {
  return kapital / (e65 * 12);
}

// ─── UTILS ──────────────────────────────────────────────────────────────────
const fmt = (n, dec = 2) =>
  isNaN(n) ? "—" : n.toLocaleString("hr-HR", { minimumFractionDigits: dec, maximumFractionDigits: dec });
const fmtEur = (n) => `${fmt(n)} EUR`;
const pct = (n) => `${(n * 100).toFixed(2)} %`;

// ─── EXCEL EXPORT ────────────────────────────────────────────────────────────
// ─── EXCEL STILOVI ────────────────────────────────────────────────────────────
function xlCell(value, opts = {}) {
  // Build a rich cell object for SheetJS
  const cell = { v: value };
  if (value === null || value === undefined) { cell.v = ""; cell.t = "s"; return cell; }
  if (typeof value === "number") cell.t = "n";
  else if (value instanceof Date) { cell.t = "d"; }
  else cell.t = "s";
  if (opts.fmt) cell.z = opts.fmt;
  return cell;
}

function xlAppendSheet(wb, data, name) {
  // data: array of arrays; each element may be a raw value or xlCell()
  const ws = {};
  let maxCol = 0;
  data.forEach((row, r) => {
    if (!row) return;
    row.forEach((val, c) => {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (maxCol < c) maxCol = c;
      if (val !== null && val !== undefined && typeof val === "object" && "v" in val) {
        ws[addr] = val;
      } else {
        const cell = { v: val };
        if (typeof val === "number") cell.t = "n";
        else cell.t = "s";
        ws[addr] = cell;
      }
    });
  });
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length - 1, c: maxCol } });
  XLSX.utils.book_append_sheet(wb, ws, name);
}

function n(v, dec = 2) {
  if (v === null || v === undefined || isNaN(v)) return xlCell(0, { fmt: `#,##0.${"0".repeat(dec)}` });
  return xlCell(Math.round(v * Math.pow(10, dec)) / Math.pow(10, dec),
    { fmt: dec === 0 ? "#,##0" : `#,##0.${"0".repeat(dec)}` });
}
function eur(v) { return n(v, 2); }
function pctCell(v) { return xlCell(Math.round(v * 10000) / 10000, { fmt: "0.00%" }); }

function exportToExcel(scenariji, username) {
  const wb = XLSX.utils.book_new();
  const datum = new Date().toLocaleDateString("hr-HR");
  const DPS_MAX_U = 663.61; const DPS_S = 0.15;

  // ── SHEET 1: PREGLED SVIH SCENARIJA ──────────────────────────────────────
  const hdr1 = [
    "Naziv scenarija", "Tip", "Datum",
    // I. stup
    "Staž (god.)", "PVB", "PF", "MF", "AVM (EUR)",
    "Osobni bodovi", "Bruto mirovina I. (EUR/mj.)", "Efektivna mirovina I. (EUR/mj.)", "Godišnja I. stupa (EUR)",
    "Dodatak 27% (EUR/mj.)",
    // II. stup
    "Bruto plaća II. (EUR/mj.)", "God. štednje II.", "Neto prinos II. (%)", "Kapital II. (EUR)",
    "Jednokratna isplata % II.", "Jednokratna isplata EUR", "Preostali kapital II. (EUR)",
    "Aprox. doživotna renta II. (EUR/mj.)",
    // III. stup
    "Mjesečna uplata III. (EUR)", "God. štednje III.", "Prinos III. (%)", "Uplata poslodavca III. (EUR/mj.)",
    "DPS godišnje (EUR)", "Kapital III. (EUR)",
    // Ukupno
    "UKUPNO (EUR/mj.)", "UKUPNO godišnje (EUR)", "Stopa zamjene",
  ];

  const body1 = scenariji.map(s => {
    const tip = s.tip === "stup3" ? "III. stup" : s.stup2 && s.stup1 ? "I.+II. stup" : s.stup1 ? "I. stup" : "—";
    const r1 = s.stup1 ? calcStup1(s.stup1) : null;
    const dod27 = r1 ? r1.efektivna * 0.27 : 0;

    const k2raw = s.stup2 ? calcStup2FV(s.stup2) : 0;
    const jed_pct = s.stup2 ? (s.stup2.jednokratni_pct || 0) : 0;
    const jed_eur = k2raw * jed_pct;
    const preostali2 = k2raw * (1 - jed_pct);
    const renta2 = s.stup2 ? rentaDozivotna(preostali2) : 0;

    const s3 = s.stup3;
    let k3 = 0, dps3 = 0;
    if (s3) {
      const pmt3 = (s3.uplata + (s3.posl || 0) + Math.min(s3.uplata * 12, DPS_MAX_U) * DPS_S / 12);
      const r3 = s3.prinos / 12; const n3 = s3.god * 12;
      k3 = r3 === 0 ? pmt3 * n3 : pmt3 * ((Math.pow(1 + r3, n3) - 1) / r3);
      dps3 = Math.min(s3.uplata * 12, DPS_MAX_U) * DPS_S;
    }

    const uk = (r1 ? r1.efektivna + dod27 : 0) + renta2;
    const ref_placa = s.stup2 ? s.stup2.bruto_placa : (s.stup3 ? 1300 : 1300);

    return [
      s.naziv || "—",
      tip,
      new Date(s.created).toLocaleDateString("hr-HR"),
      // I. stup
      r1 ? xlCell(s.stup1.staz, { fmt: "0.0" }) : xlCell("—"),
      r1 ? xlCell(s.stup1.pvb, { fmt: "0.000" }) : xlCell("—"),
      r1 ? xlCell(s.stup1.pf, { fmt: "0.000" }) : xlCell("—"),
      r1 ? xlCell(s.stup1.mf, { fmt: "0.00" }) : xlCell("—"),
      r1 ? eur(s.stup1.avm || AVM) : xlCell("—"),
      r1 ? xlCell(r1.ob, { fmt: "0.000" }) : xlCell("—"),
      r1 ? eur(r1.bruto) : xlCell("—"),
      r1 ? eur(r1.efektivna) : xlCell("—"),
      r1 ? eur(r1.efektivna * 12) : xlCell("—"),
      r1 ? eur(dod27) : xlCell("—"),
      // II. stup
      s.stup2 ? eur(s.stup2.bruto_placa) : xlCell("—"),
      s.stup2 ? xlCell(s.stup2.god_stednje, { fmt: "0" }) : xlCell("—"),
      s.stup2 ? pctCell(s.stup2.neto_prinos) : xlCell("—"),
      s.stup2 ? eur(k2raw) : xlCell("—"),
      s.stup2 ? pctCell(jed_pct) : xlCell("—"),
      s.stup2 ? eur(jed_eur) : xlCell("—"),
      s.stup2 ? eur(preostali2) : xlCell("—"),
      s.stup2 ? eur(renta2) : xlCell("—"),
      // III. stup
      s3 ? eur(s3.uplata) : xlCell("—"),
      s3 ? xlCell(s3.god, { fmt: "0" }) : xlCell("—"),
      s3 ? pctCell(s3.prinos) : xlCell("—"),
      s3 ? eur(s3.posl || 0) : xlCell("—"),
      s3 ? eur(dps3) : xlCell("—"),
      s3 ? eur(k3) : xlCell("—"),
      // Ukupno
      eur(uk),
      eur(uk * 12),
      pctCell(uk / ref_placa),
    ];
  });

  xlAppendSheet(wb, [
    [`Mirovina HR — Izvoz scenarija | ${username} | ${datum}`],
    [],
    ["⚠ Svi iznosi su informativni izračuni. Za točne podatke koristite mojamirovina.hr (e-Građani)."],
    [],
    hdr1,
    ...body1,
  ], "Scenariji");

  // ── SHEET 2: I. STUP DETALJI ─────────────────────────────────────────────
  const stup1Sc = scenariji.filter(s => s.stup1);
  if (stup1Sc.length > 0) {
    const hdr2 = ["Naziv", "Staž", "PVB", "PF", "MF", "AVM", "Osobni bodovi",
      "Bruto mirovina", "Najniža mirovina", "Najviša mirovina",
      "✅ Efektivna mirovina", "Godišnja mirovina", "Dodatak 27%", "S dodatkom (mj.)"];

    xlAppendSheet(wb, [
      ["I. MIROVINSKI STUP — Detalji scenarija"],
      ["Formula: M = PVB × Staž × PF × MF × AVM | ZOMO | AVM = 14,45 EUR (1.7.2025.)"],
      [],
      hdr2,
      ...stup1Sc.map(s => {
        const r = calcStup1(s.stup1);
        return [
          s.naziv, xlCell(s.stup1.staz, { fmt: "0.0" }), xlCell(s.stup1.pvb, { fmt: "0.000" }),
          xlCell(s.stup1.pf, { fmt: "0.000" }), xlCell(s.stup1.mf, { fmt: "0.00" }),
          eur(s.stup1.avm || AVM),
          xlCell(r.ob, { fmt: "0.000" }),
          eur(r.bruto), eur(r.najniza), eur(r.najvisa),
          eur(r.efektivna), eur(r.efektivna * 12), eur(r.efektivna * 0.27),
          eur(r.efektivna * 1.27),
        ];
      }),
    ], "I. stup");
  }

  // ── SHEET 3: II. STUP DETALJI ────────────────────────────────────────────
  const stup2Sc = scenariji.filter(s => s.stup2);
  if (stup2Sc.length > 0) {
    const hdr3 = ["Naziv", "Bruto plaća", "God. štednje", "Neto prinos",
      "Naknade OMF (HANFA)", "Ukupni kapital", "Čiste uplate", "Investicijski prinos",
      "Jed. isplata %", "Jed. isplata EUR (bruto)",
      "Preostali kapital", "Aprox. renta (mj.)", "Aprox. renta (god.)"];

    xlAppendSheet(wb, [
      ["II. MIROVINSKI STUP — Detalji scenarija"],
      ["Formula: FV = PMT × [((1+r_mj)^n_mj − 1)/r_mj] | ZOMF | Naknade OMF: max 0,24% god. (HANFA 2025.)"],
      [],
      hdr3,
      ...stup2Sc.map(s => {
        const p = s.stup2;
        const k = calcStup2FV(p);
        const uplate = p.bruto_placa * 0.05 * p.god_stednje * 12;
        const jed_pct = p.jednokratni_pct || 0;
        const jed_eur = k * jed_pct;
        const preost = k * (1 - jed_pct);
        const renta = rentaDozivotna(preost);
        return [
          s.naziv, eur(p.bruto_placa), xlCell(p.god_stednje, { fmt: "0" }),
          pctCell(p.neto_prinos), xlCell(0.0024, { fmt: "0.000%" }),
          eur(k), eur(uplate), eur(k - uplate),
          pctCell(jed_pct), eur(jed_eur),
          eur(preost), eur(renta), eur(renta * 12),
        ];
      }),
    ], "II. stup");
  }

  // ── SHEET 4: III. STUP DETALJI ───────────────────────────────────────────
  const stup3Sc = scenariji.filter(s => s.stup3);
  if (stup3Sc.length > 0) {
    const hdr4 = ["Naziv", "Mj. uplata", "Uplata poslodavca", "God. štednje", "Prinos",
      "DPS godišnje (15%)", "DPS ukupno", "Efektivna mj. uplata (s DPS-om)",
      "Ukupni kapital", "Čiste uplate", "Investicijski prinos + DPS",
      "Kapital bez DPS-a (usporedba)"];

    xlAppendSheet(wb, [
      ["III. MIROVINSKI STUP (Dobrovoljna štednja) — Detalji scenarija"],
      [`DPS: 15% od god. uplate do ${DPS_MAX_U} EUR → max ${(DPS_MAX_U * 0.15).toFixed(2)} EUR/god. | ZDOBF | Ministarstvo financija`],
      [],
      hdr4,
      ...stup3Sc.map(s => {
        const p = s.stup3;
        const dps_god = Math.min(p.uplata * 12, DPS_MAX_U) * DPS_S;
        const dps_uk = dps_god * p.god;
        const pmt_eff = (p.uplata + (p.posl || 0) + dps_god / 12);
        const r = p.prinos / 12; const nn = p.god * 12;
        const k_eff = r === 0 ? pmt_eff * nn : pmt_eff * ((Math.pow(1 + r, nn) - 1) / r);
        const k_nodps = r === 0 ? p.uplata * nn : p.uplata * ((Math.pow(1 + r, nn) - 1) / r);
        const uplate_ciste = p.uplata * 12 * p.god;
        return [
          s.naziv, eur(p.uplata), eur(p.posl || 0),
          xlCell(p.god, { fmt: "0" }), pctCell(p.prinos),
          eur(dps_god), eur(dps_uk), eur(pmt_eff),
          eur(k_eff), eur(uplate_ciste), eur(k_eff - uplate_ciste),
          eur(k_nodps),
        ];
      }),
    ], "III. stup");
  }

  // ── SHEET 5: PROJEKCIJA AVM (20 polugodišta) ─────────────────────────────
  const avm_ipc  = 0.03;
  const avm_ibp  = 0.04;
  const stopa_fn = (ipc, ibp) =>
    ipc / (ipc + ibp) <= 0.5
      ? 0.15 * ipc + 0.85 * ibp
      : 0.85 * ipc + 0.15 * ibp;

  const projHdr = ["Polugodište", "Datum", "AVM (EUR)", "Stopa usklađivanja", "ΔAVM (EUR)", "Kum. rast (%)"];
  const projRows = [];
  let avm_cur = AVM;
  const projDates = [
    "01.01.2026","01.07.2026","01.01.2027","01.07.2027","01.01.2028",
    "01.07.2028","01.01.2029","01.07.2029","01.01.2030","01.07.2030",
    "01.01.2031","01.07.2031","01.01.2032","01.07.2032","01.01.2033",
    "01.07.2033","01.01.2034","01.07.2034","01.01.2035","01.07.2035",
  ];
  projDates.forEach((d, i) => {
    const stopa = stopa_fn(avm_ipc / 2, avm_ibp / 2);
    const avm_new = avm_cur * (1 + stopa);
    projRows.push([
      xlCell(i + 1, { fmt: "0" }),
      d,
      eur(avm_new),
      pctCell(stopa),
      eur(avm_new - avm_cur),
      pctCell(avm_new / AVM - 1),
    ]);
    avm_cur = avm_new;
  });

  xlAppendSheet(wb, [
    ["PROJEKCIJA AVM — 20 polugodišta (2026.–2035.)"],
    [`Polazni AVM: ${AVM} EUR (1.7.2025.) | Swiss index: 85% rast plaća + 15% IPC | ZOMO 2025.`],
    [`Pretpostavke: godišnji IPC ${(avm_ipc * 100).toFixed(0)}%, rast plaća ${(avm_ibp * 100).toFixed(0)}%`],
    [],
    projHdr,
    ...projRows,
  ], "Projekcija AVM");

  // ── SHEET 6: FORMULE I PRAVNI IZVORI ─────────────────────────────────────
  xlAppendSheet(wb, [
    ["FORMULE I PRAVNI IZVORI"],
    [`Generirano: ${datum} | Izvor: HZMO, HANFA, mss.gov.hr, ZOMO, ZOMF`],
    [],
    ["Formula", "Matematički izraz", "Pravni izvor", "Napomena"],
    ["Osnovna mirovina I. stupa", "M = PVB × Staž × PF × MF × AVM", "ZOMO čl. 75.", ""],
    ["Osobni bodovi", "OB = PVB × Staž × PF", "ZOMO čl. 76.", ""],
    ["Najniža mirovina", "M_min = 1,06 × AVM × Staž × PF × MF", "ZOMO čl. 80.", "Od 1.7.2025."],
    ["Najviša mirovina", "M_max = Staž × 3,8 × PF × MF × AVM", "Zakon o najvišoj mirovini", ""],
    ["Prijevremena — PF", "PF = 1,00 − 0,002 × broj_mj", "ZOMO čl. 77.", "Trajno smanjenje"],
    ["Odgoda mirovine — PF", "PF = 1,00 + 0,0045 × broj_mj", "ZOMO čl. 77.", "Trajno povećanje"],
    ["Dodatak na mirovinu", "Dod = 0,27 × osnovna mirovina", "NN 156/2023", "Od 1.1.2024."],
    ["Invalidska mirovina — MF (potpuna)", "MF = 1,10", "ZOMO čl. 48.", "Od 1.1.2026."],
    ["Invalidska mirovina — MF (djelomična)", "MF = 0,90", "ZOMO čl. 48.", "Od 1.1.2026."],
    ["AVM usklađivanje", "Swiss index: 85% × ΔIBP + 15% × ΔIPC (ili obrnuto)", "ZOMO 2025, čl. 82.", "2× godišnje"],
    ["II. stup — akumulacija", "FV = PMT × [((1+r_mj)^n_mj − 1) / r_mj]", "ZOMF", "Geometrijska renta"],
    ["II. stup — doprinos", "5% bruto plaće na osobni račun", "Zakon o doprinosima", "Uplaćuje poslodavac"],
    ["II. stup — naknade OMF", "Max 0,24% god. imovine fonda (2025.)", "HANFA propis", "Pada do 0,20% od 2029."],
    ["III. stup — DPS", "DPS = 15% × god. uplate, max 99,54 EUR/god.", "ZDOBF čl. 135.", ""],
    ["III. stup — poslodavac", "Max 804 EUR/god. neoporezivo", "Zakon o porezu na dohodak", "Beneficija za zaposlenika"],
    [],
    ["AVM POVIJEST"],
    ["Datum", "AVM (EUR)", "Promjena"],
    ["1.7.2025.", xlCell(14.45, { fmt: "#,##0.00" }), "+6,48%"],
    ["1.1.2025.", xlCell(13.57, { fmt: "#,##0.00" }), "+3,03%"],
    ["1.7.2024.", xlCell(13.17, { fmt: "#,##0.00" }), ""],
    ["1.1.2024.", xlCell(12.34, { fmt: "#,##0.00" }), ""],
    ["1.1.2023.", xlCell(11.59, { fmt: "#,##0.00" }), ""],
    [],
    ["⚠ NAPOMENA"],
    ["Svi iznosi su informativni izračuni temeljeni na javno dostupnim zakonskim formulama."],
    ["Za točan osobni izračun koristite mojamirovina.hr putem e-Građani ili se obratite HZMO-u."],
  ], "Formule i izvori");

  // ── Širine stupaca ────────────────────────────────────────────────────────
  ["Scenariji", "I. stup", "II. stup", "III. stup"].forEach(name => {
    const ws = wb.Sheets[name];
    if (!ws) return;
    ws["!cols"] = Array(30).fill({ wch: 18 });
    if (ws["!cols"][0]) ws["!cols"][0] = { wch: 26 };
  });

  XLSX.writeFile(wb, `mirovina_${(username || "korisnik").replace(/\s+/g, "_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ─── BLOG DATA (static + AI-enriched) ───────────────────────────────────────
const BLOG_POSTS = [
  {
    id: 1, slug: "kako-funkcionira-drugi-stup",
    title: "Kako zaista funkcionira II. mirovinski stup?",
    excerpt: "Vaših 5% bruto plaće odlazi u obvezni mirovinski fond — ali što se tamo događa i zašto iznosi mirovina mogu biti toliko različiti?",
    category: "Osnove", readMin: 6, date: "2025-09-10",
    tags: ["II. stup", "OMF", "kapitalizacija"],
    content: `## Što je II. mirovinski stup?

Od 2002. godine svaki zaposlenik u Hrvatskoj obvezno uplaćuje **5% bruto plaće** na osobni račun u Obveznom mirovinskom fondu (OMF). Ova sredstva nisu zajednička — vaša su osobna imovina, naslijediva i odvojena od državnog proračuna.

## Kako raste vaš kapital?

OMF ulaže vaše uplate u financijske instrumente (obveznice, dionice, nekretninski fondovi). Prinos nije zajamčen, ali prosječni prinos hr. OMF-ova od 2002. do 2024. bio je **5–7% godišnje**.

Formula akumulacije:
> **FV = PMT × [((1 + r_mj)^n − 1) / r_mj]**

gdje je PMT = mjesečna uplata, r_mj = neto prinos/12, n = broj mjeseci.

**Primjer:** Osoba s prosječnom bruto plaćom 1.300 EUR, koja štedi 33 godine uz 5,5% neto prinos, akumulira **~95.000 EUR**.

## Naknade koje ne vidite

Svake godine OMF naplaćuje naknadu za upravljanje. HANFA propisuje:
- 2025.: **0,24%** godišnje od ukupne imovine fonda
- 2029. i dalje: **0,20%**

Na 33 godine štednje, razlika između 3% i 7% prinosa na kapital od 1.300 EUR/mj. = razlika **~65.000 EUR** u akumuliranom kapitalu.

## Što čuvaju za vas?

Vaša sredstva su **osobna imovina**. Ako umrete, a obitelj ne stječe pravo na obiteljsku mirovinu, akumulirana sredstva su predmet nasljeđivanja prema Zakonu o nasljeđivanju (čl. 103. ZOMF).`
  },
  {
    id: 2, slug: "nacini-isplate-usporedba",
    title: "Načini isplate iz II. stupa: što izabrati?",
    excerpt: "Doživotna renta, zajamčeno razdoblje ili jednokratna isplata? Razlike su značajne — i ovise o vašem zdravlju, obitelji i apetitu za rizik.",
    category: "Isplata", readMin: 8, date: "2025-10-15",
    tags: ["isplata", "MOD", "renta", "nasljeđivanje"],
    content: `## Četiri vrste mirovine iz II. stupa

Kada odete u mirovinu i izaberete kombiniranu isplatu (I. + II. stup), morate odabrati **vrstu mirovine** od MOD-a.

### 1. Pojedinačna mirovina
Isplaćuje se dok ste živi. Prestaje vašom smrću — ostatak kapitala zadržava MOD. Daje **najviši** početni iznos.

### 2. Pojedinačna s zajamčenim razdobljem (5–20 god.)
Ako umrete unutar zajamčenog perioda, mirovina se isplaćuje imenovanom korisniku **do kraja perioda** (min. 50% vaše mirovine). Niži početni iznos.

### 3. Zajednička mirovina
Nastavlja se isplaćivati bračnom drugu koji vas nadživi — **min. 60%** vaše mirovine. Dobar odabir za bračne parove.

### 4. Zajednička s zajamčenim razdobljem
Kombinacija gore navedenog — maksimalna zaštita obitelji, ali najniži početni iznos.

## Jednokratna isplata do 20%

Uvjet: osnovna mirovina iz I. stupa ≥ 115% najniže mirovine iz I. stupa.

Isplaćuje se bruto — MOD obračunava porez na dohodak. Preostali kapital (80%+) ide u doživotnu rentu.

## Koje opcije čuvaju kapital?

Trenutno nijedna hrvatska opcija ne vraća neisplaćeni kapital nasljednicima ako živite dulje od aktuarskog prosjeka. **Ovo je ključna razlika od osiguranja s povratom premije.** Pogledajte naš "Inovativne varijante" kalkulator za hipotetske scenarije.`
  },
  {
    id: 3, slug: "inovativne-varijante-isplate",
    title: "Varijante koje čuvaju kapital: moguće, ali ne postoje u HR",
    excerpt: "Što ako bi MOD mogao ponuditi mirovinu s povratom ostatka kapitala nasljednicima? Izračunajte koliko bi iznosila — i po kojoj cijeni.",
    category: "Inovacije", readMin: 10, date: "2025-11-01",
    tags: ["inovacije", "kapital", "nasljeđivanje", "hibridni model"],
    content: `## Zašto klasična renta 'guta' kapital?

U standardnoj doživotnoj renti, MOD preuzima **mortalitetni rizik**: ako živite kratko, ostatak kapitala ostaje u zajedničkom pulu. Ako živite dugo, MOD nastavlja isplaćivati iz tog pula. Ovo je osiguranje, ne štednja.

## Hipotetska varijanta A: Renta s povratom ostatka

**Opis:** Svaki mjesec dobivate rentu. Kad umrete, nasljednici dobivaju razliku između akumuliranog kapitala i do tada isplaćenih iznosa (ako je pozitivna).

**Matematika:** Početna renta bi bila niža (MOD ne preuzima mortalitetni rizik), ali nasljednici su zaštićeni.

Okvirno smanjenje rente: **15–25%** u odnosu na standardnu doživotnu rentu.

## Hipotetska varijanta B: Fiksna isplata kroz N godina

**Opis:** Isplaćujete se fiksni iznos N godina (npr. 20), neovisno o tome jeste li živi. Nasljednici nastavljaju primati do kraja perioda.

**Rizik:** Ako preživite N godina, isplata prestaje. Rješenje: kombinacija s minimalnom doživotnom rentom.

## Hipotetska varijanta C: Investicijska mirovina

**Opis:** Kapital ostaje uložen. Primate % od prinosa (npr. 4% godišnje = "pravilo 4%"). Glavnica raste ili pada s tržištem.

**Prednost:** Kapital se može nasljediti u cijelosti. **Rizik:** Isplate nisu zajamčene.

## Zašto ovo ne postoji u HR?

1. Regulatorni okvir (ZOMF) zahtijeva od MOD-ova da garantiraju doživotnu isplatu
2. MOD mora imati tehničke pričuve > 100% obveza u svakom trenutku
3. Nema tržišnog pritiska — 4 MOD-a, limitirana konkurencija

Kalkulator "Inovativne varijante" u ovoj aplikaciji simulira sve tri opcije.`
  },
  {
    id: 4, slug: "pvb-sto-je-i-zasto-je-vazan",
    title: "Prosječni vrijednosni bodovi: najvažniji broj vaše mirovine",
    excerpt: "PVB je omjer vaše plaće i prosječne plaće RH — i izravno određuje visinu vaše buduće mirovine. Saznajte kako ga optimizirati.",
    category: "I. stup", readMin: 5, date: "2025-08-20",
    tags: ["PVB", "plaća", "I. stup", "optimizacija"],
    content: `## Što je vrijednosni bod?

Za svaku godinu radnog vijeka HZMO izračunava:

> **VB(god.) = Vaša plaća te godine / Prosječna plaća RH te godine**

Ako ste zarađivali **točno prosječnu plaću**, vaš VB za tu godinu = 1,000.
Ako ste zarađivali **dvostruko više** od prosjeka, VB = 2,000.

## Prosječni vrijednosni bodovi (PVB)

PVB je prosjek svih godišnjih VB-ova kroz karijeru. HZMO ga izračunava za sve godine od 1970. do godine prije umirovljenja.

## Zašto je PVB ključan?

Formula mirovine je: **M = PVB × Staž × PF × MF × AVM**

Osoba s PVB 1,5 dobiva **50% veću mirovinu** od osobe s PVB 1,0 — uz isti staž.

## Optimizacija PVB-a

- Ako ste plaća raste prema kraju karijere, **isplati se raditi dulje** (visoke plaće dižu PVB više od ranijih, nižih plaća)
- Godina umirovljenja **ne ulazi** u izračun VB (uzimaju se plaće do prethodne godine)
- Bonus i otpremnina u godini odlaska ne pomažu PVB-u`
  },
  {
    id: 5, slug: "prijevremena-mirovina-cijena",
    title: "Prijevremena mirovina: koliko košta svaki rani odlazak?",
    excerpt: "Za svaki mjesec ranijeg odlaska u mirovinu, polazni faktor se trajno smanjuje za 0,2%. Na 36 mjeseci to je 7,2% — zauvijek.",
    category: "Strategija", readMin: 7, date: "2025-09-25",
    tags: ["prijevremena mirovina", "polazni faktor", "strategija"],
    content: `## Kazna za rani odlazak

Za svaki **mjesec** ranijeg odlaska u mirovinu (u odnosu na zakonsku dob), polazni faktor se smanjuje za **0,2%**.

> **PF = 1,00 − 0,002 × broj_mj_ranijeg_odlaska**

## Primjeri

| Rani odlazak | PF | Trajno smanjenje mirovine |
|---|---|---|
| 6 mj. | 0,988 | −1,2% |
| 12 mj. | 0,976 | −2,4% |
| 24 mj. | 0,952 | −4,8% |
| 36 mj. | 0,928 | −7,2% |
| 60 mj. | 0,880 | −12,0% |

## Puna mirovinska dob 2025.

Muškarci: **65 godina**
Žene (2025.): **63 god. i 9 mj.** (prijelazno razdoblje do 2030.)

## Odgoda mirovine = nagrada

Suprotno vrijedi za odgodu: **+0,45% po mjesecu** nakon pune dobi.

Za 24 mj. odgode: PF = 1,00 + 0,0045 × 24 = **1,108** — 10,8% veća mirovina zauvijek.

## Od 1.1.2026. — nova zaštita

Korisnicima s PF < 1,0 koji navrše **70 godina**, Zavod po službenoj dužnosti određuje mirovinu bez smanjenja PF-a od sljedećeg mjeseca.`
  }
];

// ─── COMPONENTS ─────────────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html { font-size: 15px; }
  body { font-family: 'Inter', sans-serif; background: #eef0f8; color: #1a2e52; min-height: 100vh; }
  input, select, textarea { font-family: 'Inter', sans-serif; }
  button { cursor: pointer; font-family: 'Inter', sans-serif; }
  h1, h2, h3 { font-family: 'Playfair Display', serif; }
  code, .mono { font-family: 'JetBrains Mono', monospace; font-size: 0.85em; }
  ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-thumb { background: #c0c8e0; border-radius: 3px; }
  .fadeIn { animation: fadeIn 0.35s ease; }
  @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }
`;

function Tooltip2({ text, children }) {
  const [show, setShow] = useState(false);
  return (
    <span style={{ position: "relative", display: "inline-flex", alignItems: "center" }}
      onMouseEnter={() => setShow(true)} onMouseLeave={() => setShow(false)}>
      {children}
      {show && (
        <span style={{
          position: "absolute", bottom: "calc(100% + 6px)", left: "50%",
          transform: "translateX(-50%)", background: C.navy, color: "#fff",
          padding: "6px 10px", borderRadius: 6, fontSize: 12, whiteSpace: "pre-wrap",
          maxWidth: 240, zIndex: 100, boxShadow: "0 4px 16px rgba(0,0,0,0.18)",
          lineHeight: 1.5, pointerEvents: "none",
        }}>{text}</span>
      )}
    </span>
  );
}

function InfoBtn({ text }) {
  return (
    <Tooltip2 text={text}>
      <Info size={13} style={{ color: C.gold, marginLeft: 4, cursor: "help", flexShrink: 0 }} />
    </Tooltip2>
  );
}

function Card({ children, style = {}, className = "" }) {
  return (
    <div className={className} style={{
      background: C.white, borderRadius: 14, padding: 24,
      boxShadow: "0 2px 16px rgba(26,46,82,0.07)", border: `1px solid ${C.border}`,
      ...style
    }}>{children}</div>
  );
}

function Badge({ label, color = C.navy }) {
  return (
    <span style={{
      background: color + "18", color, padding: "2px 9px",
      borderRadius: 20, fontSize: 11, fontWeight: 600, letterSpacing: "0.02em",
    }}>{label}</span>
  );
}

// ─── SHARED SLIDER STYLES injected once ─────────────────────────────────────
const SLIDER_CSS = `
  .rng {
    -webkit-appearance: none; appearance: none;
    width: 100%; height: 6px; border-radius: 3px;
    background: transparent; outline: none; cursor: pointer; margin: 0;
  }
  .rng::-webkit-slider-thumb {
    -webkit-appearance: none; appearance: none;
    width: 20px; height: 20px; border-radius: 50%;
    background: #fff; border: 3px solid var(--tc, #1a2e52);
    box-shadow: 0 1px 6px rgba(0,0,0,0.22);
    cursor: pointer; transition: transform 0.12s, box-shadow 0.12s;
    position: relative; z-index: 2;
  }
  .rng::-moz-range-thumb {
    width: 20px; height: 20px; border-radius: 50%;
    background: #fff; border: 3px solid var(--tc, #1a2e52);
    box-shadow: 0 1px 6px rgba(0,0,0,0.22);
    cursor: pointer; transition: transform 0.12s; z-index: 2;
  }
  .rng:hover::-webkit-slider-thumb { transform: scale(1.18); box-shadow: 0 2px 10px rgba(0,0,0,0.28); }
  .rng:hover::-moz-range-thumb    { transform: scale(1.18); }
  .rng:active::-webkit-slider-thumb { transform: scale(1.08); }
  .num-inp {
    width: 72px; text-align: center; padding: 4px 6px;
    border: 1.5px solid #d1d5db; border-radius: 7px;
    font-size: 12px; font-weight: 700; font-family: 'Inter', sans-serif;
    outline: none; background: #fafbff;
    transition: border-color 0.15s;
  }
  .num-inp:focus { border-color: #1a2e52; }
`;

// Shared track component with coloured fill + visible thumb
function RangeTrack({ value, min, max, step, color, onChange }) {
  const pct = ((value - min) / (max - min)) * 100;
  const trackBg = `linear-gradient(90deg, ${color} ${pct}%, #d1d5db ${pct}%)`;
  return (
    <input
      type="range" min={min} max={max} step={step} value={value}
      className="rng"
      style={{ "--tc": color, background: trackBg }}
      onChange={e => onChange(Number(e.target.value))}
    />
  );
}

// Clamp + parse raw text input back to a numeric value
function parseInput(raw, min, max, step) {
  const n = parseFloat(raw.replace(",", "."));
  if (isNaN(n)) return null;
  const decimals = step < 1 ? String(step).split(".")[1]?.length ?? 0 : 0;
  const clamped = Math.min(max, Math.max(min, n));
  return parseFloat(clamped.toFixed(decimals));
}

// ─── LEAN SLIDER (Ukupno quick-input + Edu MiniSlider base) ─────────────────
function Slider({ label, info, value, min, max, step = 1, format, rawUnit = "", onChange }) {
  const [draft, setDraft] = useState("");
  const [editing, setEditing] = useState(false);
  const color = C.navy;

  function commitDraft() {
    const v = parseInput(draft, min, max, step);
    if (v !== null) onChange(v);
    setEditing(false);
    setDraft("");
  }

  return (
    <div style={{ marginBottom: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <label style={{ fontSize: 13, fontWeight: 500, color: C.slate, display: "flex", alignItems: "center", gap: 2 }}>
          {label}{info && <InfoBtn text={info} />}
        </label>
        <input
          className="num-inp"
          value={editing ? draft : (format ? format(value) : String(value))}
          onFocus={() => { setEditing(true); setDraft(String(value)); }}
          onChange={e => setDraft(e.target.value)}
          onBlur={commitDraft}
          onKeyDown={e => { if (e.key === "Enter") commitDraft(); if (e.key === "Escape") { setEditing(false); setDraft(""); } }}
          style={{ borderColor: editing ? color : C.border }}
        />
      </div>
      <RangeTrack value={value} min={min} max={max} step={step} color={color} onChange={onChange} />
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: C.slateL, marginTop: 4 }}>
        <span>{format ? format(min) : min}</span><span>{format ? format(max) : max}</span>
      </div>
    </div>
  );
}

// ─── MINI SLIDER (Edukacija panel sliders) ───────────────────────────────────
function MiniSlider({ label, value, min, max, step, format, onChange, color = C.navy }) {
  const [draft, setDraft] = useState("");
  const [editing, setEditing] = useState(false);

  function commitDraft() {
    const v = parseInput(draft, min, max, step);
    if (v !== null) onChange(v);
    setEditing(false); setDraft("");
  }

  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <span style={{ fontSize: 12, color: C.slate, fontWeight: 500 }}>{label}</span>
        <input
          className="num-inp"
          value={editing ? draft : (format ? format(value) : String(value))}
          onFocus={() => { setEditing(true); setDraft(String(value)); }}
          onChange={e => setDraft(e.target.value)}
          onBlur={commitDraft}
          onKeyDown={e => { if (e.key === "Enter") commitDraft(); if (e.key === "Escape") { setEditing(false); setDraft(""); } }}
          style={{ borderColor: editing ? color : C.border }}
        />
      </div>
      <RangeTrack value={value} min={min} max={max} step={step} color={color} onChange={onChange} />
    </div>
  );
}

// ─── SMART PARAM – rich human-friendly parameter block ───────────────────────
function SmartParam({ id, icon, title, value, min, max, step = 1, format,
                      onChange, explain, dynamicExplain, zones, children }) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState("");
  const [editing, setEditing] = useState(false);

  let zoneColor = C.navy;
  let zoneLabel = "";
  if (zones) {
    for (const z of zones) {
      if (value <= z.upTo) { zoneColor = z.color; zoneLabel = z.label; break; }
    }
  }

  const displayVal = format ? format(value) : String(value);
  const dynText = dynamicExplain ? dynamicExplain(value) : null;

  function commitDraft() {
    const v = parseInput(draft, min, max, step);
    if (v !== null) onChange(v);
    setEditing(false); setDraft("");
  }

  return (
    <div style={{ marginBottom: 14, borderRadius: 12, border: `1.5px solid ${zoneColor}33`,
      background: `${zoneColor}06`, overflow: "hidden" }}>

      {/* Header: icon + title + zone badge + direct input */}
      <div style={{ padding: "12px 14px 10px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 7, flex: 1 }}>
            <span style={{ fontSize: 18, flexShrink: 0 }}>{icon}</span>
            <div>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.navy }}>{title}</div>
              {zoneLabel && (
                <span style={{ fontSize: 10, fontWeight: 700, color: zoneColor,
                  background: `${zoneColor}18`, padding: "1px 7px", borderRadius: 10 }}>
                  {zoneLabel}
                </span>
              )}
            </div>
          </div>

          {/* Editable value badge */}
          <input
            className="num-inp"
            value={editing ? draft : displayVal}
            onFocus={() => { setEditing(true); setDraft(String(value)); }}
            onChange={e => setDraft(e.target.value)}
            onBlur={commitDraft}
            onKeyDown={e => { if (e.key === "Enter") commitDraft(); if (e.key === "Escape") { setEditing(false); setDraft(""); } }}
            style={{
              width: 88, flexShrink: 0, marginLeft: 10,
              fontSize: 13, fontWeight: 800, color: editing ? C.navy : zoneColor,
              borderColor: editing ? zoneColor : `${zoneColor}55`,
              background: editing ? "#fff" : `${zoneColor}10`,
            }}
            title="Kliknite za izravni unos vrijednosti"
          />
        </div>

        {/* Slider track with visible thumb */}
        <RangeTrack value={value} min={min} max={max} step={step} color={zoneColor} onChange={onChange} />
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: C.slateL, marginTop: 4 }}>
          <span>{format ? format(min) : min}</span><span>{format ? format(max) : max}</span>
        </div>

        {/* Dynamic contextual sentence */}
        {dynText && (
          <div style={{ marginTop: 10, fontSize: 12, color: C.slate, lineHeight: 1.5,
            padding: "7px 11px", background: `${zoneColor}10`, borderRadius: 7,
            borderLeft: `3px solid ${zoneColor}` }}>
            {dynText}
          </div>
        )}
      </div>

      {/* Expandable explanation */}
      {explain && (
        <>
          <button onClick={() => setOpen(v => !v)} style={{
            width: "100%", background: "none", border: "none", borderTop: `1px solid ${C.border}`,
            padding: "7px 14px", display: "flex", justifyContent: "space-between",
            alignItems: "center", cursor: "pointer", color: C.slateL, fontSize: 11,
          }}>
            <span>ℹ Što je ovo i kako utječe na mirovinu?</span>
            {open ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
          </button>
          {open && (
            <div style={{ padding: "12px 14px", background: C.white,
              borderTop: `1px solid ${C.border}`, fontSize: 12, color: C.slate, lineHeight: 1.7 }}>
              {explain}
            </div>
          )}
        </>
      )}

      {children && <div style={{ padding: "0 14px 12px" }}>{children}</div>}
    </div>
  );
}

function ResultBox({ label, value, sub, color = C.navy, big = false }) {
  return (
    <div style={{
      background: color === C.green ? C.greenL : color === C.gold ? C.amberL : C.blueL,
      border: `2px solid ${color}22`, borderRadius: 12, padding: "14px 18px",
      textAlign: "center"
    }}>
      <div style={{ fontSize: 11, color: C.slate, fontWeight: 500, marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: big ? 26 : 20, fontWeight: 800, color, fontFamily: "'Playfair Display', serif" }}>
        {value}
      </div>
      {sub && <div style={{ fontSize: 11, color: C.slateL, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

// ─── AUTH MODAL (overlay, triggered on demand) ───────────────────────────────
function AuthModal({ onLogin, onClose, reason = "" }) {
  const [mode, setMode] = useState("login");
  const [form, setForm] = useState({ name: "", email: "", password: "" });
  const [err, setErr] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [loading, setLoading] = useState(false);

  async function handleSubmit() {
    setErr(""); setLoading(true);
    await new Promise(r => setTimeout(r, 350));
    if (mode === "register") {
      if (!form.name || !form.email || !form.password) { setErr("Sva polja su obavezna."); setLoading(false); return; }
      if (form.password.length < 6) { setErr("Lozinka mora imati najmanje 6 znakova."); setLoading(false); return; }
      const existing = await stGet(`user:${form.email}`);
      if (existing) { setErr("Račun s tom e-adresom već postoji."); setLoading(false); return; }
      const user = { name: form.name, email: form.email, password: form.password, created: Date.now() };
      await stSet(`user:${form.email}`, user);
      onLogin(user);
    } else {
      if (!form.email || !form.password) { setErr("Unesite e-adresu i lozinku."); setLoading(false); return; }
      const user = await stGet(`user:${form.email}`);
      if (!user || user.password !== form.password) { setErr("Pogrešna e-adresa ili lozinka."); setLoading(false); return; }
      onLogin(user);
    }
    setLoading(false);
  }

  const inp = { width: "100%", padding: "11px 14px", border: `1.5px solid ${C.border}`,
    borderRadius: 9, fontSize: 14, outline: "none", background: "#fafbff" };

  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 200,
      background: "rgba(15,25,50,0.55)", backdropFilter: "blur(4px)",
      display: "flex", alignItems: "center", justifyContent: "center", padding: 20,
    }} onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="fadeIn" style={{ width: "100%", maxWidth: 400, background: C.white, borderRadius: 18, padding: 32, boxShadow: "0 24px 80px rgba(0,0,0,0.3)" }}>
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <div style={{ fontSize: 24, marginBottom: 4 }}>🏛</div>
            <h2 style={{ fontSize: 20, marginBottom: 4 }}>Prijava / Registracija</h2>
            {reason && <p style={{ fontSize: 12, color: C.slateL, lineHeight: 1.5 }}>{reason}</p>}
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: C.slateL, padding: 4, cursor: "pointer" }}>
            <X size={20} />
          </button>
        </div>

        {/* Mode toggle */}
        <div style={{ display: "flex", marginBottom: 20, background: C.bg, borderRadius: 9, padding: 3 }}>
          {[["login","Prijava"],["register","Registracija"]].map(([m, l]) => (
            <button key={m} onClick={() => { setMode(m); setErr(""); }}
              style={{ flex: 1, padding: "8px 0", border: "none", borderRadius: 7,
                background: mode === m ? C.navy : "transparent",
                color: mode === m ? "#fff" : C.slateL, fontWeight: 600, fontSize: 13,
                transition: "all 0.2s", cursor: "pointer" }}>
              {l}
            </button>
          ))}
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 11 }}>
          {mode === "register" && (
            <input style={inp} placeholder="Ime i prezime" value={form.name}
              onChange={e => setForm(f => ({ ...f, name: e.target.value }))} />
          )}
          <input style={inp} placeholder="E-adresa" type="email" value={form.email}
            onChange={e => setForm(f => ({ ...f, email: e.target.value }))} />
          <div style={{ position: "relative" }}>
            <input style={{ ...inp, paddingRight: 42 }} placeholder="Lozinka"
              type={showPw ? "text" : "password"} value={form.password}
              onChange={e => setForm(f => ({ ...f, password: e.target.value }))}
              onKeyDown={e => e.key === "Enter" && handleSubmit()} />
            <button onClick={() => setShowPw(v => !v)} style={{
              position: "absolute", right: 11, top: "50%", transform: "translateY(-50%)",
              background: "none", border: "none", color: C.slateL, cursor: "pointer"
            }}>
              {showPw ? <EyeOff size={15} /> : <Eye size={15} />}
            </button>
          </div>
        </div>

        {err && <div style={{ marginTop: 10, color: C.red, fontSize: 12, padding: "7px 11px", background: C.redL, borderRadius: 7 }}>⚠ {err}</div>}

        <button onClick={handleSubmit} disabled={loading} style={{
          marginTop: 18, width: "100%", padding: "12px 0",
          background: loading ? C.slateL : `linear-gradient(135deg, ${C.gold}, ${C.goldL})`,
          border: "none", borderRadius: 10, color: C.navy, fontWeight: 700, fontSize: 14,
          boxShadow: "0 4px 14px rgba(201,146,43,0.3)", cursor: "pointer"
        }}>
          {loading ? "…" : mode === "login" ? "Prijavi se" : "Kreiraj račun"}
        </button>

        <p style={{ textAlign: "center", marginTop: 14, fontSize: 11, color: C.slateL }}>
          Demo: <span className="mono">testuser@hr.com</span> / <span className="mono">test123</span>
        </p>
      </div>
    </div>
  );
}

// ─── STUP 1 KALKULATOR ───────────────────────────────────────────────────────
function Stup1Calculator({ onSave, user, onAuthRequired }) {
  const [params, setParams] = useState({
    staz: 40, pvb: 1.05, vrsta: "starosna",
    rani_odlazak_mj: 0, odgoda_mj: 0, mf: 1.0, avm: AVM,
    spol: "M", dob_umirovljenja: 65,
  });
  const [naziv, setNaziv] = useState("Moj scenarij");
  const [saved, setSaved] = useState(false);

  const pf = params.vrsta === "prijevremena"
    ? Math.max(0.7, 1 - 0.002 * params.rani_odlazak_mj)
    : params.vrsta === "starosna" && params.odgoda_mj > 0
      ? 1 + 0.0045 * params.odgoda_mj
      : 1.0;

  const mf = params.vrsta === "starosna" ? 1.0
    : params.vrsta === "prijevremena" ? 1.0
    : params.vrsta === "invalidska_potpuna" ? 1.0
    : params.vrsta === "invalidska_djelomicna" ? 0.8 : 0.77;

  const res = calcStup1({ pvb: params.pvb, staz: params.staz, pf, mf, avm: params.avm });

  const cmp_data = [
    { name: "Prijevremena (-36 mj.)", value: calcStup1({ pvb: params.pvb, staz: Math.max(1, params.staz - 3), pf: 1 - 0.002 * 36, mf: 1.0, avm: params.avm }).efektivna },
    { name: "Standardna (puna dob)", value: calcStup1({ pvb: params.pvb, staz: params.staz, pf: 1.0, mf: 1.0, avm: params.avm }).efektivna },
    { name: "Vaš scenarij", value: res.efektivna },
    { name: "Odgoda (+24 mj.)", value: calcStup1({ pvb: params.pvb, staz: params.staz + 2, pf: 1 + 0.0045 * 24, mf: 1.0, avm: params.avm }).efektivna },
    { name: "Odgoda (+48 mj.)", value: calcStup1({ pvb: params.pvb, staz: params.staz + 4, pf: 1 + 0.0045 * 48, mf: 1.0, avm: params.avm }).efektivna },
  ];

  async function handleSave() {
    if (!user) { onAuthRequired("Za spremanje scenarija potrebna je prijava ili registracija. Besplatno je i traje 30 sekundi."); return; }
    const scenario = {
      id: Date.now(), naziv, created: Date.now(),
      stup1: { pvb: params.pvb, staz: params.staz, pf, mf, avm: params.avm },
    };
    const existing = (await stGet(`scenarios:${user.email}`)) || [];
    await stSet(`scenarios:${user.email}`, [...existing, scenario]);
    setSaved(true); setTimeout(() => setSaved(false), 2500);
    if (onSave) onSave();
  }

  const vrsteOptions = [
    { v: "starosna", l: "Starosna mirovina" },
    { v: "prijevremena", l: "Prijevremena starosna" },
    { v: "dugogodisnji", l: "Dugogodišnji osiguranik (60+41 god.)" },
    { v: "invalidska_potpuna", l: "Invalidska (potpuni gubitak)" },
    { v: "invalidska_djelomicna", l: "Invalidska (djelomični gubitak)" },
  ];

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>I. mirovinski stup</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>
          Generacijska solidarnost — sadašnji radnici financiraju sadašnje umirovljenike
        </p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        {/* Left: rich human-friendly params */}
        <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>

          {/* Vrsta mirovine – special select block */}
          <div style={{ marginBottom: 14, borderRadius: 12, border: `1.5px solid ${C.border}`,
            background: C.white, overflow: "hidden" }}>
            <div style={{ padding: "12px 14px 10px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 10 }}>
                <span style={{ fontSize: 18 }}>🎯</span>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.navy }}>Vrsta mirovine</div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                {[
                  { v: "starosna", emoji: "👴", l: "Starosna", sub: "Normalan odlazak u mirovinu po navršenju zakonske dobi" },
                  { v: "prijevremena", emoji: "🏃", l: "Prijevremena", sub: "Odlazak ranije — mirovina trajno manja" },
                  { v: "dugogodisnji", emoji: "⏳", l: "Dugogodišnji", sub: "60 god. + 41 god. staža — posebni uvjeti" },
                  { v: "invalidska_potpuna", emoji: "⚕️", l: "Invalidska", sub: "Potpuni gubitak radne sposobnosti" },
                ].map(o => (
                  <button key={o.v} onClick={() => setParams(p => ({ ...p, vrsta: o.v }))}
                    style={{ padding: "8px 10px", border: `2px solid ${params.vrsta === o.v ? C.navy : C.border}`,
                      borderRadius: 9, background: params.vrsta === o.v ? C.navy : C.white,
                      color: params.vrsta === o.v ? "#fff" : C.navy,
                      cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>
                    <div style={{ fontSize: 16, marginBottom: 2 }}>{o.emoji}</div>
                    <div style={{ fontSize: 12, fontWeight: 700 }}>{o.l}</div>
                    <div style={{ fontSize: 10, opacity: 0.75, lineHeight: 1.3, marginTop: 2 }}>{o.sub}</div>
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* STAŽ */}
          <SmartParam
            icon="📅" title="Ukupni mirovinski staž"
            value={params.staz} min={1} max={50} step={0.5}
            format={v => `${v} god.`}
            onChange={v => setParams(p => ({ ...p, staz: v }))}
            zones={[
              { upTo: 14, color: C.red,   label: "⚠ Premalo za mirovinu (min. 15)" },
              { upTo: 24, color: C.amber, label: "Kratki staž" },
              { upTo: 34, color: C.gold,  label: "Srednji staž" },
              { upTo: 42, color: C.green, label: "Dobar staž" },
              { upTo: 50, color: C.blue,  label: "Dugogodišnji radnik" },
            ]}
            dynamicExplain={v => {
              const yrs = Math.round(v);
              if (v < 15) return `⚠ S ${v} god. staža nema prava na starosnu mirovinu (minimum je 15 godina).`;
              const est = (1.05 * v * 1.0 * 1.0 * AVM).toFixed(0);
              return `S ${v} godina staža i prosječnom plaćom, vaša okvirna mirovina bila bi oko ${est} EUR/mj. — svaka dodatna godina staža povećava je za otprilike ${(1.05 * AVM).toFixed(0)} EUR/mj.`;
            }}
            explain={
              <div>
                <p><strong>Staž</strong> je ukupan broj godina koje ste proveli radeći i uplaćujući doprinose u mirovinski sustav.</p>
                <p style={{ marginTop: 6 }}>Svaka godina staža direktno povećava mirovinu — u formuli je <em>M = PVB × <strong>Staž</strong> × PF × MF × AVM</em>, dakle što je staž veći, to je mirovina veća, linearno.</p>
                <p style={{ marginTop: 6 }}>Vrijedi i vojni rok, bolovanje i porodiljni dopust. Staž se zbrajaju iz svih radnih odnosa u životu.</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.blueL, borderRadius: 6 }}>
                  💡 <strong>Praktično:</strong> 5 dodatnih godina rada = oko 5 × {(1.05 * AVM).toFixed(0)} EUR više mirovine svaki mjesec, zauvijek.
                </div>
              </div>
            }
          />

          {/* PVB */}
          <SmartParam
            icon="💰" title="Prosječni vrijednosni bodovi (PVB)"
            value={params.pvb} min={0.3} max={3.8} step={0.01}
            format={v => v.toFixed(2)}
            onChange={v => setParams(p => ({ ...p, pvb: v }))}
            zones={[
              { upTo: 0.79, color: C.red,    label: "Ispod minimuma" },
              { upTo: 0.99, color: C.amber,  label: "Ispod prosjeka" },
              { upTo: 1.09, color: C.green,  label: "Prosječna plaća" },
              { upTo: 1.49, color: C.blue,   label: "Iznad prosjeka" },
              { upTo: 2.49, color: C.purple, label: "Visoka plaća" },
              { upTo: 3.8,  color: "#c00",   label: "Vrhunska plaća" },
            ]}
            dynamicExplain={v => {
              const avgHR = 1300; // prosj. bruto HR 2024
              const approxPlaca = Math.round(v * avgHR);
              if (v < 1.0) return `PVB ${v.toFixed(2)} znači da ste prosječno zarađivali oko ${approxPlaca} EUR bruto — ispod prosjeka RH (≈${avgHR} EUR). Svaki bod manje = manja mirovina razmjerno.`;
              if (v < 1.15) return `PVB ${v.toFixed(2)} ≈ prosječna plaća u RH (≈${approxPlaca} EUR bruto). Mirovina "pratila" je prosjek ekonomije.`;
              return `PVB ${v.toFixed(2)} znači da ste prosječno zarađivali oko ${approxPlaca} EUR bruto — ${Math.round((v-1)*100)}% iznad prosjeka RH. Mirovina raste proporcionalno.`;
            }}
            explain={
              <div>
                <p><strong>PVB (Prosječni vrijednosni bodovi)</strong> — ovo je najvažniji broj koji opisuje koliko ste zarađivali kroz cijelu karijeru u usporedbi s ostatkom Hrvatske.</p>
                <ul style={{ marginTop: 8, paddingLeft: 16, display: "flex", flexDirection: "column", gap: 4 }}>
                  <li><strong>PVB = 1,0</strong> → uvijek ste zarađivali točno prosječnu plaću RH</li>
                  <li><strong>PVB = 1,5</strong> → zarađivali ste 50% više od prosjeka</li>
                  <li><strong>PVB = 0,8</strong> → zarađivali ste 20% manje od prosjeka</li>
                  <li><strong>PVB = 2,0</strong> → dvostruko više od prosjeka (visoke plaće)</li>
                </ul>
                <p style={{ marginTop: 8 }}>HZMO svake godine dijeli vašu plaću s prosječnom plaćom RH — to je godišnji vrijednosni bod. PVB je prosjek svih tih godišnjih bodova.</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.amberL, borderRadius: 6 }}>
                  💡 <strong>Primjer:</strong> Zarađivali ste 1.950 EUR bruto, a prosjek RH bio 1.300 EUR → bod te godine = 1.500. Ako je to bio prosjek vaše karijere, PVB = 1,50.
                </div>
              </div>
            }
          />

          {/* Prijevremeni odlazak */}
          {params.vrsta === "prijevremena" && (
            <SmartParam
              icon="⏩" title="Koliko ranije odlazite u mirovinu?"
              value={params.rani_odlazak_mj} min={1} max={60} step={1}
              format={v => v < 12 ? `${v} mj.` : `${Math.floor(v/12)} god. ${v%12 ? `${v%12} mj.` : ""}`}
              onChange={v => setParams(p => ({ ...p, rani_odlazak_mj: v }))}
              zones={[
                { upTo: 6,  color: C.green, label: "Malo ranije (mala kazna)" },
                { upTo: 24, color: C.amber, label: "Znatno ranije" },
                { upTo: 48, color: C.red,   label: "Dosta ranije (velika kazna)" },
                { upTo: 60, color: "#900",  label: "Maksimalna kazna" },
              ]}
              dynamicExplain={v => {
                const pf_calc = Math.max(0.7, 1 - 0.002 * v);
                const kazna = ((1 - pf_calc) * 100).toFixed(1);
                return `Odlaskom ${v} mj. ranije, vaša mirovina je trajno manja za ${kazna}% — svaki mjesec do kraja života. Polazni faktor PF = ${pf_calc.toFixed(3)}.`;
              }}
              explain={
                <div>
                  <p>Za svaki <strong>mjesec ranijeg odlaska</strong> u mirovinu, mirovina se trajno smanjuje za <strong>0,2%</strong>.</p>
                  <p style={{ marginTop: 6 }}>Ovo smanjenje je <em>permanentno</em> — ne nadoknađuje se nikad. Formula: <em>PF = 1,00 − 0,002 × broj_mj</em>.</p>
                  <div style={{ marginTop: 8, padding: "6px 10px", background: C.redL, borderRadius: 6 }}>
                    ⚠ <strong>Pažnja:</strong> 3 godine ranije = −7,2% mirovine zauvijek. Na 20 god. mirovine, to je ekvivalent 1,4 god. bez prihoda.
                  </div>
                </div>
              }
            />
          )}

          {/* Odgoda */}
          {params.vrsta === "starosna" && (
            <SmartParam
              icon="⏳" title="Odgoda iznad zakonske dobi"
              value={params.odgoda_mj} min={0} max={60} step={1}
              format={v => v === 0 ? "Ne odgađam" : v < 12 ? `${v} mj.` : `${Math.floor(v/12)} god. ${v%12 ? `${v%12} mj.` : ""}`}
              onChange={v => setParams(p => ({ ...p, odgoda_mj: v }))}
              zones={[
                { upTo: 0,  color: C.slate, label: "Bez odgode" },
                { upTo: 12, color: C.blue,  label: "Mala odgoda" },
                { upTo: 36, color: C.green, label: "Dobra odgoda" },
                { upTo: 60, color: C.purple,label: "Velika odgoda (+bonus)" },
              ]}
              dynamicExplain={v => {
                if (v === 0) return "Odlaskom točno na zakonsku dob, PF = 1,000 — bez smanjenja ni povećanja.";
                const bonus = (v * 0.0045 * 100).toFixed(1);
                return `Čekate još ${v} mj. dulje → mirovina trajno veća za +${bonus}%. PF = ${(1 + 0.0045*v).toFixed(3)}.`;
              }}
              explain={
                <div>
                  <p>Ako ostanete raditi <strong>dulje od zakonske mirovinske dobi</strong>, svaki dodatni mjesec rada povećava mirovinu za <strong>0,45%</strong>.</p>
                  <p style={{ marginTop: 6 }}>Formula: <em>PF = 1,00 + 0,0045 × broj_mj</em>.</p>
                  <div style={{ marginTop: 8, padding: "6px 10px", background: C.greenL, borderRadius: 6 }}>
                    💡 <strong>Primjer:</strong> Radite još 24 mj. = PF 1,108 = trajno +10,8% mirovine. Isplati se ako ste zdravi i možete raditi.
                  </div>
                </div>
              }
            />
          )}

          {/* AVM */}
          <SmartParam
            icon="📊" title="Aktualna vrijednost mirovine (AVM)"
            value={params.avm} min={10} max={20} step={0.01}
            format={v => `${v.toFixed(2)} EUR`}
            onChange={v => setParams(p => ({ ...p, avm: v }))}
            zones={[
              { upTo: 13,   color: C.red,   label: "Ispod sadašnje razine" },
              { upTo: 14.45,color: C.green, label: "Sadašnja razina (1.7.2025.)" },
              { upTo: 17,   color: C.blue,  label: "Projicirana buduća razina" },
              { upTo: 20,   color: C.purple,label: "Optimistična projekcija" },
            ]}
            dynamicExplain={v => {
              if (Math.abs(v - AVM) < 0.02) return `AVM ${v.toFixed(2)} EUR — ovo je sadašnja vrijednost (od 1.7.2025.). Svaki vaš osobni bod vrijedi ${v.toFixed(2)} EUR mirovine.`;
              const razlika = ((v - AVM) / AVM * 100).toFixed(1);
              return `AVM ${v.toFixed(2)} EUR = ${razlika > 0 ? '+' : ''}${razlika}% vs. sadašnja vrijednost. Ovim simulirate buduće usklađivanje AVM.`;
            }}
            explain={
              <div>
                <p><strong>AVM (Aktualna vrijednost mirovine)</strong> je "cijena jednog boda". Određuje je Vlada RH i mijenja se dva puta godišnje sukladno indeksima plaća i cijena.</p>
                <p style={{ marginTop: 6 }}>Vaša mirovina = <em>Osobni bodovi × AVM</em>. Ako AVM poraste za 5%, vaša mirovina poraste za 5%.</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.blueL, borderRadius: 6 }}>
                  💡 Sadašnja vrijednost: <strong>14,45 EUR/bod</strong> (od 1.7.2025.). Promijenite ovaj slider da vidite kako bi mirovina izgledala u budućnosti s drugačijim AVM-om.
                </div>
              </div>
            }
          />

          {/* PF/MF read-only summary */}
          <div style={{ padding: "12px 14px", background: C.navy, borderRadius: 12, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.55)", marginBottom: 3 }}>Polazni faktor (PF)</div>
              <div style={{ fontSize: 22, fontWeight: 800, color: pf >= 1 ? "#7bffb0" : pf >= 0.93 ? C.goldL : "#ff9090",
                fontFamily: "'Playfair Display', serif" }}>{pf.toFixed(3)}</div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.5)", marginTop: 2 }}>
                {pf === 1.0 ? "Bez promjene" : pf > 1 ? `+${((pf-1)*100).toFixed(1)}% (odgoda)` : `−${((1-pf)*100).toFixed(1)}% (rani odlazak)`}
              </div>
            </div>
            <div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.55)", marginBottom: 3 }}>Mirovinski faktor (MF)</div>
              <div style={{ fontSize: 22, fontWeight: 800, color: mf === 1.0 ? "#7bffb0" : C.goldL,
                fontFamily: "'Playfair Display', serif" }}>{mf.toFixed(2)}</div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.5)", marginTop: 2 }}>
                {mf === 1.0 ? "Puna mirovina" : `Smanjena (invalidska/obiteljska)`}
              </div>
            </div>
          </div>
        </div>

        {/* Right: results */}
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <Card style={{ border: `2px solid ${C.navy}22` }}>
            <h3 style={{ fontSize: 14, marginBottom: 14, color: C.navy }}>Rezultati</h3>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <ResultBox label="Osobni bodovi" value={res.ob.toFixed(3)} color={C.blue} />
              <ResultBox label="Bruto mirovina" value={`${fmt(res.bruto)} EUR`} color={C.blue} />
              <ResultBox label="Najniža mirovina" value={`${fmt(res.najniza)} EUR`} color={C.slate} />
              <ResultBox label="Najviša mirovina" value={`${fmt(res.najvisa)} EUR`} color={C.slate} />
            </div>
            <div style={{ marginTop: 12 }}>
              <ResultBox label="✅ Efektivna mirovina / mj." value={`${fmt(res.efektivna)} EUR`} color={C.green} big />
            </div>
            <div style={{ marginTop: 10, textAlign: "center", fontSize: 12, color: C.slateL }}>
              Godišnje: <strong>{fmtEur(res.efektivna * 12)}</strong>
            </div>
          </Card>

          <Card>
            <h3 style={{ fontSize: 13, marginBottom: 12, color: C.navy }}>Usporedba scenarija (isti PVB)</h3>
            <ResponsiveContainer width="100%" height={170}>
              <BarChart data={cmp_data} margin={{ top: 5, right: 10, bottom: 25, left: 0 }}>
                <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-12} textAnchor="end" />
                <YAxis tick={{ fontSize: 10 }} tickFormatter={v => `${Math.round(v)}`} />
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <Tooltip formatter={v => [`${fmt(v)} EUR`, "Mirovina"]} />
                <Bar dataKey="value" radius={[4,4,0,0]}>
                  {cmp_data.map((d, i) => <Cell key={i} fill={d.name === "Vaš scenarij" ? C.gold : (i < 2 ? C.slate + "88" : C.navy + "bb")} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </Card>

          {/* Save */}
          <Card style={{ padding: "14px 18px" }}>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              <input value={naziv} onChange={e => setNaziv(e.target.value)}
                style={{ flex: 1, padding: "8px 12px", border: `1.5px solid ${C.border}`, borderRadius: 8, fontSize: 13 }}
                placeholder="Naziv scenarija" />
              <button onClick={handleSave} style={{
                padding: "8px 18px", background: saved ? C.green : C.navy, color: "#fff",
                border: "none", borderRadius: 8, fontWeight: 600, fontSize: 13, display: "flex",
                alignItems: "center", gap: 6, transition: "background 0.2s"
              }}>
                {saved ? <CheckCircle size={15} /> : <Save size={15} />}
                {saved ? "Spremljeno!" : "Spremi"}
              </button>
            </div>
          </Card>
        </div>
      </div>

      {/* Formula box */}
      <Card style={{ marginTop: 20, background: C.navy, color: "#fff" }}>
        <h3 style={{ fontSize: 14, marginBottom: 12, color: C.goldL }}>📐 Formule (ZOMO) — s vašim vrijednostima</h3>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
          {[
            ["Osobni bodovi", `OB = ${params.pvb.toFixed(2)} × ${params.staz} × ${pf.toFixed(3)} = ${res.ob.toFixed(3)}`],
            ["Bruto mirovina", `${res.ob.toFixed(3)} × ${mf} × ${params.avm} = ${fmt(res.bruto)} EUR`],
            ["Najniža mirovina", `${MIN_MIROVINA_FAKTOR} × ${params.avm} × ${params.staz} × ${pf.toFixed(3)} = ${fmt(res.najniza)} EUR`],
          ].map(([k, v]) => (
            <div key={k} style={{ background: "rgba(255,255,255,0.08)", borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ fontSize: 11, color: C.goldL, marginBottom: 4 }}>{k}</div>
              <div className="mono" style={{ fontSize: 12, color: "#e0e8ff", wordBreak: "break-all" }}>{v}</div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

// ─── STUP 2 KALKULATOR ───────────────────────────────────────────────────────
function Stup2Calculator({ onSave, user, onAuthRequired }) {
  const [params, setParams] = useState({
    bruto_placa: 1300, god_stednje: 33,
    neto_prinos: 0.055, e65: 18,
    jednokratni_pct: 0.20,
    vrsta_rente: "pojedinacna",
    inovativna_varijanta: "standard",
    povrat_kapital_pct: 0.5,
  });
  const [naziv, setNaziv] = useState("Moj II. stup scenarij");
  const [saved, setSaved] = useState(false);

  const kapital = calcStup2FV(params);
  const jednokratni_iznos = kapital * params.jednokratni_pct;
  const preostali = kapital * (1 - params.jednokratni_pct);
  const renta_standard = rentaDozivotna(preostali, params.e65);

  // Inovativne varijante
  const renta_s_povratom = preostali / (params.e65 * 12) * (1 - 0.18); // 18% niža zbog garancije povrata
  const renta_fiksno_n = preostali / (20 * 12); // 20 godina fiksno
  const renta_investicijska = (preostali * 0.04) / 12; // 4% pravilo

  const renta_prikaz = params.inovativna_varijanta === "standard" ? renta_standard
    : params.inovativna_varijanta === "s_povratom" ? renta_s_povratom
    : params.inovativna_varijanta === "fiksno_n" ? renta_fiksno_n
    : renta_investicijska;

  // Projekcija akumulacije
  const proj_data = Array.from({ length: Math.ceil(params.god_stednje / 5) + 1 }, (_, i) => {
    const god = i * 5;
    if (god > params.god_stednje) return null;
    const k = calcStup2FV({ ...params, god_stednje: god });
    const uplata = params.bruto_placa * 0.05 * god * 12;
    return { god: `${god}g`, kapital: Math.round(k), uplate: Math.round(uplata), prinos: Math.round(k - uplata) };
  }).filter(Boolean);

  // Varijante usporedba
  const varijante_data = [
    { name: "Standard\n(doživotna)", renta: renta_standard, nasljeđuje: "Ne" },
    { name: "S povratom\nkapitala*", renta: renta_s_povratom, nasljeđuje: "Da" },
    { name: "Fiksno 20g.*", renta: renta_fiksno_n, nasljeđuje: "Djelomično" },
    { name: "Investicijska\n(4% pravilo)*", renta: renta_investicijska, nasljeđuje: "Da" },
  ];

  async function handleSave() {
    if (!user) { onAuthRequired("Za spremanje scenarija potrebna je prijava ili registracija. Besplatno je i traje 30 sekundi."); return; }
    const scenario = {
      id: Date.now(), naziv, created: Date.now(),
      stup2: { ...params },
      stup1: { pvb: 1.05, staz: 40, pf: 1.0, mf: 1.0, avm: AVM },
    };
    const existing = (await stGet(`scenarios:${user.email}`)) || [];
    await stSet(`scenarios:${user.email}`, [...existing, scenario]);
    setSaved(true); setTimeout(() => setSaved(false), 2500);
    if (onSave) onSave();
  }

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 24 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>II. mirovinski stup</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>
          Individualna kapitalizirana štednja · 5% bruto plaće · MOD
        </p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>

          {/* Bruto plaća */}
          <SmartParam
            icon="💼" title="Prosječna bruto plaća kroz karijeru"
            value={params.bruto_placa} min={500} max={5000} step={50}
            format={v => `${v.toLocaleString("hr-HR")} EUR/mj.`}
            onChange={v => setParams(p => ({ ...p, bruto_placa: v }))}
            zones={[
              { upTo: 799,  color: C.red,    label: "Ispod minimalne plaće" },
              { upTo: 1099, color: C.amber,  label: "Niža plaća" },
              { upTo: 1399, color: C.green,  label: "Blizu prosjeka RH" },
              { upTo: 2199, color: C.blue,   label: "Iznad prosjeka" },
              { upTo: 5000, color: C.purple, label: "Visoka plaća" },
            ]}
            dynamicExplain={v => {
              const mj_uplata = v * 0.05;
              return `Od vaše bruto plaće ${v.toLocaleString("hr-HR")} EUR, svaki mjesec ide ${mj_uplata.toFixed(0)} EUR na vaš osobni račun u mirovinskom fondu (5%). Godišnje: ${(mj_uplata*12).toFixed(0)} EUR.`;
            }}
            explain={
              <div>
                <p>Svaki mjesec, od vaše bruto plaće, <strong>5% ide na vaš osobni račun</strong> u Obveznom mirovinskom fondu (OMF). To su vaša osobna sredstva — nitko ih ne može uzeti.</p>
                <p style={{ marginTop: 6 }}>Uplaćuje <em>poslodavac</em> na osnovi zakona. Vi tu uplatu ne vidite direktno na plaći, ali ona se akumulira na vašem osobnom računu i raste s prinosima fonda.</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.purpleL, borderRadius: 6 }}>
                  💡 <strong>Praktično:</strong> Pri plaći 1.300 EUR bruto = 65 EUR/mj. → 780 EUR/god. → na 33 godine uz 5,5% prinos = ~95.000 EUR kapitala.
                </div>
              </div>
            }
          />

          {/* Godine štednje */}
          <SmartParam
            icon="⏱" title="Ukupno godina štednje u II. stupu"
            value={params.god_stednje} min={5} max={45} step={1}
            format={v => `${v} god.`}
            onChange={v => setParams(p => ({ ...p, god_stednje: v }))}
            zones={[
              { upTo: 14, color: C.amber,  label: "Kratka štednja" },
              { upTo: 24, color: C.gold,   label: "Srednje" },
              { upTo: 34, color: C.green,  label: "Dobro" },
              { upTo: 45, color: C.blue,   label: "Dugogodišnji štediša" },
            ]}
            dynamicExplain={v => {
              const k = calcStup2FV({ ...params, god_stednje: v });
              return `Nakon ${v} godina štednje uz sadašnje parametre, akumulirali biste oko ${fmt(k)} EUR — čak i bez povećanja plaće.`;
            }}
            explain={
              <div>
                <p>II. stup je uveden <strong>2002. godine</strong>. Osoba zaposlena 2002. ima maksimalno ~33 god. štednje do 2035.</p>
                <p style={{ marginTop: 6 }}>Što dulje štedite, to jači je efekt složene kamate (kamate na kamatu). Razlika između 25 i 35 godina može biti <em>deseci tisuća eura</em>.</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.blueL, borderRadius: 6 }}>
                  💡 <strong>Kamatni učinak:</strong> 10 dodatnih godina štednje ne znači 10% više kapitala — znači gotovo <strong>dvostruko</strong> više, zahvaljujući složenoj kamati.
                </div>
              </div>
            }
          />

          {/* Prinos */}
          <SmartParam
            icon="📈" title="Godišnji prinos mirovinskog fonda"
            value={params.neto_prinos} min={0.01} max={0.12} step={0.001}
            format={v => `${(v*100).toFixed(1)}% god.`}
            onChange={v => setParams(p => ({ ...p, neto_prinos: v }))}
            zones={[
              { upTo: 0.034, color: C.red,    label: "Konzervativno" },
              { upTo: 0.059, color: C.gold,   label: "Umjereno (prosjek HR OMF)" },
              { upTo: 0.079, color: C.green,  label: "Optimistično" },
              { upTo: 0.12,  color: C.purple, label: "Iznadprosječno" },
            ]}
            dynamicExplain={v => {
              const k_real = calcStup2FV({ ...params, neto_prinos: v });
              const k_0 = params.bruto_placa * 0.05 * params.god_stednje * 12;
              const prinos_eur = k_real - k_0;
              return `Uz ${(v*100).toFixed(1)}% prinos, kamate bi donijele ${fmt(prinos_eur)} EUR dodatno — to je ${Math.round(prinos_eur/k_0*100)}% više od pukih uplata. Naknada OMF (0,24%/god.) već je oduzeta.`;
            }}
            explain={
              <div>
                <p>OMF ulaže vaša sredstva u obveznice, dionice i nekretnine. Prinos <strong>nije zajamčen</strong>, ali HZMO propisuje minimalne standarde zaštite.</p>
                <ul style={{ marginTop: 8, paddingLeft: 16, display: "flex", flexDirection: "column", gap: 4 }}>
                  <li><strong>3,5%</strong> — konzervativna pretpostavka (loši uvjeti na tržištu)</li>
                  <li><strong>5,5%</strong> — prosjek HR OMF-ova 2002.–2024.</li>
                  <li><strong>7,5%</strong> — optimistično (dobra tržišna kretanja)</li>
                </ul>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.amberL, borderRadius: 6 }}>
                  ⚠ Razlika između 3,5% i 7,5% na 33 god. štednje = <strong>više od 60.000 EUR</strong> razlike u kapitalu pri plaći 1.300 EUR.
                </div>
              </div>
            }
          />

          {/* Jednokratna isplata */}
          <SmartParam
            icon="💵" title="Jednokratna isplata pri odlasku u mirovinu"
            value={params.jednokratni_pct} min={0} max={0.2} step={0.01}
            format={v => v === 0 ? "Ne uzimam jednokratno" : `${(v*100).toFixed(0)}% od kapitala`}
            onChange={v => setParams(p => ({ ...p, jednokratni_pct: v }))}
            zones={[
              { upTo: 0,    color: C.slate, label: "Sve u rentu" },
              { upTo: 0.09, color: C.blue,  label: "Mala jednokratna" },
              { upTo: 0.15, color: C.gold,  label: "Srednja jednokratna" },
              { upTo: 0.20, color: C.green, label: "Maksimum (20%)" },
            ]}
            dynamicExplain={v => {
              const jednokr = (kapital * v).toFixed(0);
              const preost = (kapital * (1-v)).toFixed(0);
              if (v === 0) return `Sav akumulirani kapital (${fmt(kapital)} EUR) ide u doživotnu rentu.`;
              return `Jednokratno primite ${fmt(Number(jednokr))} EUR odmah. Preostalih ${fmt(Number(preost))} EUR pretvara se u doživotnu rentu. ⚠ MOD obračunava porez na dohodak pri isplati.`;
            }}
            explain={
              <div>
                <p>Zakonom je dopušteno podići <strong>do 20% akumuliranog kapitala</strong> jednokratno pri umirovljenju — ali samo ako je vaša mirovina iz I. stupa veća od 115% najniže mirovine.</p>
                <p style={{ marginTop: 6 }}>To su vaši novci — možete ih iskoristiti kako želite (renovacija, otplata kredita, putovanje…). Na taj iznos platite porez na dohodak.</p>
                <p style={{ marginTop: 6 }}>Preostali kapital (80%+) pretvara se u doživotnu rentu kod MOD-a.</p>
              </div>
            }
          />

          {/* e65 */}
          <SmartParam
            icon="🧬" title="Preostali životni vijek u dobi 65 god."
            value={params.e65} min={12} max={26} step={0.5}
            format={v => `${v} god.`}
            onChange={v => setParams(p => ({ ...p, e65: v }))}
            zones={[
              { upTo: 15, color: C.amber,  label: "Ispod prosjeka" },
              { upTo: 20, color: C.green,  label: "Prosjek HR" },
              { upTo: 26, color: C.blue,   label: "Optimistično (duži život)" },
            ]}
            dynamicExplain={v => {
              const r = rentaDozivotna(preostali, v);
              return `Ako živite još ${v} god. nakon umirovljenja (do ${65 + v} god.), renta od ${fmt(r)} EUR/mj. isplaćuje se cijelo to vrijeme. Ukupno ćete primiti ~${fmt(r * v * 12)} EUR od II. stupa.`;
            }}
            explain={
              <div>
                <p>MOD izračunava rentu na temelju <strong>aktuarskih tablica</strong> — statističkih podataka o očekivanom životnom vijeku. Što dulje prosječno živite, to je renta niža (jer MOD mora isplaćivati dulje).</p>
                <ul style={{ marginTop: 8, paddingLeft: 16, display: "flex", flexDirection: "column", gap: 4 }}>
                  <li><strong>HR muškarci</strong> u dobi 65: prosj. još ~17–19 god. (do ~84)</li>
                  <li><strong>HR žene</strong> u dobi 65: prosj. još ~20–22 god. (do ~86)</li>
                </ul>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.blueL, borderRadius: 6 }}>
                  💡 Ovo je aproksimacija. MOD koristi vlastite aktuarske tablice i uzima u obzir spol, dob i odabrani tip rente.
                </div>
              </div>
            }
          />

          {/* Inovativne varijante */}
          <div style={{ borderRadius: 12, border: `2px solid ${C.purple}44`, background: C.purpleL, overflow: "hidden" }}>
            <div style={{ padding: "12px 14px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                <span style={{ fontSize: 18 }}>🧪</span>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: C.purple }}>Inovativne varijante isplate</div>
                  <div style={{ fontSize: 10, color: C.purple, opacity: 0.8 }}>Simulacija varijanti koje ne postoje u HR</div>
                </div>
              </div>
              {[
                { v: "standard",     emoji: "✅", l: "Standardna doživotna renta",
                  sub: "Isplaćuje se dok ste živi. Prestaje vašom smrću. Najviši iznos." },
                { v: "s_povratom",   emoji: "👨‍👩‍👧", l: "* Renta + povrat ostatka nasljednicima",
                  sub: "Nasljednici dobivaju ono što nije isplaćeno. Renta ~18% niža." },
                { v: "fiksno_n",     emoji: "📆", l: "* Fiksna renta 20 godina",
                  sub: "Isplaćuje se 20 god. bez obzira živite li vi ili nasljednici." },
                { v: "investicijska",emoji: "💹", l: "* Investicijska mirovina (4% pravilo)",
                  sub: "Kapital ostaje uložen. Primate 4% godišnje. Kapital se nasljeđuje." },
              ].map(o => (
                <button key={o.v} onClick={() => setParams(p => ({ ...p, inovativna_varijanta: o.v }))}
                  style={{ width: "100%", marginBottom: 6, padding: "8px 10px",
                    border: `2px solid ${params.inovativna_varijanta === o.v ? C.purple : C.border}`,
                    borderRadius: 9, background: params.inovativna_varijanta === o.v ? C.purple : C.white,
                    color: params.inovativna_varijanta === o.v ? "#fff" : C.navy,
                    cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>
                  <div style={{ fontSize: 12, fontWeight: 700 }}>{o.emoji} {o.l}</div>
                  <div style={{ fontSize: 10, opacity: 0.8, marginTop: 2, lineHeight: 1.3 }}>{o.sub}</div>
                </button>
              ))}
              <div style={{ fontSize: 10, color: C.purple, marginTop: 4, opacity: 0.8 }}>
                * Varijante s * su simulacije — ne postoje u HR zakonodavnom okviru (ZOMF)
              </div>
            </div>
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Akumulirani kapital</h3>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
              <ResultBox label="Ukupni kapital" value={`${fmt(kapital)} EUR`} color={C.purple} big />
              <ResultBox label="Neto uplate (bez prinosa)" value={`${fmt(params.bruto_placa * 0.05 * params.god_stednje * 12)} EUR`} color={C.slate} />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <ResultBox label={`Jednokratna isplata (${Math.round(params.jednokratni_pct*100)}%)`}
                value={`${fmt(jednokratni_iznos)} EUR`} color={C.gold} />
              <ResultBox label="Preostali kapital za rentu" value={`${fmt(preostali)} EUR`} color={C.blue} />
            </div>
          </Card>

          <Card style={{ border: `2px solid ${params.inovativna_varijanta !== "standard" ? C.purple : C.green}33` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
              <h3 style={{ fontSize: 14 }}>Doživotna renta</h3>
              {params.inovativna_varijanta !== "standard" && <Badge label="* Simulacija" color={C.purple} />}
            </div>
            <ResultBox
              label={params.inovativna_varijanta === "standard" ? "Okvirna doživotna renta / mj."
                : params.inovativna_varijanta === "s_povratom" ? "* Renta s povratom ostatka / mj."
                : params.inovativna_varijanta === "fiksno_n" ? "* Fiksna renta 20 god. / mj."
                : "* Investicijska mirovina / mj. (4%)"}
              value={`${fmt(renta_prikaz)} EUR`}
              color={params.inovativna_varijanta !== "standard" ? C.purple : C.green} big />
            {params.inovativna_varijanta !== "standard" && (
              <div style={{ marginTop: 10, fontSize: 11, color: C.purple, background: C.purpleL, padding: "8px 12px", borderRadius: 8 }}>
                <strong>Napomena:</strong> Ova varijanta ne postoji u HR zakonodavnom okviru (ZOMF). Prikazana je kao simulacija moguće buduće opcije. Stvarna renta utvrđuje se aktuarskim izračunom MOD-a.
              </div>
            )}
            {params.inovativna_varijanta === "standard" && (
              <div style={{ marginTop: 8, fontSize: 11, color: C.slateL }}>
                Aproksimacija: Kapital / (e₆₅ × 12). MOD koristi vlastite aktuarske tablice.
              </div>
            )}
          </Card>

          <Card>
            <h3 style={{ fontSize: 13, marginBottom: 12 }}>Usporedba varijanti isplate</h3>
            <ResponsiveContainer width="100%" height={160}>
              <BarChart data={varijante_data} margin={{ top: 5, right: 10, bottom: 30, left: 0 }}>
                <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-10} textAnchor="end" interval={0} />
                <YAxis tick={{ fontSize: 10 }} tickFormatter={v => `${Math.round(v)}`} />
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <Tooltip formatter={v => [`${fmt(v)} EUR/mj.`, "Renta"]} />
                <Bar dataKey="renta" radius={[4,4,0,0]}>
                  {varijante_data.map((d, i) => <Cell key={i} fill={i === 0 ? C.green : C.purple} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            <div style={{ fontSize: 10, color: C.purple, marginTop: 4 }}>* Varijante označene zvjezdicom su simulacije — ne postoje u HR</div>
          </Card>

          <Card style={{ padding: "14px 18px" }}>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              <input value={naziv} onChange={e => setNaziv(e.target.value)}
                style={{ flex: 1, padding: "8px 12px", border: `1.5px solid ${C.border}`, borderRadius: 8, fontSize: 13 }}
                placeholder="Naziv scenarija" />
              <button onClick={handleSave} style={{
                padding: "8px 18px", background: saved ? C.green : C.purple, color: "#fff",
                border: "none", borderRadius: 8, fontWeight: 600, fontSize: 13,
                display: "flex", alignItems: "center", gap: 6,
              }}>
                {saved ? <CheckCircle size={15} /> : <Save size={15} />}
                {saved ? "Spremljeno!" : "Spremi"}
              </button>
            </div>
          </Card>
        </div>
      </div>

      {/* Akumulacija chart */}
      <Card style={{ marginTop: 20 }}>
        <h3 style={{ fontSize: 15, marginBottom: 16 }}>Rast kapitala kroz godine</h3>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={proj_data}>
            <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
            <XAxis dataKey="god" tick={{ fontSize: 11 }} />
            <YAxis tick={{ fontSize: 11 }} tickFormatter={v => `${Math.round(v/1000)}k`} />
            <Tooltip formatter={v => [fmtEur(v), ""]} />
            <Legend />
            <Area type="monotone" dataKey="uplate" stackId="1" fill={C.blueL} stroke={C.blue} name="Uplate (bez prinosa)" />
            <Area type="monotone" dataKey="prinos" stackId="1" fill={C.amberL} stroke={C.gold} name="Kumulativni prinos" />
          </AreaChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

// ─── UKUPNO ──────────────────────────────────────────────────────────────────
function UkupnoView() {
  const [p1, setP1] = useState({ pvb: 1.05, staz: 40, pf: 1.0, mf: 1.0, avm: AVM });
  const [p2, setP2] = useState({ bruto_placa: 1300, god_stednje: 33, neto_prinos: 0.055, e65: 18, jednokratni_pct: 0.20 });

  const r1 = calcStup1(p1);
  const k2 = calcStup2FV(p2);
  const renta = rentaDozivotna(k2 * (1 - p2.jednokratni_pct));
  const dodatak = r1.efektivna * 0.27;
  const ukupno = r1.efektivna + dodatak + renta;

  const pie_data = [
    { name: "I. stup", value: Math.round(r1.efektivna), fill: C.navy },
    { name: "Dodatak 27%", value: Math.round(dodatak), fill: C.gold },
    { name: "II. stup renta", value: Math.round(renta), fill: C.purple },
  ];

  return (
    <div className="fadeIn">
      <h2 style={{ fontSize: 22, marginBottom: 4 }}>Ukupna kombinirana mirovina</h2>
      <p style={{ color: C.slateL, fontSize: 13, marginBottom: 24 }}>I. stup + Dodatak 27% + II. stup renta</p>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
        <Card>
          <h3 style={{ fontSize: 14, marginBottom: 14, color: C.navy }}>I. stup – brzi unos</h3>
          {[
            { k: "pvb", l: "PVB", min: 0.3, max: 3.8, step: 0.01, fmt: v => v.toFixed(2) },
            { k: "staz", l: "Staž (god.)", min: 1, max: 50, step: 1, fmt: v => `${v}` },
          ].map(f => (
            <Slider key={f.k} label={f.l} value={p1[f.k]} min={f.min} max={f.max} step={f.step}
              format={f.fmt} onChange={v => setP1(prev => ({ ...prev, [f.k]: v }))} />
          ))}
        </Card>
        <Card>
          <h3 style={{ fontSize: 14, marginBottom: 14, color: C.purple }}>II. stup – brzi unos</h3>
          {[
            { k: "bruto_placa", l: "Bruto plaća (EUR/mj.)", min: 500, max: 4000, step: 50, fmt: v => `${v.toLocaleString("hr-HR")}` },
            { k: "god_stednje", l: "Godina štednje", min: 5, max: 45, step: 1, fmt: v => `${v}` },
            { k: "neto_prinos", l: "Neto prinos (%)", min: 0.01, max: 0.12, step: 0.001, fmt: v => `${(v*100).toFixed(1)}%` },
          ].map(f => (
            <Slider key={f.k} label={f.l} value={p2[f.k]} min={f.min} max={f.max} step={f.step}
              format={f.fmt} onChange={v => setP2(prev => ({ ...prev, [f.k]: v }))} />
          ))}
        </Card>
      </div>

      <Card style={{ marginBottom: 20, background: `linear-gradient(135deg, ${C.navy}, #243a68)` }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
          {[
            { l: "I. stup (neto)", v: r1.efektivna, c: "#7fbcff" },
            { l: "+ Dodatak 27%", v: dodatak, c: C.goldL },
            { l: "+ II. stup renta", v: renta, c: "#c4a8ff" },
            { l: "= UKUPNO / MJ.", v: ukupno, c: "#7bffb0" },
          ].map((it, i) => (
            <div key={i} style={{ textAlign: "center", padding: "12px 8px" }}>
              <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)", marginBottom: 6 }}>{it.l}</div>
              <div style={{ fontSize: i === 3 ? 26 : 20, fontWeight: 800, color: it.c, fontFamily: "'Playfair Display', serif" }}>
                {fmt(it.v)} EUR
              </div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        <Card>
          <h3 style={{ fontSize: 14, marginBottom: 14 }}>Struktura mirovine</h3>
          <ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie data={pie_data} dataKey="value" cx="50%" cy="50%" innerRadius={55} outerRadius={85} paddingAngle={3} label={({ name, percent }) => `${name} ${(percent*100).toFixed(0)}%`} labelLine={false}>
                {pie_data.map((d, i) => <Cell key={i} fill={d.fill} />)}
              </Pie>
              <Tooltip formatter={v => [fmtEur(v), ""]} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <h3 style={{ fontSize: 14, marginBottom: 14 }}>Godišnji pregled</h3>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {[
              { l: "Godišnja mirovina", v: ukupno * 12 },
              { l: "Akumulirani kapital II. stupa", v: k2 },
              { l: "Jednokratna isplata (20%)", v: k2 * p2.jednokratni_pct },
              { l: "Stopa zamjene (vs. bruto plaća)", v: null, pct: ukupno / p2.bruto_placa },
            ].map((it, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "8px 12px", background: i%2===0 ? C.bg : C.white, borderRadius: 7 }}>
                <span style={{ fontSize: 12, color: C.slate }}>{it.l}</span>
                <strong style={{ fontSize: 13, color: C.navy }}>
                  {it.pct !== undefined ? pct(it.pct) : fmtEur(it.v)}
                </strong>
              </div>
            ))}
          </div>
          <div style={{ marginTop: 12, padding: "10px 12px", background: C.amberL, borderRadius: 8, fontSize: 11, color: C.amber }}>
            <strong>Napomena:</strong> Renta II. stupa je aproksimacija. Dodatak 27% važi za kombiniranu mirovinu od 1.1.2024. (NN 156/2023).
          </div>
        </Card>
      </div>
    </div>
  );
}

// ─── MOJ PROFIL / SCENARIJI ──────────────────────────────────────────────────
function MojiScenarijiView({ user, onAuthRequired }) {
  const [scenariji, setScenariji] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!user) { setLoading(false); return; }
    stGet(`scenarios:${user.email}`).then(d => {
      setScenariji(d || []);
      setLoading(false);
    });
  }, [user?.email]);

  async function deleteScenario(id) {
    const updated = scenariji.filter(s => s.id !== id);
    await stSet(`scenarios:${user.email}`, updated);
    setScenariji(updated);
  }

  if (loading) return <div style={{ padding: 40, textAlign: "center", color: C.slateL }}>Učitavanje…</div>;

  // Guest: show CTA to login
  if (!user) return (
    <div className="fadeIn">
      <Card style={{ textAlign: "center", padding: "60px 40px", maxWidth: 520, margin: "40px auto" }}>
        <div style={{ fontSize: 52, marginBottom: 16 }}>🔒</div>
        <h2 style={{ fontSize: 22, marginBottom: 10 }}>Prijava za scenarije</h2>
        <p style={{ color: C.slateL, fontSize: 14, lineHeight: 1.7, marginBottom: 24 }}>
          Scenariji se čuvaju vezano uz vaš korisnički račun kako biste im mogli pristupiti i u budućim sesijama.<br /><br />
          Kalkulator možete koristiti bez prijave — registrirajte se samo kada želite <strong>sačuvati ili izvesti</strong> rezultate.
        </p>
        <button onClick={() => onAuthRequired("Registracija je besplatna i traje 30 sekundi.")}
          style={{ padding: "13px 32px", background: `linear-gradient(135deg, ${C.gold}, ${C.goldL})`,
            border: "none", borderRadius: 10, color: C.navy, fontWeight: 700, fontSize: 15,
            boxShadow: "0 4px 14px rgba(201,146,43,0.3)", cursor: "pointer" }}>
          Prijavi se / Registriraj
        </button>
      </Card>
    </div>
  );

  return (
    <div className="fadeIn">
      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <h2 style={{ fontSize: 22, marginBottom: 4 }}>Moji scenariji</h2>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <span style={{ fontSize: 13, color: C.slateL }}>{scenariji.length} {scenariji.length === 1 ? "scenarij" : "scenarija"}</span>
            {["I. stup","II. stup","III. stup"].map(tip => {
              const cnt = scenariji.filter(s =>
                tip === "I. stup" ? (s.stup1 && !s.stup2 && !s.stup3) :
                tip === "II. stup" ? (s.stup2) :
                (s.stup3)
              ).length;
              if (!cnt) return null;
              return <Badge key={tip} label={`${tip}: ${cnt}`} color={tip === "I. stup" ? C.navy : tip === "II. stup" ? C.purple : C.green} />;
            })}
          </div>
        </div>
        <button
          onClick={() => exportToExcel(scenariji, user.name)}
          disabled={scenariji.length === 0}
          style={{
            padding: "11px 22px",
            background: scenariji.length === 0
              ? C.slateL
              : `linear-gradient(135deg, ${C.green}, #1a5c35)`,
            color: "#fff", border: "none", borderRadius: 9, fontWeight: 700,
            display: "flex", alignItems: "center", gap: 8, fontSize: 13,
            boxShadow: scenariji.length > 0 ? "0 4px 14px rgba(45,122,79,0.35)" : "none",
            cursor: scenariji.length === 0 ? "not-allowed" : "pointer",
            transition: "all 0.2s",
          }}>
          <Download size={16} />
          Izvoz u Excel ({scenariji.length})
        </button>
      </div>

      {/* Info strip */}
      {scenariji.length > 0 && (
        <div style={{ background: C.greenL, border: `1px solid ${C.green}33`, borderRadius: 10,
          padding: "10px 16px", marginBottom: 20, fontSize: 12, color: C.slate, display: "flex", gap: 16, flexWrap: "wrap" }}>
          <span>📊 Excel izvoz uključuje: <strong>Pregled svih scenarija</strong></span>
          {scenariji.some(s => s.stup1) && <span>· <strong>I. stup</strong> detalji</span>}
          {scenariji.some(s => s.stup2) && <span>· <strong>II. stup</strong> detalji</span>}
          {scenariji.some(s => s.stup3) && <span>· <strong>III. stup</strong> + DPS</span>}
          <span>· <strong>Projekcija AVM</strong> 2026.–2035.</span>
          <span>· <strong>Formule i pravni izvori</strong></span>
        </div>
      )}

      {scenariji.length === 0 ? (
        <Card style={{ textAlign: "center", padding: 60 }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>📋</div>
          <h3 style={{ marginBottom: 8 }}>Nema spremljenih scenarija</h3>
          <p style={{ color: C.slateL, fontSize: 14 }}>Izračunajte mirovinu u bilo kojem kalkulatoru i kliknite "Spremi".</p>
        </Card>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          {scenariji.map(s => {
            const r1 = s.stup1 ? calcStup1(s.stup1) : null;
            const k2 = s.stup2 ? calcStup2FV(s.stup2) : null;
            const renta2 = k2 ? rentaDozivotna(k2 * (1 - (s.stup2.jednokratni_pct || 0))) : null;
            const k3 = s.stup3 ? (() => {
              const p = s.stup3;
              const dps_mj = Math.min(p.uplata * 12, 663.61) * 0.15 / 12;
              const pmt = p.uplata + (p.posl || 0) + dps_mj;
              const r = p.prinos / 12; const n = p.god * 12;
              return r === 0 ? pmt * n : pmt * ((Math.pow(1 + r, n) - 1) / r);
            })() : null;

            const tip = s.stup3 ? "III. stup" : s.stup2 && s.stup1 ? "I.+II." : s.stup1 ? "I. stup" : "II. stup";
            const tipColor = s.stup3 ? C.green : s.stup2 ? C.purple : C.navy;

            return (
              <Card key={s.id} style={{ position: "relative", border: `1.5px solid ${tipColor}22` }}>
                <button onClick={() => deleteScenario(s.id)} style={{
                  position: "absolute", top: 12, right: 12, background: "none", border: "none",
                  color: C.red, cursor: "pointer", opacity: 0.5,
                }}>
                  <Trash2 size={15} />
                </button>

                <div style={{ display: "flex", alignItems: "flex-start", gap: 8, marginBottom: 10, paddingRight: 24 }}>
                  <div>
                    <h3 style={{ fontSize: 15, marginBottom: 3 }}>{s.naziv || "Bez naziva"}</h3>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <Badge label={tip} color={tipColor} />
                      <span style={{ fontSize: 10, color: C.slateL }}>
                        {new Date(s.created).toLocaleDateString("hr-HR")}
                      </span>
                    </div>
                  </div>
                </div>

                {r1 && (
                  <div style={{ background: C.blueL, borderRadius: 8, padding: "8px 12px", marginBottom: 6 }}>
                    <div style={{ fontSize: 10, color: C.slate, fontWeight: 600, marginBottom: 2 }}>I. STUP</div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ fontSize: 16, fontWeight: 800, color: C.navy }}>{fmtEur(r1.efektivna)}<span style={{ fontSize: 10, fontWeight: 400 }}>/mj.</span></div>
                      <div style={{ fontSize: 10, color: C.slateL, textAlign: "right" }}>
                        PVB {s.stup1.pvb.toFixed(2)} · Staž {s.stup1.staz} god.<br/>PF {s.stup1.pf.toFixed(3)}
                      </div>
                    </div>
                  </div>
                )}

                {k2 !== null && (
                  <div style={{ background: C.purpleL, borderRadius: 8, padding: "8px 12px", marginBottom: 6 }}>
                    <div style={{ fontSize: 10, color: C.purple, fontWeight: 600, marginBottom: 2 }}>II. STUP</div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 14, fontWeight: 700, color: C.purple }}>{fmtEur(k2)} kapital</div>
                        <div style={{ fontSize: 11, color: C.slateL }}>Renta: {fmtEur(renta2)}/mj.</div>
                      </div>
                      <div style={{ fontSize: 10, color: C.slateL, textAlign: "right" }}>
                        {s.stup2.god_stednje} god. štednje<br/>{(s.stup2.neto_prinos * 100).toFixed(1)}% prinos
                      </div>
                    </div>
                  </div>
                )}

                {k3 !== null && (
                  <div style={{ background: C.greenL, borderRadius: 8, padding: "8px 12px", marginBottom: 6 }}>
                    <div style={{ fontSize: 10, color: C.green, fontWeight: 600, marginBottom: 2 }}>III. STUP + DPS</div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ fontSize: 14, fontWeight: 700, color: C.green }}>{fmtEur(k3)} kapital</div>
                      <div style={{ fontSize: 10, color: C.slateL, textAlign: "right" }}>
                        {fmtEur(s.stup3.uplata)}/mj.<br/>{s.stup3.god} god. štednje
                      </div>
                    </div>
                  </div>
                )}

                {/* Ukupno */}
                {(r1 || k2) && (
                  <div style={{ background: C.navy, borderRadius: 8, padding: "7px 12px", marginTop: 4 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 10, color: "rgba(255,255,255,0.6)" }}>Ukupno (I.+27%+II.)</span>
                      <span style={{ fontSize: 15, fontWeight: 800, color: "#7bffb0" }}>
                        {fmtEur((r1 ? r1.efektivna * 1.27 : 0) + (renta2 || 0))}/mj.
                      </span>
                    </div>
                  </div>
                )}
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── S&P 500 MIROVINSKI KALKULATOR ──────────────────────────────────────────
// Izvor: Fidelity.com — S&P 500 prosječni godišnji prinos zadnjih 40 god. = 11,5% (zaklj. prosinac 2025.)
// Izvor: NYU Stern / Damodaran — dugoročni CAGR 1928.–2024. = 10,2% s reinvestiranim dividendama
const SP500_40Y_NOMINAL   = 0.115;  // 11,5% god. nominalno, zadnjih 40 god. (Fidelity, prosinac 2025.)
const SP500_40Y_REALNI    = 0.072;  // ~7,2% realno (prilagođeno inflacijom, procjena)
const SP500_LONG_NOMINAL  = 0.102;  // 10,2% god. dugoročni CAGR od 1928. (NYU Stern / Damodaran)

function SP500MirovinaBlog() {
  const [placa, setPlaca] = useState(1300);
  const [staz, setStaz] = useState(40);
  const [doprinos_pct, setDoprinos] = useState(0.20);
  const [prinos_mode, setPrinosMode] = useState("40y_nominal"); // 40y_nominal | 40y_realni | long_nominal | custom
  const [custom_prinos, setCustomPrinos] = useState(0.07);
  const [showDetail, setShowDetail] = useState(false);

  const prinos_map = {
    "40y_nominal": SP500_40Y_NOMINAL,
    "40y_realni":  SP500_40Y_REALNI,
    "long_nominal":SP500_40Y_LONG_NOMINAL || SP500_LONG_NOMINAL,
    "custom":      custom_prinos,
  };
  const prinos = prinos_map[prinos_mode];

  // Akumulacija: FV anuitetna renta (kraj perioda)
  function fvRenta(pmt_mj, god, r_god) {
    const r = r_god / 12; const n = god * 12;
    if (r === 0) return pmt_mj * n;
    return pmt_mj * ((Math.pow(1 + r, n) - 1) / r);
  }

  const pmt = placa * doprinos_pct;
  const kapital_sp500 = fvRenta(pmt, staz, prinos);
  const renta_4pct = kapital_sp500 * 0.04 / 12;   // 4% pravilo
  const renta_3pct = kapital_sp500 * 0.03 / 12;   // konzervativno

  // Usporedba s HR sustavom (20% ide u I.+II.)
  const pvb_est = placa / 1300;
  const hr_i = calcStup1({ pvb: pvb_est, staz, pf: 1.0, mf: 1.0, avm: AVM }).efektivna;
  const hr_dod = hr_i * 0.27;
  const hr_ii = rentaDozivotna(calcStup2FV({ bruto_placa: placa, god_stednje: staz, neto_prinos: 0.055 }) * 0.8);
  const hr_uk = hr_i + hr_dod + hr_ii;

  const razlika_kapital = kapital_sp500 - (hr_uk * 18 * 12); // HR ~18 god. isplate
  const razlika_renta   = renta_4pct - hr_uk;

  // Godišnji presjek — rast 1k EUR jednokratno investirano
  const snapshotData = [
    { god: "5g",  hr: fvRenta(pmt,5,0.055), sp: fvRenta(pmt,5,prinos) },
    { god: "10g", hr: fvRenta(pmt,10,0.055), sp: fvRenta(pmt,10,prinos) },
    { god: "15g", hr: fvRenta(pmt,15,0.055), sp: fvRenta(pmt,15,prinos) },
    { god: "20g", hr: fvRenta(pmt,20,0.055), sp: fvRenta(pmt,20,prinos) },
    { god: "25g", hr: fvRenta(pmt,25,0.055), sp: fvRenta(pmt,25,prinos) },
    { god: "30g", hr: fvRenta(pmt,30,0.055), sp: fvRenta(pmt,30,prinos) },
    { god: `${staz}g`, hr: fvRenta(pmt,staz,0.055), sp: kapital_sp500 },
  ].filter((d,i,a) => i === a.findIndex(x => x.god === d.god));

  const tw = v => `${Math.round(v/1000)}k`;

  const prinosOpcije = [
    { id: "40y_nominal", label: "11,5% — S&P 500 zadnjih 40 god. nominalno", sublabel: "Fidelity, prosinac 2025." },
    { id: "long_nominal", label: "10,2% — S&P 500 od 1928. nominalno", sublabel: "NYU Stern / Damodaran" },
    { id: "40y_realni",  label: "7,2% — S&P 500 zadnjih 40 god. realno", sublabel: "Prilagođeno inflacijom (procjena)" },
    { id: "custom",      label: "Vlastiti prinos", sublabel: "" },
  ];

  return (
    <div style={{ marginBottom: 32 }}>
      {/* Hero header */}
      <div style={{ background: `linear-gradient(135deg, #0a1628, #1a2e52, #0d3321)`,
        borderRadius: 16, padding: "28px 32px", marginBottom: 20, position: "relative", overflow: "hidden" }}>
        <div style={{ position: "absolute", right: -10, top: -10, fontSize: 100, opacity: 0.06 }}>📈</div>
        <div style={{ fontSize: 10, fontWeight: 700, color: "#7bffb0", textTransform: "uppercase",
          letterSpacing: "0.12em", marginBottom: 10 }}>
          Misaoni eksperiment · Izvor: Fidelity.com · NYU Stern
        </div>
        <h2 style={{ fontSize: 22, color: "#fff", marginBottom: 10, lineHeight: 1.3 }}>
          Što da je 20% plaće išlo u S&P 500 umjesto u mirovinski sustav?
        </h2>
        <p style={{ fontSize: 13, color: "rgba(255,255,255,0.75)", lineHeight: 1.8, maxWidth: 620 }}>
          Prosječni godišnji prinos S&P 500 indeksa (s reinvestiranim dividendama) zadnjih 40 godina iznosi
          <strong style={{ color: "#7bffb0" }}> 11,5% nominalno</strong> (Fidelity, prosinac 2025.) odnosno
          <strong style={{ color: "#7bffb0" }}> 10,2% dugoročno</strong> od 1928. (NYU Stern / Damodaran).
          Koliko bi akumulirala osoba prosječne hrv. plaće da je isti iznos doprinosa (20% bruto)
          investirala u indeksni fond umjesto u mirovinski sustav?
        </p>
        <div style={{ marginTop: 14, padding: "10px 14px", background: "rgba(255,200,0,0.1)",
          borderRadius: 8, fontSize: 11, color: "#ffd966", lineHeight: 1.6 }}>
          ⚠ Ovaj prikaz je <strong>misaoni eksperiment</strong> — ne investicijska preporuka.
          Prošli prinosi S&P 500 ne garantiraju buduće. Volatilnost, devizni rizik i porezni tretman
          nisu uključeni. Usporedba zanemaruje socijalnu komponentu, invalidsku zaštitu i zajamčenu isplatu mirovinskog sustava.
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "340px 1fr", gap: 20 }}>
        {/* Lijevo: parametri */}
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 16, color: C.navy }}>Parametri simulacije</h3>

            <div style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 8 }}>Pretpostavljeni prinos S&P 500</div>
              {prinosOpcije.map(o => (
                <button key={o.id} onClick={() => setPrinosMode(o.id)}
                  style={{ width: "100%", marginBottom: 6, padding: "8px 11px",
                    border: `2px solid ${prinos_mode === o.id ? C.green : C.border}`,
                    borderRadius: 9, background: prinos_mode === o.id ? C.greenL : C.white,
                    cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: prinos_mode === o.id ? C.green : C.navy }}>
                    {o.label}
                  </div>
                  {o.sublabel && <div style={{ fontSize: 10, color: C.slateL, marginTop: 2 }}>{o.sublabel}</div>}
                </button>
              ))}
              {prinos_mode === "custom" && (
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6 }}>
                  <span style={{ fontSize: 12, color: C.slate }}>Godišnji prinos:</span>
                  <input className="num-inp"
                    value={(custom_prinos * 100).toFixed(1)}
                    onChange={e => setCustomPrinos(Math.min(0.25, Math.max(0.01, Number(e.target.value) / 100 || 0.07)))}
                    style={{ width: 64 }} />
                  <span style={{ fontSize: 12, color: C.slate }}>%</span>
                </div>
              )}
            </div>

            <MiniSlider label="Prosječna bruto plaća (EUR/mj.)" value={placa} min={700} max={5000} step={50}
              format={v => `${v.toLocaleString("hr-HR")} EUR`} onChange={setPlaca} color={C.green} />

            <MiniSlider label="Godina investiranja (radni staž)" value={staz} min={10} max={45} step={1}
              format={v => `${v} god.`} onChange={setStaz} color={C.green} />

            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 500, color: C.slate, marginBottom: 8 }}>
                Udio plaće koji se ulaže (HR: 20% bruto = I.+II. stup)
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {[0.10, 0.15, 0.20, 0.25].map(p => (
                  <button key={p} onClick={() => setDoprinos(p)}
                    style={{ flex: "1 1 60px", padding: "6px 4px",
                      border: `2px solid ${doprinos_pct === p ? C.green : C.border}`,
                      borderRadius: 8, background: doprinos_pct === p ? C.green : C.white,
                      color: doprinos_pct === p ? "#fff" : C.navy,
                      fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                    {Math.round(p*100)}%
                    {p === 0.20 && <div style={{ fontSize: 9, opacity: 0.8 }}>HR sustav</div>}
                  </button>
                ))}
              </div>
            </div>

            {/* Live summary */}
            <div style={{ background: C.navy, borderRadius: 10, padding: "12px 14px" }}>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.55)", marginBottom: 6 }}>
                Mjesečna investicija
              </div>
              <div style={{ fontSize: 22, fontWeight: 800, color: "#7bffb0",
                fontFamily: "'Playfair Display',serif" }}>
                {fmt(pmt)} EUR/mj.
              </div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.5)", marginTop: 2 }}>
                = {doprinos_pct*100}% od {placa.toLocaleString("hr-HR")} EUR bruto
              </div>
            </div>
          </Card>
        </div>

        {/* Desno: rezultati */}
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

          {/* Glavni rezultat */}
          <div style={{ background: `linear-gradient(135deg, #0d3321, #1a5c35)`,
            borderRadius: 14, padding: "20px 24px", color: "#fff" }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#9fefc0", textTransform: "uppercase",
              letterSpacing: "0.1em", marginBottom: 12 }}>
              S&P 500 simulacija · {(prinos*100).toFixed(1)}% god. · {staz} god.
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
              <div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.55)", marginBottom: 4 }}>Akumulirani kapital</div>
                <div style={{ fontSize: 28, fontWeight: 900, color: "#7bffb0",
                  fontFamily: "'Playfair Display',serif", lineHeight: 1 }}>
                  {fmt(kapital_sp500)} EUR
                </div>
              </div>
              <div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.55)", marginBottom: 4 }}>Renta — 4% pravilo/mj.</div>
                <div style={{ fontSize: 22, fontWeight: 800, color: "#a8f0c0",
                  fontFamily: "'Playfair Display',serif" }}>{fmt(renta_4pct)} EUR</div>
                <div style={{ fontSize: 10, color: "rgba(255,255,255,0.4)" }}>Konzervativno (3%): {fmt(renta_3pct)} EUR</div>
              </div>
              <div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.55)", marginBottom: 4 }}>Stopa zamjene</div>
                <div style={{ fontSize: 22, fontWeight: 800, color: "#a8f0c0",
                  fontFamily: "'Playfair Display',serif" }}>{pct(renta_4pct / placa)}</div>
                <div style={{ fontSize: 10, color: "rgba(255,255,255,0.4)" }}>vs. bruto plaće</div>
              </div>
            </div>
          </div>

          {/* Direktna usporedba */}
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>
              Usporedba: S&P 500 vs. HR mirovinski sustav
            </h3>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
              {[
                { l: "Akumulirani kapital — S&P 500", v: kapital_sp500, c: C.green },
                { l: "Equiv. HR kapital (18 god. × uk. mir.)", v: hr_uk * 18 * 12, c: C.navy },
                { l: "Renta S&P (4% pravilo) / mj.", v: renta_4pct, c: C.green },
                { l: "HR ukupna mirovina (I.+27%+II.) / mj.", v: hr_uk, c: C.navy },
              ].map((it, i) => (
                <div key={i} style={{ background: it.c+"0a", border:`1.5px solid ${it.c}22`,
                  borderRadius:10, padding:"10px 12px" }}>
                  <div style={{ fontSize:10, color:C.slateL, marginBottom:4, lineHeight:1.3 }}>{it.l}</div>
                  <div style={{ fontSize:18, fontWeight:800, color:it.c, fontFamily:"'Playfair Display',serif" }}>
                    {fmt(it.v)} EUR
                  </div>
                </div>
              ))}
            </div>

            {/* Razlika */}
            <div style={{ padding: "12px 16px", borderRadius: 10,
              background: razlika_renta > 0 ? C.greenL : C.redL,
              border: `1.5px solid ${razlika_renta > 0 ? C.green : C.red}33` }}>
              <div style={{ fontSize: 12, fontWeight: 700,
                color: razlika_renta > 0 ? C.green : C.red, marginBottom: 4 }}>
                {razlika_renta > 0 ? "📈 S&P 500 bi dao VIŠU rentu za:" : "📉 HR sustav bi dao VIŠU rentu za:"}
              </div>
              <div style={{ fontSize: 20, fontWeight: 900,
                color: razlika_renta > 0 ? C.green : C.red, fontFamily:"'Playfair Display',serif" }}>
                {fmt(Math.abs(razlika_renta))} EUR/mj.
                <span style={{ fontSize: 13, fontWeight: 500, marginLeft: 8 }}>
                  ({pct(Math.abs(razlika_renta) / hr_uk)} razlike)
                </span>
              </div>
              <div style={{ fontSize: 11, color: C.slateL, marginTop: 6 }}>
                Uz prinos {(prinos*100).toFixed(1)}% · {staz} god. · {Math.round(doprinos_pct*100)}% plaće
              </div>
            </div>
          </Card>

          {/* Graf */}
          <Card>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 12 }}>
              Rast kapitala kroz godine — S&P 500 vs. HR OMF (5,5%)
            </div>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={snapshotData} margin={{ top: 5, right: 10, bottom: 5, left: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <XAxis dataKey="god" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 10 }} tickFormatter={tw} />
                <Tooltip formatter={v => [fmtEur(v), ""]} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="sp" name={`S&P 500 (${(prinos*100).toFixed(1)}%)`} fill={C.green} radius={[4,4,0,0]} />
                <Bar dataKey="hr" name="HR OMF (5,5%)" fill={C.navy+"aa"} radius={[4,4,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </Card>
        </div>
      </div>

      {/* Zašto razlika postoji — detaljno */}
      <Card style={{ marginTop: 14, borderLeft: `4px solid ${C.gold}` }}>
        <button onClick={() => setShowDetail(v => !v)}
          style={{ width: "100%", background: "none", border: "none", cursor: "pointer",
            display: "flex", justifyContent: "space-between", alignItems: "center", padding: 0 }}>
          <h3 style={{ fontSize: 14, color: C.navy }}>
            🔍 Zašto je razlika tolika — i što nismo uzeli u obzir?
          </h3>
          {showDetail ? <ChevronUp size={16} color={C.slateL} /> : <ChevronDown size={16} color={C.slateL} />}
        </button>
        {showDetail && (
          <div style={{ marginTop: 14, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.green, marginBottom: 8 }}>
                ✅ Prednosti privatne investicijske štednje
              </div>
              {[
                ["Složena kamata na višem prinosu", `S&P 500 historijski ${(SP500_40Y_NOMINAL*100).toFixed(1)}% vs. HR OMF ~5,5% → razlika se eksponencijalno povećava s vremenom`],
                ["Kapital je osobna imovina", "Nasljeđuje se u potpunosti. Nema 'gubitka' pri ranoj smrti. Preostaje obitelji."],
                ["Nezavisnost od demografije", "Ne ovisi o omjeru radnika i umirovljenika. Demografska bomba nema utjecaj."],
                ["Fleksibilnost isplate", "Vlastiti ritam trošenja. Pravilo 4% ili 3% omogućava doživotnu isplatu s rastom kapitala."],
              ].map(([t,d],i) => (
                <div key={i} style={{ marginBottom: 10, paddingBottom: 10,
                  borderBottom: i < 3 ? `1px solid ${C.border}` : "none" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, marginBottom: 2 }}>{t}</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>{d}</div>
                </div>
              ))}
            </div>
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.red, marginBottom: 8 }}>
                ⚠ Što simulacija ne uzima u obzir
              </div>
              {[
                ["Volatilnost i redoslijed prinosa", "S&P 500 je izgubio 50% u 2000.–2002. i 2008.–2009. Osoba koja ode u mirovinu na dnu tržišta suočava se s katastrofom (sequence-of-returns risk)."],
                ["Devizni rizik", "S&P 500 je u USD. Tečaj EUR/USD može bitno promijeniti realnu vrijednost za hr. umirovljenika."],
                ["Invalidska i obiteljska zaštita", "HR sustav pruža invalidsku i obiteljsku mirovinu. Privatna štednja ne garantira ovaj 'osigurani' dio."],
                ["Zajamčena doživotna isplata", "Doživotna renta iz MOD-a isplaćuje se bez obzira koliko dugo živite. Privatni kapital se troši i može nestati."],
                ["Porezi i troškovi", "Prinosi na investicije podliježu porezima (10% u HR). OMF naknade već oduzete od prinosa. Brokeri naplaćuju provizije."],
                ["Behavioralni rizici", "Studije pokazuju da prosječni investitor ostvaruje 4-5% lošiji prinos od indeksa zbog loše alokacije (buy high, sell low)."],
              ].map(([t,d],i) => (
                <div key={i} style={{ marginBottom: 10, paddingBottom: 10,
                  borderBottom: i < 5 ? `1px solid ${C.border}` : "none" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, marginBottom: 2 }}>{t}</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>{d}</div>
                </div>
              ))}
            </div>
            <div style={{ gridColumn: "1 / -1", padding: "12px 16px", background: C.amberL,
              borderRadius: 10, fontSize: 12, color: C.amber, lineHeight: 1.7 }}>
              <strong>Što ovi podaci pokazuju:</strong> S matematičkog gledišta, uz pretpostavku
              povijesnih S&P 500 prinosa i dugo investicijsko razdoblje, privatna kapitalizirana
              štednja generira bitno više bogatstva od mješovitog HR sustava. No, mirovinski
              sustav nije samo štednja — on je i <strong>osiguranje</strong>: od dugovječnosti,
              invalidnosti i prerane smrti supružnika. Privatna štednja te rizike ne pokriva
              automatski. Svaki od ovih modela nosi vlastite prednosti, rizike i kompromise —
              koji su prihvatljivi ovisi o vrijednostima, okolnostima i preferencijama svakog pojedinca.
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

// ─── BLOG ────────────────────────────────────────────────────────────────────
function BlogView() {
  const [post, setPost] = useState(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiQ, setAiQ] = useState("");
  const [aiA, setAiA] = useState("");
  const [filter, setFilter] = useState("Sve");

  const cats = ["Sve", ...new Set(BLOG_POSTS.map(p => p.category))];
  const filtered = filter === "Sve" ? BLOG_POSTS : BLOG_POSTS.filter(p => p.category === filter);

  async function askAI() {
    if (!aiQ.trim()) return;
    setAiLoading(true); setAiA("");
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6", max_tokens: 1000,
          system: `Ti si stručnjak za hrvatski mirovinski sustav. Odgovaraj na pitanja korisnika jasno, točno i sažeto na hrvatskom jeziku. Citiraj relevantne zakone (ZOMO, ZOMF) i službene izvore (HZMO, HANFA, mss.gov.hr). AVM od 1.7.2025. = 14,45 EUR. Nikad ne izmišljaj podatke — ako ne znaš, reci to jasno.`,
          messages: [{ role: "user", content: aiQ }]
        })
      });
      const d = await res.json();
      setAiA(d.content?.[0]?.text || "Greška u odgovoru.");
    } catch (e) {
      setAiA("Nije moguće dohvatiti odgovor. Provjerite vezu.");
    }
    setAiLoading(false);
  }

  function renderMd(text) {
    return text.split("\n").map((line, i) => {
      if (line.startsWith("## ")) return <h2 key={i} style={{ fontSize: 18, marginTop: 24, marginBottom: 8, color: C.navy }}>{line.slice(3)}</h2>;
      if (line.startsWith("### ")) return <h3 key={i} style={{ fontSize: 15, marginTop: 16, marginBottom: 6, color: C.navyL }}>{line.slice(4)}</h3>;
      if (line.startsWith("> ")) return <blockquote key={i} style={{ borderLeft: `3px solid ${C.gold}`, paddingLeft: 12, color: C.slate, fontStyle: "italic", margin: "12px 0" }}>{line.slice(2)}</blockquote>;
      if (line.startsWith("| ")) return null;
      if (line === "") return <br key={i} />;
      const boldLine = line.replace(/\*\*(.*?)\*\*/g, (_, t) => `<strong>${t}</strong>`);
      return <p key={i} style={{ lineHeight: 1.7, marginBottom: 6 }} dangerouslySetInnerHTML={{ __html: boldLine }} />;
    });
  }

  if (post) {
    return (
      <div className="fadeIn">
        <button onClick={() => setPost(null)} style={{ background: "none", border: "none", color: C.blue, fontSize: 13, marginBottom: 20, display: "flex", alignItems: "center", gap: 4, cursor: "pointer" }}>
          ← Natrag na blog
        </button>
        <Card style={{ maxWidth: 740 }}>
          <div style={{ marginBottom: 6 }}>
            <Badge label={post.category} color={C.blue} />
            <span style={{ fontSize: 11, color: C.slateL, marginLeft: 10 }}>{post.readMin} min čitanja · {post.date}</span>
          </div>
          <h1 style={{ fontSize: 26, marginTop: 12, marginBottom: 18, lineHeight: 1.3 }}>{post.title}</h1>
          <div style={{ fontSize: 14, color: C.slate, lineHeight: 1.7 }}>{renderMd(post.content)}</div>
          <div style={{ marginTop: 20, display: "flex", gap: 6, flexWrap: "wrap" }}>
            {post.tags.map(t => <Badge key={t} label={`#${t}`} color={C.slateL} />)}
          </div>
        </Card>
      </div>
    );
  }

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 24 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>Blog & Vodiči</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>Sve što trebate znati o hrvatskoj mirovini</p>
      </div>

      {/* Category filter */}
      <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
        {cats.map(c => (
          <button key={c} onClick={() => setFilter(c)} style={{
            padding: "6px 16px", border: `1.5px solid ${filter === c ? C.navy : C.border}`,
            borderRadius: 20, background: filter === c ? C.navy : C.white,
            color: filter === c ? "#fff" : C.slate, fontSize: 13, fontWeight: 500,
          }}>{c}</button>
        ))}
      </div>

      {/* Posts grid */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 32 }}>
        {filtered.map(p => (
          <Card key={p.id} style={{ cursor: "pointer", transition: "transform 0.15s, box-shadow 0.15s" }}
            onClick={() => setPost(p)}
            className="blog-card">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
              <Badge label={p.category} color={C.blue} />
              <span style={{ fontSize: 11, color: C.slateL }}>{p.readMin} min</span>
            </div>
            <h3 style={{ fontSize: 16, lineHeight: 1.35, marginBottom: 8 }}>{p.title}</h3>
            <p style={{ fontSize: 13, color: C.slateL, lineHeight: 1.6, marginBottom: 12 }}>{p.excerpt}</p>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                {p.tags.slice(0, 2).map(t => <Badge key={t} label={`#${t}`} color={C.slateL} />)}
              </div>
              <span style={{ fontSize: 12, color: C.blue, fontWeight: 600 }}>Čitaj →</span>
            </div>
          </Card>
        ))}
      </div>

      {/* ── S&P 500 vs HR mirovinski sustav ── */}
      <SP500MirovinaBlog />

      {/* AI Q&A */}
      <Card style={{ border: `2px solid ${C.gold}33` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
          <div style={{ background: C.amberL, borderRadius: 10, padding: "8px 10px" }}>
            <Star size={18} color={C.gold} />
          </div>
          <div>
            <h3 style={{ fontSize: 15, marginBottom: 2 }}>Pitaj AI stručnjaka</h3>
            <p style={{ fontSize: 12, color: C.slateL }}>Odgovori temeljeni na ZOMO, HANFA i HZMO izvorima</p>
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <input value={aiQ} onChange={e => setAiQ(e.target.value)}
            onKeyDown={e => e.key === "Enter" && askAI()}
            placeholder="npr. Kada mi je isplativo odgoditi mirovinu? Što je PVB? Kako funkcionira jednokratna isplata 20%?"
            style={{ flex: 1, padding: "11px 14px", border: `1.5px solid ${C.border}`, borderRadius: 9, fontSize: 13 }} />
          <button onClick={askAI} disabled={aiLoading} style={{
            padding: "11px 22px", background: aiLoading ? C.slateL : `linear-gradient(135deg, ${C.gold}, ${C.goldL})`,
            border: "none", borderRadius: 9, color: C.navy, fontWeight: 700, fontSize: 13,
            minWidth: 100,
          }}>
            {aiLoading ? "…" : "Pitaj"}
          </button>
        </div>
        {aiLoading && (
          <div style={{ padding: "14px", background: C.amberL, borderRadius: 9, fontSize: 13, color: C.amber }}>
            AI analizira vaše pitanje…
          </div>
        )}
        {aiA && (
          <div style={{ padding: "16px", background: C.bg, borderRadius: 9, fontSize: 13, lineHeight: 1.7, color: C.navy, whiteSpace: "pre-wrap" }}>
            {aiA}
          </div>
        )}
        <div style={{ marginTop: 10, fontSize: 11, color: C.slateL }}>
          Primjeri pitanja: "Kolika je kazna za prijevremenu mirovinu?", "Što je AVM i kako se usklađuje?", "Mogu li nasljediti II. stup?"
        </div>
      </Card>
    </div>
  );
}

// ─── EDUKACIJA ───────────────────────────────────────────────────────────────
const AVG_HR = 1300; // prosj. bruto plaća RH 2024.

function EduSection({ id, emoji, title, color, children }) {
  return (
    <div id={id} style={{ scrollMarginTop: 80, marginBottom: 48 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 20 }}>
        <div style={{ width: 44, height: 44, borderRadius: 12, background: color + "22",
          display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, flexShrink: 0 }}>
          {emoji}
        </div>
        <h2 style={{ fontSize: 22, color: C.navy, letterSpacing: "-0.02em" }}>{title}</h2>
      </div>
      {children}
    </div>
  );
}

function EduView() {
  const [section, setSection] = useState("pregled");

  // Interactive state per section
  const [staz, setStaz] = useState(35);
  const [pvb, setPvb] = useState(1.0);
  const [rani, setRani] = useState(0);
  const [odgoda, setOdgoda] = useState(0);
  const [plaça, setPlaça] = useState(1300);
  const [prinos, setPrinos] = useState(0.055);
  const [godStednje, setGodStednje] = useState(33);
  const [investPrinos, setInvestPrinos] = useState(0.07);
  const [investGodine, setInvestGodine] = useState(35);
  const [investMjesecno, setInvestMjesecno] = useState(200);

  // Derived live numbers
  const pf_rani = Math.max(0.7, 1 - 0.002 * rani);
  const pf_odg  = 1 + 0.0045 * odgoda;
  const pf_std  = 1.0;
  const mirovina_std  = calcStup1({ pvb, staz, pf: pf_std,  mf: 1, avm: AVM }).efektivna;
  const mirovina_rani = calcStup1({ pvb, staz, pf: pf_rani, mf: 1, avm: AVM }).efektivna;
  const mirovina_odg  = calcStup1({ pvb, staz, pf: pf_odg,  mf: 1, avm: AVM }).efektivna;

  // PVB sweep chart
  const pvbData = [0.5,0.7,0.8,1.0,1.2,1.5,2.0,2.5,3.0].map(p => ({
    pvb: p.toFixed(1), m: Math.round(calcStup1({ pvb: p, staz, pf: pf_std, mf: 1, avm: AVM }).efektivna)
  }));

  // Staž sweep chart
  const stazData = [15,20,25,30,35,40,45,50].map(s => ({
    staz: `${s}g`, m: Math.round(calcStup1({ pvb, staz: s, pf: pf_std, mf: 1, avm: AVM }).efektivna)
  }));

  // PF impact
  const pfData = [
    { name: "60 mj. ranije\n(−12%)", pf: 0.880 },
    { name: "36 mj. ranije\n(−7,2%)", pf: 0.928 },
    { name: "12 mj. ranije\n(−2,4%)", pf: 0.976 },
    { name: "Točno na dob\n(0%)", pf: 1.000 },
    { name: "+24 mj. odgode\n(+10,8%)", pf: 1.108 },
    { name: "+48 mj. odgode\n(+21,6%)", pf: 1.216 },
  ].map(d => ({ ...d, m: Math.round(calcStup1({ pvb, staz, pf: d.pf, mf: 1, avm: AVM }).efektivna) }));

  // II. stup akumulacija
  const kapital2 = calcStup2FV({ bruto_placa: plaça, god_stednje: godStednje, neto_prinos: prinos });
  const kapital2_konz = calcStup2FV({ bruto_placa: plaça, god_stednje: godStednje, neto_prinos: 0.035 });
  const kapital2_opt  = calcStup2FV({ bruto_placa: plaça, god_stednje: godStednje, neto_prinos: 0.075 });
  const renta2 = rentaDozivotna(kapital2 * 0.8);

  const akumData = Array.from({ length: godStednje + 1 }, (_, i) => {
    if (i % 5 !== 0 && i !== godStednje) return null;
    const k = calcStup2FV({ bruto_placa: plaça, god_stednje: i, neto_prinos: prinos });
    const uplate = plaça * 0.05 * i * 12;
    return { god: `${i}g`, kapital: Math.round(k), uplate: Math.round(uplate), prinos_eur: Math.round(k - uplate) };
  }).filter(Boolean);

  // REFORMA: privatna investicijska štednja
  // Isti novac koji ide u oba stupa (20% bruto) → privatni račun
  const uk_doprinos = plaça * 0.20; // umjesto 15%→drž. + 5%→OMF
  const invest_FV_uk = (() => {
    const pmt = uk_doprinos; const r = investPrinos / 12; const n = investGodine * 12;
    return r === 0 ? pmt * n : pmt * ((Math.pow(1+r,n)-1)/r);
  })();
  // Sadašnji sustav: I.stup + II.stup kapital
  const sadStup1 = mirovina_std * 12 * 20; // 20 god. isplate
  const sadStup2 = kapital2;
  const sadUkupno = sadStup1 + sadStup2;

  // Privatna štednja: samo doprinos 20% + prinos
  const privatnaMjesecno = investMjesecno;
  const privatnaFV = (() => {
    const pmt = privatnaMjesecno; const r = investPrinos / 12; const n = investGodine * 12;
    return r === 0 ? pmt * n : pmt * ((Math.pow(1+r,n)-1)/r);
  })();

  // Usporedba: 20% bruto → tržište vs. sadašnji sustav
  const reforma_comparison = [3,5,7,10].map(p => {
    const r = p / 100 / 12; const n = investGodine * 12; const pmt = uk_doprinos;
    const kapital = r === 0 ? pmt*n : pmt*((Math.pow(1+r,n)-1)/r);
    return { name: `${p}% prinos`, kapital: Math.round(kapital), renta: Math.round(kapital * 0.04 / 12) };
  });

  const navItems = [
    { id: "pregled", label: "Pregled sustava" },
    { id: "staz",    label: "📅 Mirovinski staž" },
    { id: "pvb",     label: "💰 Vrijednosni bodovi" },
    { id: "pf",      label: "⏩ Polazni faktor" },
    { id: "avm",     label: "📊 AVM" },
    { id: "stup2",   label: "🏦 II. stup" },
    { id: "prava",   label: "⚖️ Prava osiguranika" },
    { id: "reforma", label: "🔄 Reforma sustava" },
  ];

  const scrollTo = (id) => {
    setSection(id);
    document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const tw = (v) => `${(v/1000).toFixed(0)}k`;

  return (
    <div className="fadeIn" style={{ display: "grid", gridTemplateColumns: "200px 1fr", gap: 28, alignItems: "start" }}>

      {/* Sticky sidebar nav */}
      <div style={{ position: "sticky", top: 80, background: C.white, borderRadius: 14,
        border: `1px solid ${C.border}`, padding: "14px 0",
        boxShadow: "0 2px 12px rgba(26,46,82,0.07)" }}>
        <div style={{ padding: "0 14px 10px", fontSize: 10, fontWeight: 700, color: C.slateL,
          textTransform: "uppercase", letterSpacing: "0.08em" }}>Navigacija</div>
        {navItems.map(n => (
          <button key={n.id} onClick={() => scrollTo(n.id)} style={{
            width: "100%", padding: "8px 14px", border: "none", background: "none",
            textAlign: "left", fontSize: 12, fontWeight: section === n.id ? 700 : 400,
            color: section === n.id ? C.navy : C.slate,
            borderLeft: `3px solid ${section === n.id ? C.gold : "transparent"}`,
            cursor: "pointer", transition: "all 0.15s",
          }}>{n.label}</button>
        ))}
      </div>

      {/* Main content */}
      <div>
        {/* Hero */}
        <div style={{ background: `linear-gradient(135deg, ${C.navy}, #2a4080)`, borderRadius: 16,
          padding: "32px 36px", marginBottom: 36, color: "#fff", position: "relative", overflow: "hidden" }}>
          <div style={{ position: "absolute", right: -20, top: -20, fontSize: 120, opacity: 0.06 }}>🏛</div>
          <div style={{ fontSize: 11, fontWeight: 700, color: C.goldL, textTransform: "uppercase",
            letterSpacing: "0.1em", marginBottom: 8 }}>Edukacijski centar</div>
          <h1 style={{ fontSize: 28, marginBottom: 12, letterSpacing: "-0.02em", lineHeight: 1.2 }}>
            Kako funkcionira hrvatska mirovina?
          </h1>
          <p style={{ fontSize: 14, color: "rgba(255,255,255,0.75)", lineHeight: 1.7, maxWidth: 560 }}>
            Interaktivni vodiči za svaki parametar koji utječe na vašu mirovinu —
            s live grafovima i stvarnim primjerima. Shvatite formule, usporedite scenarije
            i procijenite potrebu za promjenom sustava.
          </p>
          <div style={{ display: "flex", gap: 20, marginTop: 20, flexWrap: "wrap" }}>
            {[
              { v: `${fmt(mirovina_std)} EUR`, l: "Vaša projicirana mirovina" },
              { v: `${fmt(kapital2)} EUR`,     l: "Akumulirani kapital II. stupa" },
              { v: `${pct(mirovina_std / plaça)}`, l: "Stopa zamjene" },
            ].map((it, i) => (
              <div key={i} style={{ background: "rgba(255,255,255,0.1)", borderRadius: 10, padding: "10px 16px" }}>
                <div style={{ fontSize: 18, fontWeight: 800, color: C.goldL, fontFamily: "'Playfair Display',serif" }}>{it.v}</div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)", marginTop: 2 }}>{it.l}</div>
              </div>
            ))}
          </div>
        </div>

        {/* ── PREGLED SUSTAVA ─────────────────────────────────────────── */}
        <EduSection id="pregled" emoji="🗺" title="Pregled mirovinskog sustava" color={C.navy}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginBottom: 24 }}>
            {[
              { n: "I. stup", emoji: "🏛", boja: C.navy, opis: "Generacijska solidarnost", detalj: "15% bruto plaće ide u državni fond. Sadašnji radnici financiraju sadašnje umirovljenike. Mirovina se izračunava formulom: PVB × Staž × PF × MF × AVM.", tag: "OBVEZNO · HZMO" },
              { n: "II. stup", emoji: "🏦", boja: C.purple, opis: "Individualna kapitalizirana štednja", detalj: "5% bruto plaće ide na vaš osobni račun u OMF-u. Tamo se ulaže i raste s tržišnim prinosima. Vaša osobna imovina.", tag: "OBVEZNO · REGOS / MOD" },
              { n: "III. stup", emoji: "🌱", boja: C.green, opis: "Dobrovoljna mirovinska štednja", detalj: "Vlastite uplate u dobrovoljni mirovinski fond. Država potiče s DPS-om do 99,54 EUR/god. Potpuna fleksibilnost.", tag: "DOBROVOLJNO" },
            ].map(s => (
              <div key={s.n} style={{ borderRadius: 14, border: `2px solid ${s.boja}22`,
                background: s.boja + "08", padding: 18, position: "relative" }}>
                <div style={{ fontSize: 28, marginBottom: 8 }}>{s.emoji}</div>
                <div style={{ fontSize: 11, fontWeight: 700, color: s.boja, background: s.boja + "18",
                  padding: "2px 8px", borderRadius: 10, display: "inline-block", marginBottom: 8 }}>{s.tag}</div>
                <h3 style={{ fontSize: 15, marginBottom: 4 }}>{s.n}</h3>
                <div style={{ fontSize: 12, fontWeight: 600, color: s.boja, marginBottom: 8 }}>{s.opis}</div>
                <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.6 }}>{s.detalj}</p>
              </div>
            ))}
          </div>

          {/* Formula prikaz */}
          <Card style={{ background: C.navy }}>
            <div style={{ color: C.goldL, fontSize: 11, fontWeight: 700, marginBottom: 12, textTransform: "uppercase", letterSpacing: "0.08em" }}>Formula I. stupa (ZOMO)</div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              {[
                { sym: "M", desc: "Mirovina (EUR/mj.)", color: "#7bffb0" },
                { sym: "=", desc: "", color: "#fff" },
                { sym: "PVB", desc: "Prosječni vrijednosni bodovi", color: "#7fbcff" },
                { sym: "×", desc: "", color: "#fff" },
                { sym: "Staž", desc: "Ukupne godine rada", color: "#ffb84d" },
                { sym: "×", desc: "", color: "#fff" },
                { sym: "PF", desc: "Polazni faktor (±odlazak)", color: "#c4a8ff" },
                { sym: "×", desc: "", color: "#fff" },
                { sym: "MF", desc: "Mirovinski faktor (vrsta)", color: "#ff9090" },
                { sym: "×", desc: "", color: "#fff" },
                { sym: "AVM", desc: "Aktualna vrijednost = 14,45 EUR", color: C.goldL },
              ].map((t, i) => t.sym === "=" || t.sym === "×"
                ? <span key={i} style={{ fontSize: 22, color: "#fff", fontWeight: 300 }}>{t.sym}</span>
                : (
                  <div key={i} style={{ textAlign: "center" }}>
                    <div style={{ fontSize: 20, fontWeight: 800, color: t.color, fontFamily: "'Playfair Display',serif" }}>{t.sym}</div>
                    <div style={{ fontSize: 9, color: "rgba(255,255,255,0.5)", maxWidth: 80, lineHeight: 1.3, marginTop: 2 }}>{t.desc}</div>
                  </div>
                )
              )}
            </div>
            <div style={{ marginTop: 14, padding: "10px 14px", background: "rgba(255,255,255,0.07)", borderRadius: 8 }}>
              <span style={{ fontSize: 12, color: "rgba(255,255,255,0.6)" }}>S vašim parametrima: </span>
              <span className="mono" style={{ fontSize: 12, color: "#7bffb0" }}>
                M = {pvb.toFixed(2)} × {staz} × {pf_std.toFixed(3)} × 1.00 × {AVM} = <strong>{fmt(mirovina_std)} EUR</strong>
              </span>
            </div>
          </Card>
        </EduSection>

        {/* ── MIROVINSKI STAŽ ─────────────────────────────────────────── */}
        <EduSection id="staz" emoji="📅" title="Mirovinski staž" color={C.blue}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
            <div>
              <Card style={{ marginBottom: 14 }}>
                <p style={{ fontSize: 13, lineHeight: 1.8, color: C.slate, marginBottom: 12 }}>
                  <strong>Mirovinski staž</strong> je ukupan broj godina koje ste proveli radeći uz uplatu doprinosa
                  u mirovinski sustav — bez obzira radi li se o punom ili nepunom radnom vremenu, vojnom roku,
                  bolovanju ili porodiljnom dopustu.
                </p>
                <p style={{ fontSize: 13, lineHeight: 1.8, color: C.slate, marginBottom: 12 }}>
                  U formuli <em>M = PVB × <strong>Staž</strong> × PF × MF × AVM</em>, staž je direktni množitelj —
                  svaka dodatna godina povećava mirovinu linearno za <strong>{fmt(pvb * AVM)} EUR/mj.</strong> (uz vaš PVB).
                </p>
                <div style={{ padding: "10px 14px", background: C.blueL, borderRadius: 9, fontSize: 12 }}>
                  <div style={{ fontWeight: 700, marginBottom: 4 }}>⚖️ Minimalni uvjeti:</div>
                  <ul style={{ paddingLeft: 16, display: "flex", flexDirection: "column", gap: 3, color: C.slate }}>
                    <li>Starosna mirovina: <strong>min. 15 god.</strong> staža + puna dob (65 M / ~64 Ž)</li>
                    <li>Prijevremena: <strong>min. 30 god.</strong> staža (muškarci), 28 god. (žene)</li>
                    <li>Dugogodišnji osiguranik: <strong>41 god.</strong> + 60 god. starosti</li>
                  </ul>
                </div>
              </Card>
              <Card>
                <MiniSlider label="Prilagodite staž" value={staz} min={15} max={50} step={1}
                  format={v => `${v} god.`} onChange={setStaz} color={C.blue} />
                <MiniSlider label="PVB (za kontekst)" value={pvb} min={0.5} max={2.5} step={0.05}
                  format={v => v.toFixed(2)} onChange={setPvb} color={C.navy} />
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 8 }}>
                  <div style={{ background: C.blueL, borderRadius: 9, padding: "10px 12px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, color: C.slateL }}>Efektivna mirovina</div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: C.blue, fontFamily: "'Playfair Display',serif" }}>{fmt(mirovina_std)} EUR</div>
                  </div>
                  <div style={{ background: C.amberL, borderRadius: 9, padding: "10px 12px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, color: C.slateL }}>Svaka god. staža vrijedi</div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: C.gold, fontFamily: "'Playfair Display',serif" }}>+{fmt(pvb * pf_std * AVM)} EUR</div>
                  </div>
                </div>
              </Card>
            </div>
            <Card>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 12 }}>Mirovina prema godinama staža (PVB = {pvb.toFixed(2)})</div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={stazData} margin={{ top: 5, right: 10, bottom: 5, left: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                  <XAxis dataKey="staz" tick={{ fontSize: 11 }} />
                  <YAxis tick={{ fontSize: 10 }} tickFormatter={v => `${v}`} />
                  <Tooltip formatter={v => [`${fmt(v)} EUR/mj.`, "Mirovina"]} />
                  <Bar dataKey="m" radius={[4,4,0,0]}>
                    {stazData.map((d,i) => <Cell key={i} fill={d.staz === `${staz}g` ? C.gold : C.blue + "88"} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
              <div style={{ marginTop: 10, fontSize: 11, color: C.slateL, lineHeight: 1.5 }}>
                💡 Graf pokazuje linearni rast: svaka godina dodaje isti iznos. Nema "bonusa" za dugi staž — ali nema ni penalizacije osim prijevremenog odlaska (sljedeće poglavlje).
              </div>
            </Card>
          </div>
        </EduSection>

        {/* ── VRIJEDNOSNI BODOVI ───────────────────────────────────────── */}
        <EduSection id="pvb" emoji="💰" title="Prosječni vrijednosni bodovi (PVB)" color={C.purple}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
            <div>
              <Card style={{ marginBottom: 14 }}>
                <p style={{ fontSize: 13, lineHeight: 1.8, color: C.slate, marginBottom: 10 }}>
                  PVB je ključan pojam koji opisuje <strong>kakva je vaša plaća bila u usporedbi s prosjekom RH</strong> — tijekom cijele karijere.
                </p>
                <div style={{ background: C.purpleL, borderRadius: 10, padding: "14px", marginBottom: 10 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.purple, marginBottom: 8 }}>Kako HZMO izračunava PVB:</div>
                  <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.8 }}>
                    <strong>1. Korak:</strong> Za svaku godinu rada: VB = vaša plaća ÷ prosječna plaća RH<br/>
                    <strong>2. Korak:</strong> PVB = prosjek svih godišnjih VB-ova<br/>
                    <strong>3. Korak:</strong> PVB ulazi u formulu: M = <strong>PVB</strong> × Staž × PF × MF × AVM
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                  {[
                    { plaça: 780,  pvb: "0,60", boja: C.red,    label: "Minimalac" },
                    { plaça: 1300, pvb: "1,00", boja: C.green,  label: "Prosjek RH" },
                    { plaça: 2600, pvb: "2,00", boja: C.purple, label: "Dvostruki prosjek" },
                  ].map(p => (
                    <div key={p.plaça} style={{ background: p.boja+"11", borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                      <div style={{ fontSize: 10, color: p.boja, fontWeight: 700 }}>{p.label}</div>
                      <div style={{ fontSize: 12, color: C.slate }}>{p.plaça} EUR/mj.</div>
                      <div style={{ fontSize: 16, fontWeight: 800, color: p.boja }}>PVB {p.pvb}</div>
                    </div>
                  ))}
                </div>
              </Card>
              <Card>
                <MiniSlider label="Prilagodite PVB" value={pvb} min={0.3} max={3.0} step={0.05}
                  format={v => `${v.toFixed(2)} (≈ ${Math.round(v * AVG_HR)} EUR bruto)`}
                  onChange={setPvb} color={C.purple} />
                <MiniSlider label="Staž (za kontekst)" value={staz} min={15} max={50} step={1}
                  format={v => `${v} god.`} onChange={setStaz} color={C.navy} />
                <div style={{ marginTop: 10, padding: "10px 12px", background: C.purpleL, borderRadius: 9 }}>
                  <div style={{ fontSize: 11, color: C.purple }}>Efektivna mirovina uz PVB {pvb.toFixed(2)}:</div>
                  <div style={{ fontSize: 24, fontWeight: 800, color: C.purple, fontFamily: "'Playfair Display',serif" }}>
                    {fmt(mirovina_std)} EUR/mj.
                  </div>
                  <div style={{ fontSize: 11, color: C.slateL, marginTop: 2 }}>
                    +0,1 PVB = +{fmt(calcStup1({pvb:pvb+0.1,staz,pf:1,mf:1,avm:AVM}).efektivna - mirovina_std)} EUR/mj. zauvijek
                  </div>
                </div>
              </Card>
            </div>
            <Card>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 12 }}>
                Mirovina prema PVB-u (Staž = {staz} god.)
              </div>
              <ResponsiveContainer width="100%" height={200}>
                <LineChart data={pvbData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                  <XAxis dataKey="pvb" tick={{ fontSize: 11 }} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip formatter={v => [`${fmt(v)} EUR/mj.`, "Mirovina"]} />
                  <Line type="monotone" dataKey="m" stroke={C.purple} strokeWidth={3} dot={{ r: 5, fill: C.purple }} />
                </LineChart>
              </ResponsiveContainer>
              <div style={{ marginTop: 14 }}>
                <div style={{ fontSize: 11, fontWeight: 600, color: C.slate, marginBottom: 8 }}>Što PVB znači u praksi</div>
                {[
                  { pvb_val: 0.6, opis: "Minimalna plaća cijele karijere", plaça_eur: "~780 EUR bruto" },
                  { pvb_val: 1.0, opis: "Uvijek prosječna plaća RH", plaça_eur: "~1.300 EUR bruto" },
                  { pvb_val: 1.5, opis: "50% iznad prosjeka", plaça_eur: "~1.950 EUR bruto" },
                  { pvb_val: 2.0, opis: "Visoka stručna plaća", plaça_eur: "~2.600 EUR bruto" },
                ].map(row => {
                  const m = calcStup1({ pvb: row.pvb_val, staz, pf: 1, mf: 1, avm: AVM }).efektivna;
                  return (
                    <div key={row.pvb_val} style={{ display: "flex", justifyContent: "space-between",
                      alignItems: "center", padding: "6px 10px", marginBottom: 4,
                      background: Math.abs(pvb - row.pvb_val) < 0.08 ? C.purpleL : C.bg, borderRadius: 7 }}>
                      <div>
                        <span style={{ fontSize: 12, fontWeight: 700, color: C.purple }}>PVB {row.pvb_val.toFixed(1)}</span>
                        <span style={{ fontSize: 11, color: C.slateL, marginLeft: 6 }}>{row.opis}</span>
                      </div>
                      <span style={{ fontSize: 12, fontWeight: 700, color: C.navy }}>{fmt(m)} EUR</span>
                    </div>
                  );
                })}
              </div>
            </Card>
          </div>
        </EduSection>

        {/* ── POLAZNI FAKTOR ───────────────────────────────────────────── */}
        <EduSection id="pf" emoji="⏩" title="Polazni faktor (PF) — kad odlazite u mirovinu" color={C.gold}>
          <Card style={{ marginBottom: 16, borderLeft: `4px solid ${C.gold}` }}>
            <p style={{ fontSize: 13, lineHeight: 1.8, color: C.slate }}>
              <strong>Polazni faktor (PF)</strong> je "nagrada ili kazna" za odluku kada ćete otići u mirovinu.
              Odlazite li <span style={{ color: C.red, fontWeight: 700 }}>ranije</span> od zakonske dobi — PF je manji od 1,0
              i vaša mirovina je trajno niža. Odlazite li <span style={{ color: C.green, fontWeight: 700 }}>kasnije</span>
              — PF je veći od 1,0 i mirovina je trajno viša. Ovo povećanje/smanjenje je <strong>permanentno</strong> —
              ne mijenja se ni kad napunite "pravu" dob.
            </p>
          </Card>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
            <div>
              <Card style={{ marginBottom: 14 }}>
                <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 12 }}>Prijevremeni odlazak</div>
                <MiniSlider label="Mjeseci ranije (0 = bez kazne)" value={rani} min={0} max={60} step={1}
                  format={v => v === 0 ? "Puna dob (bez kazne)" : `${v} mj. ranije`}
                  onChange={setRani} color={C.red} />
                <div style={{ padding: "10px 12px", background: C.redL, borderRadius: 9, fontSize: 12 }}>
                  <div>PF = 1,00 − 0,002 × {rani} = <strong style={{ color: C.red }}>{pf_rani.toFixed(3)}</strong></div>
                  <div style={{ marginTop: 4, color: C.red }}>Trajno smanjenje: <strong>−{((1-pf_rani)*100).toFixed(1)}%</strong></div>
                  <div style={{ marginTop: 4 }}>Mirovina: {fmt(mirovina_std)} EUR → <strong style={{ color: C.red }}>{fmt(mirovina_rani)} EUR</strong> (−{fmt(mirovina_std-mirovina_rani)} EUR/mj.)</div>
                </div>
              </Card>
              <Card>
                <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 12 }}>Odgoda iznad zakonske dobi</div>
                <MiniSlider label="Dodatni mjeseci rada" value={odgoda} min={0} max={60} step={1}
                  format={v => v === 0 ? "Bez odgode" : `+${v} mj.`}
                  onChange={setOdgoda} color={C.green} />
                <div style={{ padding: "10px 12px", background: C.greenL, borderRadius: 9, fontSize: 12 }}>
                  <div>PF = 1,00 + 0,0045 × {odgoda} = <strong style={{ color: C.green }}>{pf_odg.toFixed(3)}</strong></div>
                  <div style={{ marginTop: 4, color: C.green }}>Trajno povećanje: <strong>+{((pf_odg-1)*100).toFixed(1)}%</strong></div>
                  <div style={{ marginTop: 4 }}>Mirovina: {fmt(mirovina_std)} EUR → <strong style={{ color: C.green }}>{fmt(mirovina_odg)} EUR</strong> (+{fmt(mirovina_odg-mirovina_std)} EUR/mj.)</div>
                </div>
              </Card>
            </div>
            <Card>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 12 }}>
                Utjecaj odluke o odlasku na mirovinu (Staž={staz}, PVB={pvb.toFixed(2)})
              </div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={pfData} margin={{ top: 5, right: 10, bottom: 60, left: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                  <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-15} textAnchor="end" interval={0} />
                  <YAxis tick={{ fontSize: 10 }} domain={[0, 'auto']} />
                  <Tooltip formatter={v => [`${fmt(v)} EUR/mj.`, "Mirovina"]} />
                  <Bar dataKey="m" radius={[4,4,0,0]}>
                    {pfData.map((d,i) => (
                      <Cell key={i} fill={d.pf < 1 ? `hsl(${Math.round(d.pf*360-240)},70%,50%)` : d.pf === 1 ? C.slate : C.green} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginTop: 10 }}>
                {[
                  { label: "Rani odlazak (−36 mj.)", m: mirovina_rani, color: C.red },
                  { label: "Standardni odlazak", m: mirovina_std, color: C.slate },
                  { label: "Odgoda (+36 mj.)", m: calcStup1({pvb,staz,pf:1+0.0045*36,mf:1,avm:AVM}).efektivna, color: C.green },
                ].map((it, i) => (
                  <div key={i} style={{ background: it.color+"11", borderRadius: 9, padding: "8px 10px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, color: it.color, fontWeight: 600, lineHeight: 1.3 }}>{it.label}</div>
                    <div style={{ fontSize: 16, fontWeight: 800, color: it.color, fontFamily: "'Playfair Display',serif" }}>{fmt(it.m)}</div>
                    <div style={{ fontSize: 9, color: C.slateL }}>EUR/mj.</div>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        </EduSection>

        {/* ── AVM ─────────────────────────────────────────────────────── */}
        <EduSection id="avm" emoji="📊" title="Aktualna vrijednost mirovine (AVM)" color={C.green}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
            <Card>
              <p style={{ fontSize: 13, lineHeight: 1.8, color: C.slate, marginBottom: 14 }}>
                <strong>AVM</strong> je "cijena jednog osobnog boda" — iznos koji Vlada RH određuje i koji
                izravno množi sve vaše bodove. Vrijedi jednako za sve umirovljenike u sustavu.
              </p>
              <div style={{ background: C.greenL, borderRadius: 10, padding: 14, marginBottom: 14 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.green, marginBottom: 8 }}>Kretanje AVM-a (odabrane godine):</div>
                {[
                  { god: "2019.", avm: 10.18 }, { god: "2021.", avm: 10.73 }, { god: "2023.", avm: 12.34 },
                  { god: "2024.", avm: 13.57 }, { god: "1.7.2025.", avm: 14.45, curr: true },
                ].map(row => (
                  <div key={row.god} style={{ display: "flex", justifyContent: "space-between",
                    padding: "4px 8px", background: row.curr ? C.green+"22" : "transparent",
                    borderRadius: 6, marginBottom: 2 }}>
                    <span style={{ fontSize: 12, color: C.slate }}>{row.god}</span>
                    <span style={{ fontSize: 12, fontWeight: row.curr ? 800 : 500, color: row.curr ? C.green : C.navy }}>
                      {row.avm.toFixed(2)} EUR {row.curr && "← Sadašnji"}
                    </span>
                  </div>
                ))}
              </div>
              <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.7 }}>
                <strong>Kako se usklađuje (Swiss index, od ZMO 2025.):</strong>
                <ul style={{ paddingLeft: 16, marginTop: 6 }}>
                  <li>2× godišnje (1. veljače i 1. kolovoza)</li>
                  <li>Formula: 85% × rast plaća + 15% × inflacija (ili obrnuto, što je povoljnije)</li>
                  <li>Garantira da mirovine prate ekonomski razvoj</li>
                </ul>
              </div>
            </Card>
            <Card>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 14 }}>
                Što znači promjena AVM-a za mirovinu?
              </div>
              {[10, 11, 12, 13, AVM, 16, 18, 20].map(a => {
                const m = calcStup1({ pvb, staz, pf: 1, mf: 1, avm: a }).efektivna;
                const pct_diff = ((a - AVM) / AVM * 100).toFixed(0);
                const isNow = Math.abs(a - AVM) < 0.1;
                return (
                  <div key={a} style={{ display: "flex", alignItems: "center", gap: 8,
                    padding: "6px 10px", marginBottom: 4,
                    background: isNow ? C.navy+"11" : C.bg, borderRadius: 8,
                    border: isNow ? `1.5px solid ${C.navy}33` : "none" }}>
                    <div style={{ width: 52, fontSize: 12, fontWeight: isNow ? 800 : 500,
                      color: isNow ? C.navy : C.slate }}>{a.toFixed(2)} EUR</div>
                    <div style={{ flex: 1, height: 8, borderRadius: 4, background: C.border, overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${(m / 1200 * 100)}%`,
                        background: isNow ? C.navy : C.green+"88", transition: "width 0.3s" }} />
                    </div>
                    <div style={{ width: 70, fontSize: 12, fontWeight: 700,
                      color: isNow ? C.navy : a > AVM ? C.green : C.slateL, textAlign: "right" }}>
                      {fmt(m)} EUR
                    </div>
                    {isNow && <span style={{ fontSize: 9, color: C.navy, fontWeight: 700 }}>← SADA</span>}
                    {!isNow && a !== AVM && <span style={{ fontSize: 9, color: a > AVM ? C.green : C.slateL, width: 30 }}>{pct_diff > 0 ? "+" : ""}{pct_diff}%</span>}
                  </div>
                );
              })}
            </Card>
          </div>
        </EduSection>

        {/* ── II. STUP ─────────────────────────────────────────────────── */}
        <EduSection id="stup2" emoji="🏦" title="II. stup — Individualna kapitalizirana štednja" color={C.purple}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
            <div>
              <Card style={{ marginBottom: 14 }}>
                <p style={{ fontSize: 13, lineHeight: 1.8, color: C.slate, marginBottom: 12 }}>
                  Svaki zaposleni odvaja <strong>5% bruto plaće</strong> na osobni račun u Obveznom mirovinskom fondu (OMF).
                  Ta sredstva su vaša osobna imovina — ne idu u "zajednički lonac" nego direktno na vaš račun.
                </p>
                <div style={{ background: C.purpleL, borderRadius: 10, padding: 14, marginBottom: 10 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.purple, marginBottom: 8 }}>Formula akumulacije:</div>
                  <div className="mono" style={{ fontSize: 13, color: C.navy }}>
                    FV = PMT × [((1+r)ⁿ − 1) / r]
                  </div>
                  <div style={{ fontSize: 11, color: C.slateL, marginTop: 6 }}>
                    PMT = uplata/mj. (5% bruto) · r = prinos/12 · n = broj mj. štednje
                  </div>
                </div>
                <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.7 }}>
                  <strong>Ključna razlika od I. stupa:</strong> Vaš kapital se <em>nasljeđuje</em>.
                  Akumulirana sredstva su predmet nasljeđivanja ako umrete prije umirovljenja,
                  a obitelj nema pravo na obiteljsku mirovinu (čl. 103. ZOMF).
                </div>
              </Card>
              <Card>
                <MiniSlider label="Bruto plaća" value={plaça} min={800} max={4000} step={50}
                  format={v => `${v.toLocaleString("hr-HR")} EUR/mj.`} onChange={setPlaça} color={C.purple} />
                <MiniSlider label="Neto prinos" value={prinos} min={0.02} max={0.10} step={0.005}
                  format={v => `${(v*100).toFixed(1)}%`} onChange={setPrinos} color={C.purple} />
                <MiniSlider label="Godine štednje" value={godStednje} min={5} max={45} step={1}
                  format={v => `${v} god.`} onChange={setGodStednje} color={C.purple} />
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginTop: 8 }}>
                  {[
                    { l: "Konz. (3,5%)", v: kapital2_konz, c: C.amber },
                    { l: "Umjer. (5,5%)", v: kapital2, c: C.purple },
                    { l: "Opt. (7,5%)", v: kapital2_opt, c: C.green },
                  ].map(it => (
                    <div key={it.l} style={{ background: it.c+"11", borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                      <div style={{ fontSize: 9, color: it.c, fontWeight: 600 }}>{it.l}</div>
                      <div style={{ fontSize: 14, fontWeight: 800, color: it.c, fontFamily: "'Playfair Display',serif" }}>{tw(it.v)} EUR</div>
                    </div>
                  ))}
                </div>
              </Card>
            </div>
            <Card>
              <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 12 }}>
                Rast kapitala kroz {godStednje} god. štednje
              </div>
              <ResponsiveContainer width="100%" height={200}>
                <AreaChart data={akumData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                  <XAxis dataKey="god" tick={{ fontSize: 11 }} />
                  <YAxis tick={{ fontSize: 10 }} tickFormatter={tw} />
                  <Tooltip formatter={v => [fmtEur(v), ""]} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Area type="monotone" dataKey="uplate" stackId="1" fill={C.blueL} stroke={C.blue} name="Čiste uplate" />
                  <Area type="monotone" dataKey="prinos_eur" stackId="1" fill={C.purpleL} stroke={C.purple} name="Investicijski prinos" />
                </AreaChart>
              </ResponsiveContainer>
              <div style={{ marginTop: 14, padding: "10px 12px", background: C.purpleL, borderRadius: 9 }}>
                <div style={{ fontSize: 11, color: C.purple, fontWeight: 700 }}>Akumulirani kapital: <span style={{ fontSize: 18, fontFamily: "'Playfair Display',serif" }}>{fmt(kapital2)} EUR</span></div>
                <div style={{ fontSize: 11, color: C.slateL, marginTop: 4 }}>
                  Od toga: {fmt(plaça*0.05*godStednje*12)} EUR uplate + {fmt(kapital2 - plaça*0.05*godStednje*12)} EUR investicijskog prinosa
                </div>
                <div style={{ fontSize: 11, color: C.purple, marginTop: 4 }}>
                  Okvirna doživotna renta (80%): <strong>{fmt(renta2)} EUR/mj.</strong>
                </div>
              </div>
            </Card>
          </div>
        </EduSection>

        {/* ── REFORMA ──────────────────────────────────────────────────── */}
        <EduSection id="reforma" emoji="🔄" title="Trebamo li reformu mirovinskog sustava?" color={C.red}>

          {/* Impact statement */}
          <div style={{ background: `linear-gradient(135deg, #1a0a0a, #3a1515)`, borderRadius: 16,
            padding: "28px 32px", marginBottom: 24, color: "#fff", position: "relative", overflow: "hidden" }}>
            <div style={{ position: "absolute", right: 10, top: 10, fontSize: 80, opacity: 0.06 }}>⚠</div>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#ffaaaa", textTransform: "uppercase",
              letterSpacing: "0.1em", marginBottom: 10 }}>Demografska matematika</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 20, marginBottom: 16 }}>
              {[
                { stat: "1,41 : 1", desc: "Omjer zaposlenih i umirovljenika u 2025.", src: "HZMO, REGOS" },
                { stat: "~38%", desc: "Prosječna stopa zamjene I. stupa (% bruto plaće)", src: "OECD, HZMO" },
                { stat: "< 1,0 : 1", desc: "Projicirani omjer do 2050. (DZS demografske projekcije)", src: "DZS, Eurostat" },
              ].map(it => (
                <div key={it.stat}>
                  <div style={{ fontSize: 28, fontWeight: 900, color: "#ff9090", fontFamily: "'Playfair Display',serif", lineHeight: 1 }}>{it.stat}</div>
                  <div style={{ fontSize: 12, color: "rgba(255,255,255,0.7)", marginTop: 6, lineHeight: 1.5 }}>{it.desc}</div>
                  <div style={{ fontSize: 10, color: "#ff9090", marginTop: 4, opacity: 0.7 }}>Izvor: {it.src}</div>
                </div>
              ))}
            </div>
            <p style={{ fontSize: 13, color: "rgba(255,255,255,0.8)", lineHeight: 1.7, maxWidth: 640 }}>
              Generacijska solidarnost funkcionira dok ima dovoljno radnika koji financiraju umirovljenike.
              S omjerom od samo <strong style={{ color: "#ffaaaa" }}>1,41 radnika po umirovljeniku</strong> i
              demografskim trendom starenja, matematika govori jasno: <strong style={{ color: "#ffaaaa" }}>
              budući umirovljenici ne mogu računati na isti standard od I. stupa kao sadašnji.</strong>
            </p>
          </div>

          {/* Simulator: sadašnji vs. privatni */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 24 }}>
            <Card style={{ borderTop: `4px solid ${C.navy}` }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, marginBottom: 14 }}>
                🏛 Sadašnji miješani sustav (I. + II. stup)
              </div>
              <MiniSlider label="Bruto plaća" value={plaça} min={800} max={4000} step={50}
                format={v => `${v.toLocaleString("hr-HR")} EUR`} onChange={setPlaça} color={C.navy} />
              <MiniSlider label="Staž" value={staz} min={20} max={50} step={1}
                format={v => `${v} god.`} onChange={setStaz} color={C.navy} />
              {(() => {
                const m1 = calcStup1({ pvb: plaça / AVG_HR, staz, pf: 1, mf: 1, avm: AVM }).efektivna;
                const k2 = calcStup2FV({ bruto_placa: plaça, god_stednje: staz, neto_prinos: prinos });
                const r2 = rentaDozivotna(k2 * 0.8);
                const dod = m1 * 0.27;
                const uk = m1 + dod + r2;
                const stopaz = uk / plaça;
                return (
                  <div style={{ marginTop: 8 }}>
                    {[
                      { l: "I. stup (osnovna + dodatak 27%)", v: m1 + dod, c: C.navy },
                      { l: "II. stup doživotna renta (80% kap.)", v: r2, c: C.purple },
                      { l: "UKUPNO / mj.", v: uk, c: C.green, big: true },
                      { l: "Stopa zamjene", v: null, pv: stopaz, c: C.gold },
                    ].map((it, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between",
                        padding: "6px 10px", marginBottom: 4,
                        background: it.big ? C.navy : C.bg, borderRadius: 8 }}>
                        <span style={{ fontSize: 12, color: it.big ? "#fff" : C.slate }}>{it.l}</span>
                        <strong style={{ fontSize: it.big ? 15 : 12, color: it.big ? "#7bffb0" : it.c }}>
                          {it.pv !== undefined ? pct(it.pv) : fmtEur(it.v)}
                        </strong>
                      </div>
                    ))}
                  </div>
                );
              })()}
            </Card>

            <Card style={{ borderTop: `4px solid ${C.green}` }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.green, marginBottom: 4 }}>
                🌱 Hipotetski: privatna investicijska štednja
              </div>
              <div style={{ fontSize: 11, color: C.slateL, marginBottom: 14, lineHeight: 1.5 }}>
                Što da je svih 20% doprinosa (15%+5%) išlo na privatni investicijski račun?
              </div>
              <MiniSlider label="Godišnji prinos investicija" value={investPrinos} min={0.03} max={0.12} step={0.005}
                format={v => `${(v*100).toFixed(1)}%`} onChange={setInvestPrinos} color={C.green} />
              {(() => {
                const pmt = plaça * 0.20; // 20% bruto
                const r = investPrinos / 12;
                const n = staz * 12;
                const kapital = r === 0 ? pmt*n : pmt*((Math.pow(1+r,n)-1)/r);
                const renta_4pct = kapital * 0.04 / 12;
                const renta_3pct = kapital * 0.03 / 12;
                const dob_mj = rentaDozivotna(kapital);
                const stopaz = renta_4pct / plaça;
                return (
                  <div>
                    {[
                      { l: "Akumulirani kapital", v: kapital, c: C.green },
                      { l: "Doživotna renta (3% god. – konzervativno)", v: renta_3pct, c: C.navy },
                      { l: "Renta po pravilu 4% (optimistično)", v: renta_4pct, c: C.green, big: true },
                      { l: "Stopa zamjene (4% pravilo)", v: null, pv: stopaz, c: C.gold },
                    ].map((it, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between",
                        padding: "6px 10px", marginBottom: 4,
                        background: it.big ? C.green : C.bg, borderRadius: 8 }}>
                        <span style={{ fontSize: 12, color: it.big ? "#fff" : C.slate }}>{it.l}</span>
                        <strong style={{ fontSize: it.big ? 15 : 12, color: it.big ? "#fff" : it.c }}>
                          {it.pv !== undefined ? pct(it.pv) : fmtEur(it.v)}
                        </strong>
                      </div>
                    ))}
                    <div style={{ marginTop: 10, fontSize: 11, color: C.slateL, lineHeight: 1.5,
                      padding: "8px 10px", background: C.greenL, borderRadius: 8 }}>
                      ⚠ Ovo je simulacija — privatna štednja nosi <strong>investicijski rizik</strong>.
                      Kapital može pasti s tržištem. Sustav bi trebao uključivati zaštitni mehanizam.
                    </div>
                  </div>
                );
              })()}
            </Card>
          </div>

          {/* Reforma argumenti */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
            <Card style={{ borderLeft: `4px solid ${C.red}` }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.red, marginBottom: 12 }}>
                ⚠ Problemi sadašnjeg sustava
              </div>
              {[
                { t: "Demografska vremenita bomba", d: "Omjer radnik/umirovljenik pada. 2050. bit će manje od 1 radnika po umirovljeniku. I. stup tada neće biti održiv bez drastičnog smanjenja mirovina ili povećanja poreza." },
                { t: "Niska stopa zamjene", d: "Hrvatska I. stup daje prosječno ~38% bruto plaće. OECD prosjek je 49%, Danska 80%. Dodavanjem II. stupa dolazimo do ~50%, ali to nije dovoljno za dostojanstvenu starost." },
                { t: "Međugeneracijska nepravda", d: "Mlađe generacije plaćaju visoke doprinose za sadašnje umirovljenike, a same neće imati dovoljno uplatitelja kada dođu u starost. Sustav je asimetričan na njihovu štetu." },
                { t: "Kapital bez nasljeđivanja", d: "U standardnoj doživotnoj renti iz II. stupa, preostali kapital ostaje MOD-u, ne nasljednicima. Generacijski kapital se 'gubi' u sustavu, ne prenosi na djecu." },
              ].map((it, i) => (
                <div key={i} style={{ marginBottom: 10, paddingBottom: 10,
                  borderBottom: i < 3 ? `1px solid ${C.border}` : "none" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, marginBottom: 3 }}>{it.t}</div>
                  <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.6 }}>{it.d}</div>
                </div>
              ))}
            </Card>

            <Card style={{ borderLeft: `4px solid ${C.green}` }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.green, marginBottom: 12 }}>
                🌱 Argumenti za privatizaciju štednje
              </div>
              {[
                { t: "Kapital koji raste s ekonomijom", d: "Investicijska štednja prati rast tržišta. Globalni dionički indeksi dugoročno rastu 7–10% godišnje nominalno. Uz složenu kamatu, 35 godina štednje stvara supstancijalan kapital." },
                { t: "Nasljednost i međugeneracijski kapital", d: "Privatna štednja je osobna imovina. Može se nasljeđivati — obitelj gradi kapital kroz generacije, a ne gubi ga u državnom pulu." },
                { t: "Neovisnost od demografije", d: "Privatna štednja ne ovisi o omjeru radnika i umirovljenika. Svaka osoba sama financira svoju starost — sustav ne može 'bankrotirati' zbog promjene demografije." },
                { t: "Primjeri iz prakse", d: "Čile (1981.), Australija (superannuation), Singapur (CPF) — sustavi s visokim udjelom kapitalzirane štednje pokazuju bolje ishode stopa zamjene uz veće akumulirani kapital." },
              ].map((it, i) => (
                <div key={i} style={{ marginBottom: 10, paddingBottom: 10,
                  borderBottom: i < 3 ? `1px solid ${C.border}` : "none" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, marginBottom: 3 }}>{it.t}</div>
                  <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.6 }}>{it.d}</div>
                </div>
              ))}
            </Card>
          </div>

          {/* Live usporedba po prinosu */}
          <Card>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <div style={{ fontSize: 13, fontWeight: 700 }}>
                Usporedba: Privatna štednja (20% bruto) pri različitim prinosima
              </div>
              <div style={{ fontSize: 11, color: C.slateL }}>Plaća {plaça.toLocaleString("hr-HR")} EUR · {staz} god.</div>
            </div>
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={reforma_comparison} margin={{ top: 5, right: 20, bottom: 5, left: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis yAxisId="left" tick={{ fontSize: 10 }} tickFormatter={tw} />
                <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 10 }} />
                <Tooltip formatter={(v, n) => [n === "kapital" ? fmtEur(v) : `${fmt(v)} EUR/mj.`, n === "kapital" ? "Akumulirani kapital" : "Renta (4% pravilo)"]} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar yAxisId="left" dataKey="kapital" name="kapital" radius={[4,4,0,0]}>
                  {reforma_comparison.map((d,i) => <Cell key={i} fill={[C.amber,C.gold,C.green,C.blue][i]} />)}
                </Bar>
                <Bar yAxisId="right" dataKey="renta" name="renta" radius={[4,4,0,0]} fill={C.navy+"66"} />
              </BarChart>
            </ResponsiveContainer>
            <div style={{ marginTop: 14, padding: "12px 16px", background: C.bg, borderRadius: 10,
              fontSize: 12, color: C.slate, lineHeight: 1.7 }}>
              Uz prinos od 5–7% (realni dugoročni prosjek globalnih tržišta),
              potpuno privatizirana štednja pri istim doprinosima (20% bruto) generira
              <strong> 2–4× veći kapital</strong> od sadašnjeg miješanog sustava i
              <strong> 60–120% višu stopu zamjene</strong> — uz uvjet da tržišta dugoročno rastu.
              Ključni rizik privatizacije: volatilnost tržišta, nedostatak zajamčene isplate i
              izostanak zaštite od invalidnosti. Koji omjer privatnog i javnog
              komponente mirovine je prihvatljiv — svaki korisnik procjenjuje sam, prema
              vlastitim prioritetima i toleranciji na rizik.
            </div>
          </Card>

          <div style={{ marginTop: 20, padding: "16px 20px", background: C.amberL,
            borderRadius: 12, borderLeft: `4px solid ${C.gold}` }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.amber, marginBottom: 6 }}>
              ⚠ Napomena o reformskim prijedlozima
            </div>
            <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.7 }}>
              Ove analize su edukativne i ne predstavljaju političku preporuku. Svaki reformski model
              nosi kompromise: privatizacija rješava demografski problem ali povećava tržišni rizik.
              Čileanski model pokazao je i probleme (niske mirovine za povremene radnike i žene).
              Australski i singapurski model kombinirani s minimalnom državnom garantiranom mirovinom
              daju uravnoteženije ishode. Svaka reforma u HR zahtijevala bi dugačak tranzicijski period
              i zaštitu prava stečenih u sadašnjem sustavu.
            </p>
          </div>
        </EduSection>

        {/* ── PRAVA OSIGURANIKA ─────────────────────────────────────── */}
        <EduSection id="prava" emoji="⚖️" title="Prava osiguranika u II. stupu" color={C.blue}>

          {/* Hero pojašnjenje */}
          <div style={{ background: `linear-gradient(135deg, #0f2040, #1a3566)`, borderRadius: 14,
            padding: "24px 28px", marginBottom: 22, color: "#fff" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: C.goldL, textTransform: "uppercase",
              letterSpacing: "0.08em", marginBottom: 10 }}>Ključna razlika od I. stupa</div>
            <p style={{ fontSize: 14, color: "rgba(255,255,255,0.88)", lineHeight: 1.8, maxWidth: 640 }}>
              Sredstva u II. stupu su <strong style={{ color: "#7bffb0" }}>vaša osobna imovina</strong> —
              ne idu u zajednički državni fond, nego na vaš individualni račun koji se vodi u
              Središnjem registru osiguranika (<strong style={{ color: C.goldL }}>REGOS</strong>).
              Ni Vlada ni HZMO ne mogu ih koristiti za isplatu tuđih mirovina.
            </p>
            <div style={{ display: "flex", gap: 16, marginTop: 16, flexWrap: "wrap" }}>
              {[
                { emoji: "🔒", label: "Osobna imovina", sub: "Ne može se oduzeti, blokirati ni koristiti za tuđe potrebe" },
                { emoji: "📋", label: "Neotuđiva", sub: "Ne možete je prodati, zamijeniti ni prenijeti trećima" },
                { emoji: "👨‍👩‍👧", label: "Nasljedna*", sub: "Uz zakonske uvjete (detalji niže)" },
                { emoji: "🔍", label: "Transparentna", sub: "Stanje računa vidljivo u REGOS portalu i e-Građani" },
              ].map(it => (
                <div key={it.emoji} style={{ background: "rgba(255,255,255,0.09)", borderRadius: 10,
                  padding: "10px 14px", flex: "1 1 160px" }}>
                  <div style={{ fontSize: 20, marginBottom: 4 }}>{it.emoji}</div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#7fbcff", marginBottom: 3 }}>{it.label}</div>
                  <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)", lineHeight: 1.4 }}>{it.sub}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Tko su akteri — vizualni flow */}
          <Card style={{ marginBottom: 18 }}>
            <h3 style={{ fontSize: 14, marginBottom: 16, color: C.navy }}>Tko čini sustav II. stupa i koja je uloga svakog aktera?</h3>
            <div style={{ display: "flex", alignItems: "stretch", gap: 0, overflowX: "auto" }}>
              {[
                { emoji: "👷", naziv: "Vi (Osiguranik)", boja: C.blue,
                  opis: "Vlasnik sredstava. Platite 5% bruto — poslodavac to uplaćuje u vaše ime. Vi birate fond i kategoriju.",
                  obveze: ["Odaberite OMF pri zaposlenju", "Pratite stanje na e-Građani / REGOS portalu", "Pri umirovljenju odaberite MOD"] },
                { emoji: "🏛", naziv: "REGOS", boja: C.navy,
                  opis: "Središnji registar osiguranika. Vodi evidenciju svih osobnih računa, prima uplate od HZMO-a i raspoređuje ih u fondove.",
                  obveze: ["Vodi vaš osobni račun", "Rasporedit će vas u fond A ako ne odaberete", "Prebacuje sredstva MOD-u pri umirovljenju"] },
                { emoji: "📈", naziv: "Mirovinsko društvo (OMF)", boja: C.purple,
                  opis: "Profesionalni upravitelj fonda. Ulaže vaša sredstva u financijske instrumente. Pod nadzorom HANFA-e.",
                  obveze: ["Ulaže novac prema strategiji fonda", "Naplaćuje naknade (max 0,24% god.)", "Mora imati jamstveni polog"] },
                { emoji: "🏦", naziv: "MOD", boja: C.green,
                  opis: "Mirovinsko osiguravajuće društvo. Od vaših akumuliranih sredstava formira i isplaćuje doživotnu rentu.",
                  obveze: ["Sklapa ugovor o mirovini s vama", "Isplaćuje doživotnu rentu", "Preuzima dugovječnosni rizik"] },
                { emoji: "🔍", naziv: "HANFA", boja: C.gold,
                  opis: "Regulator i nadzornik. Kontrolira zakonitost poslovanja svih mirovinskih društava i MOD-ova.",
                  obveze: ["Nadzire OMF i MOD", "Propisuje limite ulaganja", "Provodi inspekcije i kazne"] },
              ].map((a, i) => (
                <div key={i} style={{ display: "flex", alignItems: "stretch", flex: "none" }}>
                  <div style={{ width: 190, padding: "14px 14px", background: a.boja + "0a",
                    border: `1.5px solid ${a.boja}22`, borderRadius: 10, margin: "0 4px" }}>
                    <div style={{ fontSize: 24, marginBottom: 6 }}>{a.emoji}</div>
                    <div style={{ fontSize: 12, fontWeight: 800, color: a.boja, marginBottom: 6 }}>{a.naziv}</div>
                    <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5, marginBottom: 10 }}>{a.opis}</div>
                    <div style={{ fontSize: 10, color: C.slateL, fontWeight: 600, marginBottom: 4 }}>ZADUŽENJA:</div>
                    {a.obveze.map((o, j) => (
                      <div key={j} style={{ fontSize: 10, color: C.slate, lineHeight: 1.4,
                        padding: "3px 0", borderBottom: j < a.obveze.length-1 ? `1px solid ${C.border}` : "none" }}>
                        › {o}
                      </div>
                    ))}
                  </div>
                  {i < 4 && (
                    <div style={{ display: "flex", alignItems: "center", color: C.slateL,
                      fontSize: 18, padding: "0 2px", flexShrink: 0 }}>→</div>
                  )}
                </div>
              ))}
            </div>
          </Card>

          {/* MOD detaljno */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
            <Card style={{ borderLeft: `4px solid ${C.green}` }}>
              <h3 style={{ fontSize: 14, color: C.green, marginBottom: 14 }}>🏦 Što je MOD i kako upravlja vašim novcem?</h3>
              <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.7, marginBottom: 12 }}>
                <strong>MOD (Mirovinsko osiguravajuće društvo)</strong> je privatna financijska institucija
                koja od vaših akumuliranih sredstava u OMF-u formira doživotnu mirovinu.
                U HR trenutno postoje <strong>4 registrirana MOD-a</strong>:
              </p>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
                {[
                  { n: "HRMOD d.d.", v: "Hrvatska poštanska banka" },
                  { n: "Erste Mirovinski", v: "Erste Gruppe" },
                  { n: "PBZ CO d.o.o.", v: "Privredna banka Zagreb" },
                  { n: "Croatia MOD", v: "Croatia osiguranje" },
                ].map(m => (
                  <div key={m.n} style={{ display: "flex", justifyContent: "space-between",
                    padding: "5px 10px", background: C.greenL, borderRadius: 7, fontSize: 11 }}>
                    <strong style={{ color: C.navy }}>{m.n}</strong>
                    <span style={{ color: C.slateL }}>{m.v}</span>
                  </div>
                ))}
              </div>
              <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.7 }}>
                <p style={{ marginBottom: 8 }}>
                  <strong>Kako MOD izračunava rentu?</strong> Aktuarskim modelom uzima u obzir:
                </p>
                <ul style={{ paddingLeft: 16, display: "flex", flexDirection: "column", gap: 3 }}>
                  <li>Vaš akumulirani kapital iz OMF-a</li>
                  <li>Spol i dob pri umirovljenju (aktuarski tablice)</li>
                  <li>Odabranu vrstu mirovine (doživotna, zajednička, zajamčeno razdoblje)</li>
                  <li>Naknadu MOD-a (zakonom ograničena)</li>
                </ul>
                <div style={{ marginTop: 10, padding: "8px 10px", background: C.amberL,
                  borderRadius: 7, fontSize: 11, color: C.amber }}>
                  ⚠ MOD preuzima <strong>dugovječnosni rizik</strong> — obvezan je isplaćivati rentu
                  bez obzira koliko dugo živite. Zato je renta niža od "kapitala / e₆₅".
                </div>
              </div>
            </Card>

            <Card style={{ borderLeft: `4px solid ${C.purple}` }}>
              <h3 style={{ fontSize: 14, color: C.purple, marginBottom: 14 }}>📈 Kategorije OMF fondova (A, B, C)</h3>
              <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.7, marginBottom: 12 }}>
                Svako mirovinsko društvo upravlja s <strong>tri kategorije fonda</strong> s različitim
                rizičnim profilima. Ukupno 12 fondova na tržištu (4 društva × 3 kategorije).
              </p>
              {[
                { kat: "A", boja: C.green, profil: "Agresivni rast",
                  dionice: "do 100% dionica", za: "Mladi (>10 god. do mirovine)",
                  opis: "Najviši očekivani prinos dugoročno, ali i najveće kratkoročne oscilacije. Primjeren za osiguranike s dugim investicijskim horizontom." },
                { kat: "B", boja: C.blue, profil: "Umjereni rast",
                  dionice: "do 55% dionica", za: "Srednja dob (5–10 god. do mir.)",
                  opis: "Balans između rasta i sigurnosti. REGOS automatski prebacuje iz A u B 15 godina nakon prvog zaposlenja, ili 5 god. prije mirovine." },
                { kat: "C", boja: C.slate, profil: "Konzervativni",
                  dionice: "do 30% dionica", za: "Pred-mirovinska dob (<5 god.)",
                  opis: "Naglasak na zaštiti kapitala. REGOS automatski prebacuje iz A u B ako do mirovine ostane manje od 5 god. Niži prinos, niži rizik gubitka." },
              ].map(f => (
                <div key={f.kat} style={{ marginBottom: 10, padding: "10px 12px",
                  background: f.boja + "0a", border: `1px solid ${f.boja}22`, borderRadius: 9 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ fontSize: 16, fontWeight: 900, color: f.boja,
                        background: f.boja+"18", padding: "1px 8px", borderRadius: 6 }}>Kat. {f.kat}</span>
                      <span style={{ fontSize: 12, fontWeight: 700, color: C.navy }}>{f.profil}</span>
                    </div>
                    <span style={{ fontSize: 10, color: f.boja, fontWeight: 600 }}>{f.dionice}</span>
                  </div>
                  <div style={{ fontSize: 11, color: C.slateL, marginBottom: 3 }}>📍 {f.za}</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>{f.opis}</div>
                </div>
              ))}
              <div style={{ fontSize: 11, color: C.slateL, marginTop: 6 }}>
                Kategoriju možete promijeniti <strong>jednom godišnje</strong> u svom kalendarskom mjesecu
                rođenja (bez naknade). Fond (mirovinsko društvo) možete promijeniti u bilo koje vrijeme
                (bez naknade nakon 3 god. članstva).
              </div>
            </Card>
          </div>

          {/* Timeline isplate */}
          <Card style={{ marginBottom: 18 }}>
            <h3 style={{ fontSize: 14, color: C.navy, marginBottom: 18 }}>📋 Kad i kako se isplaćuju sredstva — korak po korak</h3>
            <div style={{ position: "relative", paddingLeft: 32 }}>
              {/* Vertical line */}
              <div style={{ position: "absolute", left: 10, top: 8, bottom: 8,
                width: 2, background: `linear-gradient(${C.navy}, ${C.green})`, borderRadius: 1 }} />
              {[
                { n: "1", boja: C.navy, naslov: "Ostvarivanje prava na mirovinu", emoji: "📝",
                  opis: "Podnosite zahtjev HZMO-u za starosnu ili prijevremenu mirovinu. HZMO izdaje Rješenje o mirovini i šalje podatke REGOS-u.",
                  detalj: null },
                { n: "2", boja: C.navy, naslov: "Odluka: kombinirana ili samo I. stup?", emoji: "⚖️",
                  opis: "Imate pravo izabrati isplatu SAMO iz I. stupa — tada se sva vaša OMF sredstva prenose u I. stup i mirovina se uvećava. Ili odabirete KOMBINIRANU mirovinu (I. + II. stup).",
                  detalj: "Istup iz II. stupa moguć je samo za osiguranike koji su u II. stup ušli dobrovoljno (oni koji su 2002. imali 40–50 god.). Ostali ne mogu izaći." },
                { n: "3", boja: C.blue, naslov: "Odabir MOD-a (za kombiniranu)", emoji: "🏦",
                  opis: "U roku od 15 dana od zaprimanja REGOS obavijesti, odabirete MOD na šalterima FINA-e, Mirovinskom informativnom centru ili putem e-Građani. Ako ne odaberete, REGOS vas raspoređuje u MOD po službenoj dužnosti.",
                  detalj: "Na raspolaganju su 4 MOD-a. Usporedite uvjete prije odabira — naknade i vrste mirovine mogu se razlikovati." },
                { n: "4", boja: C.blue, naslov: "Jednokratna isplata do 20% (opcija)", emoji: "💵",
                  opis: "Možete zatražiti jednokratnu isplatu do 20% akumuliranog kapitala PRIJE formiranja doživotne rente. MOD obračunava porez na dohodak na taj iznos. Uvjet: osnovna mirovina iz I. stupa mora biti veća od najniže mirovine.",
                  detalj: null },
                { n: "5", boja: C.purple, naslov: "Ugovor o mirovini s MOD-om", emoji: "📄",
                  opis: "MOD i vi potpisujete Ugovor o mirovini. U njemu se određuje: vrsta mirovine, iznos rente, uvjeti usklađivanja, eventualno zajamčeno razdoblje i ime korisnika.",
                  detalj: "Jednom sklopljeni ugovor nije moguće promijeniti. Zato je izbor vrste mirovine iznimno važna i TRAJNA odluka." },
                { n: "6", boja: C.green, naslov: "Isplata doživotne rente", emoji: "💰",
                  opis: "Renta se isplaćuje svaki mjesec dok ste živi — bez iznimke, bez obzira na duljinu života. MOD preuzima ovaj rizik. Usklađivanje rente (IPC) ovisi o odabranom programu.",
                  detalj: null },
              ].map((k, i) => (
                <div key={i} style={{ display: "flex", gap: 16, marginBottom: 18, position: "relative" }}>
                  <div style={{ position: "absolute", left: -28, top: 2, width: 20, height: 20,
                    borderRadius: "50%", background: k.boja, color: "#fff",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontSize: 10, fontWeight: 800, flexShrink: 0, zIndex: 1 }}>{k.n}</div>
                  <div style={{ background: k.boja + "08", border: `1px solid ${k.boja}22`,
                    borderRadius: 10, padding: "12px 14px", flex: 1 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                      <span style={{ fontSize: 16 }}>{k.emoji}</span>
                      <strong style={{ fontSize: 13, color: C.navy }}>{k.naslov}</strong>
                    </div>
                    <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.6, marginBottom: k.detalj ? 6 : 0 }}>
                      {k.opis}
                    </p>
                    {k.detalj && (
                      <div style={{ fontSize: 11, color: C.slateL, background: C.bg,
                        borderRadius: 7, padding: "6px 10px", lineHeight: 1.5 }}>
                        💡 {k.detalj}
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </Card>

          {/* Nasljeđivanje — centralni prikaz */}
          <Card style={{ marginBottom: 18, borderTop: `4px solid ${C.gold}` }}>
            <h3 style={{ fontSize: 14, color: C.navy, marginBottom: 6 }}>👨‍👩‍👧 Nasljeđivanje sredstava II. stupa — što zakon zaista kaže</h3>
            <p style={{ fontSize: 12, color: C.slateL, marginBottom: 16, lineHeight: 1.6 }}>
              Zakon o obveznim mirovinskim fondovima, čl. 103. st. 1. (NN 19/14, 93/15, 64/18, 115/18, 58/20)
            </p>

            {/* Flowchart nasljeđivanja */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr auto 1fr auto 1fr", gap: 8,
              alignItems: "center", marginBottom: 20 }}>
              {/* Slučaj 1: smrt prije umirovljenja */}
              <div style={{ background: C.navy + "0a", border: `1.5px solid ${C.navy}22`,
                borderRadius: 10, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 22, marginBottom: 6 }}>💀</div>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.navy, marginBottom: 6 }}>
                  Smrt PRIJE umirovljenja
                </div>
                <div style={{ fontSize: 10, color: C.slate, lineHeight: 1.5 }}>
                  Cijeli akumulirani kapital na osobnom računu u OMF-u
                </div>
              </div>

              <div style={{ fontSize: 18, color: C.slateL, textAlign: "center" }}>→</div>

              <div>
                <div style={{ background: C.greenL, border: `1.5px solid ${C.green}33`,
                  borderRadius: 10, padding: 12, marginBottom: 8, textAlign: "center" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.green, marginBottom: 4 }}>
                    ✅ Ako NEMA korisnika obiteljske mirovine (ZOMO)
                  </div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>
                    Sredstva su <strong>predmet nasljeđivanja</strong> prema Zakonu o nasljeđivanju.
                    Nasljednici (fizičke ili pravne osobe) podnose pravomoćno Rješenje o nasljeđivanju.
                  </div>
                </div>
                <div style={{ background: C.amberL, border: `1.5px solid ${C.gold}33`,
                  borderRadius: 10, padding: 12, textAlign: "center" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.amber, marginBottom: 4 }}>
                    ⚠ Ako IMA korisnika obiteljske mirovine (ZOMO)
                  </div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>
                    Obitelj prima obiteljsku mirovinu iz I. stupa.
                    Sredstva iz II. stupa ostaju u sustavu — <strong>ne nasljeđuju se</strong>.
                  </div>
                </div>
              </div>

              <div style={{ fontSize: 18, color: C.slateL, textAlign: "center" }}>→</div>

              <div style={{ background: C.blueL, border: `1.5px solid ${C.blue}22`,
                borderRadius: 10, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 22, marginBottom: 6 }}>📄</div>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.blue, marginBottom: 6 }}>Postupak isplate nasljednicima</div>
                <div style={{ fontSize: 10, color: C.slate, lineHeight: 1.5, textAlign: "left" }}>
                  1. Pravomoćno Rješenje o nasljeđivanju<br/>
                  2. Original ili ovjeren original<br/>
                  3. Zahtjev REGOS-u / FINA-i<br/>
                  4. Isplata na račun nasljednika<br/>
                  <span style={{ color: C.slateL, marginTop: 4, display: "block" }}>
                    Nasljednici mogu biti i pravne osobe, udruge, JLS (grad, općina) — ako se zakonski nasljednici odriču nasljedstva.
                  </span>
                </div>
              </div>
            </div>

            {/* Smrt nakon umirovljenja */}
            <div style={{ background: C.redL, border: `1.5px solid ${C.red}22`,
              borderRadius: 10, padding: 16, marginBottom: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.red, marginBottom: 8 }}>
                ⚠ Smrt NAKON umirovljenja (u tijeku isplate rente)
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.navy, marginBottom: 4 }}>Bez zajamčenog razdoblja</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>
                    Renta prestaje smrću. Preostali neisplaćeni kapital ostaje MOD-u.
                    Nasljednici ne dobivaju ništa.
                  </div>
                </div>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.green, marginBottom: 4 }}>Sa zajamčenim razdobljem (5–20 god.)</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>
                    Smrt unutar zajamčenog perioda → renta se nastavlja isplaćivati
                    <strong> imenovanom korisniku</strong> (min. 50% vaše rente) do kraja zajamčenog roka.
                  </div>
                </div>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.blue, marginBottom: 4 }}>Zajednička mirovina (bračni par)</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>
                    Smrću jednog bračnog druga, nadživjeli bračni drug nastavlja primati
                    <strong> min. 60%</strong> prvobitne rente do kraja života.
                  </div>
                </div>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.purple, marginBottom: 4 }}>Zajednička + zajamčeno razdoblje</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>
                    Kombinacija — najniži početni iznos rente, ali
                    <strong> maksimalna zaštita</strong> i bračnog druga i imenovanog korisnika.
                  </div>
                </div>
              </div>
            </div>

            <div style={{ padding: "10px 14px", background: C.purpleL, borderRadius: 9,
              fontSize: 11, color: C.purple, lineHeight: 1.6 }}>
              💡 <strong>Praktični savjet:</strong> Vrstu mirovine odabirete jednom pri sklapanju ugovora s MOD-om
              — ta odluka je <strong>nepromjenjiva</strong>. Razgovarajte s partnerom i savjetnikom
              prije odabira. Ako imate ovisne članove obitelji, razmatrajteo zajedničku mirovinu ili zajamčeno razdoblje,
              iako će početna renta biti niža.
            </div>
          </Card>

          {/* Prava i zaštite — tablični prikaz */}
          <Card>
            <h3 style={{ fontSize: 14, color: C.navy, marginBottom: 16 }}>✅ Sva vaša prava u II. stupu — pregled</h3>
            <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
              {[
                { pravo: "Pravo na uvid u stanje računa",
                  kako: "e-Građani portal, REGOS web, godišnji izviještaj poštom",
                  izvor: "ZOMF čl. 70." },
                { pravo: "Pravo izbora OMF fonda i kategorije",
                  kako: "FINA šalter, REGOS portal, e-Građani — jednom godišnje (kategorija), češće (fond)",
                  izvor: "ZOMF čl. 55." },
                { pravo: "Pravo prijenosa u drugi fond (bez naknade nakon 3 god.)",
                  kako: "Zahtjev REGOS-u, bez naknade za izlaz po isteku 3 god. članstva",
                  izvor: "ZOMF čl. 56." },
                { pravo: "Pravo na informacije o ulaganjima fonda",
                  kako: "Statut fonda, godišnje i polugodišnje izvješće — javno objavljeno",
                  izvor: "ZOMF čl. 75." },
                { pravo: "Pravo na jednokratnu isplatu do 20%",
                  kako: "Zahtjev MOD-u pri umirovljenju. Uvjet: I. stup mirovina > najniža mirovina",
                  izvor: "ZOMF čl. 102a." },
                { pravo: "Pravo na istup iz II. stupa (samo za dobrovoljne ulaznike)",
                  kako: "Zahtjev REGOS-u pri ostvarivanju prava na mirovinu. Samo za osiguranike koji su 2002. imali 40–50 god.",
                  izvor: "ZOMF čl. 104." },
                { pravo: "Pravo nasljednika na kapitalizirana sredstva",
                  kako: "Pravomoćno Rješenje o nasljeđivanju — REGOS isplaćuje zakonskim nasljednicima",
                  izvor: "ZOMF čl. 103." },
                { pravo: "Pravo na pritužbu HANFA-i",
                  kako: "Pisana pritužba HANFA-i ako mirovinsko društvo ili MOD krši zakone ili propise",
                  izvor: "Zakon o HANFA-i" },
              ].map((r, i) => (
                <div key={i} style={{ display: "grid", gridTemplateColumns: "2fr 3fr 1fr",
                  gap: 12, padding: "9px 12px",
                  background: i % 2 === 0 ? C.bg : C.white,
                  borderBottom: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.navy }}>{r.pravo}</div>
                  <div style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>{r.kako}</div>
                  <div style={{ fontSize: 10, color: C.slateL, fontStyle: "italic" }}>{r.izvor}</div>
                </div>
              ))}
            </div>
            <div style={{ marginTop: 12, fontSize: 11, color: C.slateL, padding: "8px 12px",
              background: C.bg, borderRadius: 7 }}>
              Izvor: Zakon o obveznim mirovinskim fondovima (ZOMF), NN 19/14, 93/15, 64/18, 115/18, 58/20 ·
              HANFA (hanfa.hr) · mss.gov.hr · HZMO (mirovinsko.hr)
            </div>
          </Card>
        </EduSection>
      </div>
    </div>
  );
}

// ─── III. STUP + INVALIDSKA + PLANER + USPOREDBA ────────────────────────────

// ── III. STUP ──────────────────────────────────────────────────────────────
function Stup3View({ user, onAuthRequired }) {
  const [uplata, setUplata] = useState(55.3);   // ~663.6/12
  const [god, setGod] = useState(25);
  const [prinos, setPrinos] = useState(0.055);
  const [posl, setPosl] = useState(0);
  const [naziv, setNaziv] = useState("III. stup scenarij");
  const [saved, setSaved] = useState(false);

  const DPS_MAX_UPLATA = 663.61;
  const DPS_STOPA = 0.15;
  const DPS_MAX = 99.54;

  const god_uplata = uplata * 12;
  const dps_god = Math.min(god_uplata, DPS_MAX_UPLATA) * DPS_STOPA;
  const dps_total = dps_god * god;
  const uk_uplata_god = god_uplata + dps_god + posl * 12;
  const pmt_eff = uk_uplata_god / 12;
  const r = prinos / 12; const n = god * 12;
  const kapital = r === 0 ? pmt_eff * n : pmt_eff * ((Math.pow(1+r,n)-1)/r);
  const kapital_bez_dps = r === 0 ? uplata*n*12 : uplata*12*((Math.pow(1+r,n)-1)/r);
  const roi_dps = (dps_total + (kapital - kapital_bez_dps)).toFixed(0);

  const projs = [5,10,15,20,25,30,35,40].filter(g => g <= god + 1).map(g => {
    const nn = g * 12;
    const k = r === 0 ? pmt_eff*nn : pmt_eff*((Math.pow(1+r,nn)-1)/r);
    return { god: `${g}g`, k: Math.round(k) };
  });

  async function handleSave() {
    if (!user) { onAuthRequired("Za spremanje scenarija potrebna je prijava."); return; }
    const sc = { id: Date.now(), naziv, created: Date.now(), tip: "stup3",
      stup3: { uplata, god, prinos, posl } };
    const ex = (await stGet(`scenarios:${user.email}`)) || [];
    await stSet(`scenarios:${user.email}`, [...ex, sc]);
    setSaved(true); setTimeout(() => setSaved(false), 2500);
  }

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>III. mirovinski stup</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>Dobrovoljna mirovinska štednja · Državni poticaj (DPS) 15% · Dostupno od 55 god.</p>
      </div>

      {/* DPS callout */}
      <div style={{ background: `linear-gradient(135deg, ${C.green}, #1a5c35)`, borderRadius: 14,
        padding: "20px 24px", marginBottom: 20, color: "#fff", display: "flex", gap: 24, flexWrap: "wrap" }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 11, color: "#9fefc0", fontWeight: 700, textTransform: "uppercase",
            letterSpacing: "0.08em", marginBottom: 6 }}>Državni poticaj (DPS) — jedinstven u HR</div>
          <p style={{ fontSize: 13, color: "rgba(255,255,255,0.85)", lineHeight: 1.7 }}>
            Hrvatska država daje <strong style={{ color: "#7bffb0" }}>15% od vaše godišnje uplate</strong> besplatno — do
            maksimalnih <strong style={{ color: "#7bffb0" }}>{DPS_MAX.toFixed(2)} EUR godišnje</strong>.
            Ostvaruje se uplatom <strong style={{ color: "#7bffb0" }}>{DPS_MAX_UPLATA} EUR godišnje</strong> ({(DPS_MAX_UPLATA/12).toFixed(2)} EUR/mj.).
            <br/>Na III. stup <strong style={{ color: "#7bffb0" }}>ne plaća se porez</strong> ni u fazi štednje ni pri isplati.
          </p>
        </div>
        <div style={{ display: "flex", gap: 14 }}>
          {[
            { v: `${DPS_MAX.toFixed(2)} EUR`, l: "Max DPS godišnje" },
            { v: "15%", l: "Stopa poticaja" },
            { v: "55 god.", l: "Dostupno od" },
          ].map((it,i) => (
            <div key={i} style={{ textAlign: "center", background: "rgba(255,255,255,0.12)", borderRadius: 10, padding: "10px 14px" }}>
              <div style={{ fontSize: 20, fontWeight: 800, color: "#7bffb0", fontFamily: "'Playfair Display',serif" }}>{it.v}</div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.65)", marginTop: 2 }}>{it.l}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        <div>
          <Card style={{ marginBottom: 14 }}>
            <h3 style={{ fontSize: 14, marginBottom: 16 }}>Parametri štednje</h3>
            <SmartParam icon="💳" title="Vaša mjesečna uplata" value={uplata} min={10} max={500} step={1}
              format={v => `${v.toFixed(0)} EUR/mj.`} onChange={setUplata}
              zones={[
                { upTo: 55.3, color: C.gold, label: "Optimum za max DPS" },
                { upTo: 150,  color: C.blue, label: "Iznad DPS praga" },
                { upTo: 500,  color: C.purple, label: "Visoka štednja" },
              ]}
              dynamicExplain={v => {
                const god_u = v * 12;
                const dps = Math.min(god_u, DPS_MAX_UPLATA) * DPS_STOPA;
                if (v < 55.3) return `Uplatom ${v.toFixed(0)} EUR/mj. (${god_u.toFixed(0)} EUR/god.) dobivate DPS od ${dps.toFixed(2)} EUR/god. Za max DPS trebate ${(DPS_MAX_UPLATA/12).toFixed(2)} EUR/mj.`;
                return `Ostvarujete maksimalni DPS od ${DPS_MAX.toFixed(2)} EUR/god. Uplate iznad ${(DPS_MAX_UPLATA/12).toFixed(2)} EUR/mj. ne donose više DPS, ali povećavaju kapital.`;
              }}
              explain={<div><p style={{ fontSize: 12, lineHeight: 1.7, color: C.slate }}>DPS se izračunava na godišnjem iznosu uplata, a isplaćuje se krajem godine za prethodnu kalendarsku godinu. Minimalna uplata za DPS nije propisana, ali ni poticaj ni prinos nema smisla bez redovitih uplata.</p></div>}
            />
            <SmartParam icon="⏱" title="Broj godina štednje" value={god} min={1} max={40} step={1}
              format={v => `${v} god.`} onChange={setGod}
              zones={[
                { upTo: 9,  color: C.amber, label: "Kratko" },
                { upTo: 19, color: C.gold,  label: "Srednje" },
                { upTo: 40, color: C.green, label: "Dugo — maksimalan učinak" },
              ]}
              dynamicExplain={v => `Za ${v} god. akumulirate ukupno ${fmt(kapital)} EUR. Sama složena kamata i DPS dodaju ${fmt(kapital - uplata*12*v)} EUR na vaše direktne uplate.`}
              explain={<div><p style={{ fontSize: 12, lineHeight: 1.7, color: C.slate }}>Sredstva su dostupna od 55. godine. Raniji start = eksponencijalno više novca zbog složene kamate. Čak i 5 godina više štednje može udvostručiti krajnji kapital.</p></div>}
            />
            <SmartParam icon="📈" title="Godišnji prinos fonda" value={prinos} min={0.01} max={0.10} step={0.005}
              format={v => `${(v*100).toFixed(1)}%`} onChange={setPrinos}
              zones={[
                { upTo: 0.034, color: C.amber, label: "Konzervativno" },
                { upTo: 0.065, color: C.green, label: "Prosjek DMF-ova" },
                { upTo: 0.10,  color: C.blue,  label: "Optimistično" },
              ]}
              dynamicExplain={v => `Uz ${(v*100).toFixed(1)}% prinos, na ${god} god. štednje prinos generira ${fmt(kapital - uplata*12*god - dps_total)} EUR iznad direktnih uplata i DPS-a.`}
              explain={<div><p style={{ fontSize: 12, lineHeight: 1.7, color: C.slate }}>Dobrovoljni mirovinski fondovi ulažu konzervativnije od II. stupa — manji udio dionica. Prosječni prinos HR DMF-ova 2010.–2024. bio je 4–6% godišnje.</p></div>}
            />
            <div style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 500, color: C.slate, marginBottom: 8 }}>
                Uplata poslodavca (ako uplaćuje u vaš III. stup)
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <RangeTrack value={posl} min={0} max={100} step={5} color={C.green}
                  onChange={setPosl} />
                <input className="num-inp" value={posl} onChange={e => setPosl(Math.min(100, Math.max(0, Number(e.target.value) || 0)))}
                  style={{ width: 60 }} />
                <span style={{ fontSize: 11, color: C.slateL, whiteSpace: "nowrap" }}>EUR/mj.</span>
              </div>
              <div style={{ fontSize: 11, color: C.green, marginTop: 4 }}>
                Uplata poslodavca do 67 EUR/mj. (804 EUR/god.) oslobođena je poreza na dohodak i doprinosa.
              </div>
            </div>
          </Card>
          <Card style={{ padding: "14px 18px" }}>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              <input value={naziv} onChange={e => setNaziv(e.target.value)}
                style={{ flex: 1, padding: "8px 12px", border: `1.5px solid ${C.border}`, borderRadius: 8, fontSize: 13 }} />
              <button onClick={handleSave} style={{
                padding: "8px 18px", background: saved ? C.green : "#166534", color: "#fff",
                border: "none", borderRadius: 8, fontWeight: 600, fontSize: 13, display: "flex", alignItems: "center", gap: 6 }}>
                {saved ? <CheckCircle size={15}/> : <Save size={15}/>}
                {saved ? "Spremljeno!" : "Spremi"}
              </button>
            </div>
          </Card>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Rezultati</h3>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
              <ResultBox label={`Akumulirani kapital (${god} god.)`} value={`${fmt(kapital)} EUR`} color={C.green} big />
              <ResultBox label="Ukupni DPS kroz karijeru" value={`${fmt(dps_total)} EUR`} color={C.gold} />
              <ResultBox label="Vaše direktne uplate" value={`${fmt(uplata*12*god)} EUR`} color={C.blue} />
              <ResultBox label="Prinos + DPS bonus" value={`${fmt(kapital - uplata*12*god)} EUR`} color={C.purple} />
            </div>
            <div style={{ padding: "10px 12px", background: C.greenL, borderRadius: 9, fontSize: 12 }}>
              <div style={{ fontWeight: 700, color: C.green }}>DPS je "besplatni novac" od države:</div>
              <div style={{ color: C.slate, marginTop: 4 }}>
                Na svakih 663,61 EUR/god. koje uplatite, dobivate automatski <strong>99,54 EUR više</strong> od države —
                to je zajamčeni prinos od 15% PRIJE investicijskog prinosa fonda.
              </div>
            </div>
          </Card>
          <Card>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 12 }}>Rast kapitala kroz godine</div>
            <ResponsiveContainer width="100%" height={170}>
              <AreaChart data={projs}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <XAxis dataKey="god" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 10 }} tickFormatter={v => `${Math.round(v/1000)}k`} />
                <Tooltip formatter={v => [fmtEur(v), "Kapital"]} />
                <Area type="monotone" dataKey="k" fill={C.greenL} stroke={C.green} strokeWidth={2} />
              </AreaChart>
            </ResponsiveContainer>
          </Card>
          <Card style={{ background: C.navy }}>
            <div style={{ color: C.goldL, fontSize: 11, fontWeight: 700, marginBottom: 10, textTransform: "uppercase" }}>Prednosti III. stupa vs. obična štednja</div>
            {[
              ["DPS — besplatni državni bonus", `+${fmt(dps_total)} EUR kroz karijeru`],
              ["Bez poreza na prihod od kapitala", "Obična štednja: 10% porez na kamate"],
              ["Prinos fonda > bankarska štednja", "Prosj. 4–6% vs. 0.5–2% bankarski depozit"],
              ["Porezna olakšica za poslodavca", "Do 804 EUR/god. neoporezivo"],
            ].map(([k,v],i) => (
              <div key={i} style={{ display:"flex", justifyContent:"space-between", padding:"6px 8px",
                background: i%2===0 ? "rgba(255,255,255,0.07)" : "transparent", borderRadius:6, marginBottom:3 }}>
                <span style={{ fontSize:11, color:"rgba(255,255,255,0.7)" }}>{k}</span>
                <span style={{ fontSize:11, fontWeight:700, color:"#7bffb0" }}>{v}</span>
              </div>
            ))}
          </Card>
        </div>
      </div>
    </div>
  );
}

// ── INVALIDSKA I OBITELJSKA MIROVINA ───────────────────────────────────────
function InvalidskaView() {
  const [pvb, setPvb] = useState(1.0);
  const [staz, setStaz] = useState(25);
  const [vrsta, setVrsta] = useState("potpuna");
  const [clan, setClan] = useState("bracni");

  // MF od 1.1.2026.: potpuna = 1.1, djelomična = 0.9
  const mf_map = { potpuna: 1.1, djelomicna: 0.9, profesionalna: 0.9 };
  const mf = mf_map[vrsta] || 1.0;
  const res = calcStup1({ pvb, staz, pf: 1.0, mf, avm: AVM });

  // Obiteljska mirovina: % od osiguranikove mirovine
  const obitelj_pct = { bracni: 0.70, djeca1: 0.70, djeca2: 0.80, djeca3: 0.90, roditelji: 0.50 };
  const obiteljna = res.efektivna * (obitelj_pct[clan] || 0.70);

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>Invalidska i obiteljska mirovina</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>ZOMO · Izmjene MF od 1.1.2026. (HZMO)</p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        {/* Invalidska */}
        <div>
          <div style={{ background: `linear-gradient(135deg, ${C.red}cc, #8b0000)`, borderRadius: 14,
            padding: "16px 20px", marginBottom: 14, color: "#fff" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#ffaaaa", textTransform: "uppercase", marginBottom: 6 }}>Invalidska mirovina</div>
            <p style={{ fontSize: 12, color: "rgba(255,255,255,0.85)", lineHeight: 1.7 }}>
              Pravo na invalidsku mirovinu stječe osiguranik koji izgubi radnu sposobnost. Formula je ista kao
              za starosnu mirovinu — razlikuje se samo <strong style={{ color: "#ffaaaa" }}>mirovinski faktor (MF)</strong>.
              Od 1.1.2026. MF-ovi su povećani novim ZOMO-om.
            </p>
          </div>
          <Card style={{ marginBottom: 14 }}>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Vrsta invalidske mirovine</h3>
            {[
              { v: "potpuna", emoji: "🚫", l: "Potpuni gubitak radne sposobnosti", mf: "MF = 1,1 (od 1.1.2026.)", boja: C.red,
                opis: "Osiguranik ne može raditi ni jedan posao. Mirovina se isplaćuje do navršenja dobi za starosnu mirovinu, kada se automatski pretvara u starosnu." },
              { v: "djelomicna", emoji: "⚕️", l: "Djelomični gubitak radne sposobnosti / profesionalna nesposobnost", mf: "MF = 0,9 (od 1.1.2026.)", boja: C.amber,
                opis: "Osiguranik može raditi u smanjenom opsegu (do pola radnog vremena). Od 1.1.2026. dozvoljen rad uz isplatu rente do 3,5h/dan." },
            ].map(o => (
              <button key={o.v} onClick={() => setVrsta(o.v)}
                style={{ width: "100%", marginBottom: 8, padding: "10px 12px",
                  border: `2px solid ${vrsta === o.v ? o.boja : C.border}`,
                  borderRadius: 10, background: vrsta === o.v ? o.boja+"11" : C.white,
                  cursor: "pointer", textAlign: "left" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: C.navy }}>{o.emoji} {o.l}</span>
                  <span style={{ fontSize: 10, fontWeight: 700, color: o.boja,
                    background: o.boja+"18", padding: "2px 7px", borderRadius: 8 }}>{o.mf}</span>
                </div>
                <div style={{ fontSize: 11, color: C.slateL, marginTop: 4, lineHeight: 1.4 }}>{o.opis}</div>
              </button>
            ))}
          </Card>
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Parametri osiguranika</h3>
            <SmartParam icon="💰" title="PVB (prosječni vrednosni bodovi)" value={pvb} min={0.3} max={3.0} step={0.05}
              format={v => v.toFixed(2)} onChange={setPvb} zones={[
                { upTo: 0.99, color: C.amber, label: "Ispod prosjeka" },
                { upTo: 1.09, color: C.green, label: "Prosjek" },
                { upTo: 3.0,  color: C.blue,  label: "Iznad prosjeka" },
              ]}
              dynamicExplain={v => `PVB ${v.toFixed(2)} uz ${staz} god. staža → invaldiska mirovina: ${fmt(calcStup1({pvb:v,staz,pf:1,mf,avm:AVM}).efektivna)} EUR/mj.`}
              explain={null} />
            <SmartParam icon="📅" title="Mirovinski staž u trenutku invalidnosti" value={staz} min={1} max={45} step={1}
              format={v => `${v} god.`} onChange={setStaz} zones={[
                { upTo: 4,  color: C.red,   label: "⚠ Ispod minimuma za neke slučajeve" },
                { upTo: 14, color: C.amber, label: "Kratki staž" },
                { upTo: 45, color: C.green, label: "Dobar staž" },
              ]}
              dynamicExplain={v => `Sa ${v} god. staža i PVB ${pvb.toFixed(2)}: invalidska mirovina = ${fmt(calcStup1({pvb,staz:v,pf:1,mf,avm:AVM}).efektivna)} EUR/mj.`}
              explain={<div><p style={{ fontSize:12, lineHeight:1.7, color:C.slate }}>Za invalidnost uzrokovanu bolešću potrebno je određeno „čekalno" razdoblje staža ovisno o dobi. Za ozljedu na radu i profesionalnu bolest pravo nastaje bez obzira na staž.</p></div>} />
            <div style={{ padding:"10px 12px", background: vrsta==="potpuna" ? C.redL : C.amberL, borderRadius:9, textAlign:"center" }}>
              <div style={{ fontSize:11, color:C.slate }}>Invalidska mirovina / mj.</div>
              <div style={{ fontSize:26, fontWeight:800, color: vrsta==="potpuna" ? C.red : C.amber, fontFamily:"'Playfair Display',serif" }}>
                {fmt(res.efektivna)} EUR
              </div>
              <div style={{ fontSize:11, color:C.slateL, marginTop:2 }}>
                MF = {mf.toFixed(1)} · Osobni bodovi: {res.ob.toFixed(3)}
              </div>
            </div>
          </Card>
        </div>

        {/* Obiteljska */}
        <div>
          <div style={{ background: `linear-gradient(135deg, #1a3566, #2a4f8c)`, borderRadius: 14,
            padding: "16px 20px", marginBottom: 14, color: "#fff" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: C.goldL, textTransform: "uppercase", marginBottom: 6 }}>Obiteljska mirovina</div>
            <p style={{ fontSize: 12, color: "rgba(255,255,255,0.85)", lineHeight: 1.7 }}>
              Pravo na obiteljsku mirovinu stječu članovi obitelji umrlog osiguranika ili korisnika mirovine.
              Iznos je postotak od mirovine koju bi osiguranik primao (ili je primao).
            </p>
          </div>
          <Card style={{ marginBottom: 14 }}>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Tko ima pravo i koliki postotak?</h3>
            {[
              { v: "bracni", emoji: "💑", l: "Bračni / izvanbračni drug", pct: "70%",
                uvjeti: "Navršenih 50 god. ili invalidnost, ili briga za djecu s pravom na obiteljsku mirovinu" },
              { v: "djeca1", emoji: "👶", l: "1 dijete (bez oba roditelja)", pct: "70%",
                uvjeti: "Do 15 god. ili do kraja školovanja (max. 26 god.), ili trajno bez radne sposobnosti" },
              { v: "djeca2", emoji: "👦👧", l: "2 djece (bez oba roditelja)", pct: "80%",
                uvjeti: "Isti uvjeti kao za 1 dijete, postotak se dijeli ravnomjerno" },
              { v: "djeca3", emoji: "👨‍👧‍👦", l: "3+ djece (bez oba roditelja)", pct: "90%",
                uvjeti: "Maksimalni postotak za djecu bez oba roditelja" },
              { v: "roditelji", emoji: "👴👵", l: "Roditelji / braća / sestre", pct: "50%",
                uvjeti: "Ako ih je osiguranik uzdržavao i nemaju dovoljno vlastitih prihoda" },
            ].map(o => (
              <button key={o.v} onClick={() => setClan(o.v)}
                style={{ width: "100%", marginBottom: 6, padding: "8px 12px",
                  border: `2px solid ${clan===o.v ? C.blue : C.border}`,
                  borderRadius: 9, background: clan===o.v ? C.blueL : C.white, cursor: "pointer", textAlign: "left" }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                  <span style={{ fontSize:12, fontWeight:700, color:C.navy }}>{o.emoji} {o.l}</span>
                  <span style={{ fontSize:13, fontWeight:800, color:C.blue }}>{o.pct}</span>
                </div>
                <div style={{ fontSize:10, color:C.slateL, marginTop:3, lineHeight:1.4 }}>{o.uvjeti}</div>
              </button>
            ))}
          </Card>
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Iznos obiteljske mirovine</h3>
            <SmartParam icon="💰" title="PVB umrlog osiguranika" value={pvb} min={0.3} max={3.0} step={0.05}
              format={v => v.toFixed(2)} onChange={setPvb} zones={[
                { upTo: 0.99, color: C.amber }, { upTo: 1.09, color: C.green }, { upTo: 3.0, color: C.blue }
              ]}
              dynamicExplain={v => `Osnovna mirovina: ${fmt(calcStup1({pvb:v,staz,pf:1,mf:1,avm:AVM}).efektivna)} EUR → Obiteljska (${Math.round((obitelj_pct[clan]||0.7)*100)}%): ${fmt(calcStup1({pvb:v,staz,pf:1,mf:1,avm:AVM}).efektivna*(obitelj_pct[clan]||0.7))} EUR`}
              explain={null} />
            <SmartParam icon="📅" title="Staž osiguranika" value={staz} min={1} max={45} step={1}
              format={v => `${v} god.`} onChange={setStaz} zones={[
                { upTo: 14, color: C.amber }, { upTo: 45, color: C.green }
              ]}
              dynamicExplain={v => `Osnovna mirovina (PVB ${pvb.toFixed(2)}, ${v} god.): ${fmt(calcStup1({pvb,staz:v,pf:1,mf:1,avm:AVM}).efektivna)} EUR`}
              explain={null} />
            <div style={{ padding:"10px 12px", background: C.blueL, borderRadius:9, textAlign:"center", marginBottom:10 }}>
              <div style={{ fontSize:11, color:C.slate }}>Obiteljska mirovina / mj.</div>
              <div style={{ fontSize:26, fontWeight:800, color:C.blue, fontFamily:"'Playfair Display',serif" }}>
                {fmt(obiteljna)} EUR
              </div>
              <div style={{ fontSize:11, color:C.slateL, marginTop:2 }}>
                {Math.round((obitelj_pct[clan]||0.7)*100)}% od {fmt(res.efektivna)} EUR (osnovna mirovina)
              </div>
            </div>
            <div style={{ fontSize:11, color:C.slateL, lineHeight:1.6, padding:"8px 10px", background:C.bg, borderRadius:8 }}>
              ⚠ Ovi iznosi su aproksimativni. Stvarni uvjeti i iznosi ovisni su o konkretnim okolnostima.
              Za točne podatke obratite se HZMO-u ili koristite mojamirovina.hr (e-Građani).
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}

// ── "KOLIKO MI NEDOSTAJE?" PLANER ──────────────────────────────────────────
function PlanerView() {
  const [zeljenaMirovina, setZeljenaMirovina] = useState(900);
  const [trenutnaPlacha, setTrenutnaPlaca] = useState(1300);
  const [trenutniStaz, setTrenutniStaz] = useState(20);
  const [dob, setDob] = useState(42);
  const [spol, setSpol] = useState("M");
  const [prinos2, setPrinos2] = useState(0.055);

  const mirDob = spol === "M" ? 65 : 64;
  const godDoMir = Math.max(0, mirDob - dob);
  const buduciStaz = trenutniStaz + godDoMir;
  const pvb_est = trenutnaPlacha / 1300;
  const r1 = calcStup1({ pvb: pvb_est, staz: buduciStaz, pf: 1, mf: 1, avm: AVM });
  const k2 = calcStup2FV({ bruto_placa: trenutnaPlacha, god_stednje: godDoMir, neto_prinos: prinos2 });
  const renta2 = rentaDozivotna(k2 * 0.8);
  const dod = r1.efektivna * 0.27;
  const ukupno = r1.efektivna + dod + renta2;
  const gap = zeljenaMirovina - ukupno;
  const gap_poz = gap > 0;

  // Što treba promijeniti za doći do željene mirovine
  const potrebnaPvb = zeljenaMirovina / (buduciStaz * 1.27 * AVM);
  const potrebnaPlaca = potrebnaPvb * 1300;
  const potrebniStaz = zeljenaMirovina / (pvb_est * 1.27 * AVM);
  const potrebniStup3_god = dps => {
    if (!gap_poz) return 0;
    const r = 0.055 / 12; const n = godDoMir * 12;
    const k_pot = gap * 18 * 12;
    return n === 0 ? k_pot : k_pot * r / (Math.pow(1+r,n)-1);
  };
  const stup3_mj = potrebniStup3_god();

  const progressData = [
    { name: "I. stup", value: Math.round(r1.efektivna), fill: C.navy },
    { name: "Dodatak 27%", value: Math.round(dod), fill: C.gold },
    { name: "II. stup renta", value: Math.round(renta2), fill: C.purple },
    { name: gap_poz ? "Nedostaje ↑" : "Višak ↑", value: Math.abs(Math.round(gap)), fill: gap_poz ? C.red+"66" : C.green+"66" },
  ];

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>Koliko mi nedostaje?</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>Unesite željenu mirovinu — kalkulator pokazuje razliku i što učiniti da je dosegnete</p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        <Card>
          <h3 style={{ fontSize: 14, marginBottom: 16 }}>Vaša situacija</h3>
          <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
            {[["M","Muškarac 🧔","Puna dob: 65 god."],["Ž","Žena 👩","Puna dob: ~64 god."]].map(([v,l,s]) => (
              <button key={v} onClick={() => setSpol(v)} style={{
                flex:1, padding:"8px 10px", border:`2px solid ${spol===v ? C.navy : C.border}`,
                borderRadius:9, background:spol===v ? C.navy : C.white,
                color:spol===v?"#fff":C.navy, cursor:"pointer" }}>
                <div style={{ fontSize:13, fontWeight:700 }}>{l}</div>
                <div style={{ fontSize:10, opacity:0.7, marginTop:2 }}>{s}</div>
              </button>
            ))}
          </div>
          <SmartParam icon="🎯" title="Željena mirovina" value={zeljenaMirovina} min={300} max={3000} step={50}
            format={v => `${v.toLocaleString("hr-HR")} EUR/mj.`} onChange={setZeljenaMirovina}
            zones={[
              { upTo: 499,  color: C.red,    label: "Ispod najniže mirovine" },
              { upTo: 799,  color: C.amber,  label: "Skromna mirovina" },
              { upTo: 1199, color: C.gold,   label: "Prosječna" },
              { upTo: 1799, color: C.green,  label: "Udobna mirovina" },
              { upTo: 3000, color: C.blue,   label: "Visoka mirovina" },
            ]}
            dynamicExplain={v => `Stopa zamjene: ${pct(v/trenutnaPlacha)} vaše sadašnje bruto plaće.`}
            explain={null} />
          <SmartParam icon="💼" title="Trenutna bruto plaća" value={trenutnaPlacha} min={700} max={5000} step={50}
            format={v => `${v.toLocaleString("hr-HR")} EUR`} onChange={setTrenutnaPlaca}
            zones={[
              { upTo: 1099, color: C.amber }, { upTo: 1499, color: C.green }, { upTo: 5000, color: C.blue }
            ]}
            dynamicExplain={v => `Procijenjeni PVB: ${(v/1300).toFixed(2)} (prosjek RH = 1.300 EUR)`}
            explain={null} />
          <SmartParam icon="📅" title="Trenutni mirovinski staž" value={trenutniStaz} min={0} max={45} step={1}
            format={v => `${v} god.`} onChange={setTrenutniStaz}
            zones={[
              { upTo: 14, color: C.amber }, { upTo: 29, color: C.gold }, { upTo: 45, color: C.green }
            ]}
            dynamicExplain={v => `Do mirovine ostaje još ${godDoMir} god. Na kraju karijere imat ćete ukupno ${v + godDoMir} god. staža.`}
            explain={null} />
          <SmartParam icon="🎂" title="Trenutna dob" value={dob} min={18} max={64} step={1}
            format={v => `${v} god.`} onChange={setDob}
            zones={[
              { upTo: 34, color: C.green, label: "Puno vremena za planiranje" },
              { upTo: 49, color: C.gold,  label: "Sredina karijere" },
              { upTo: 59, color: C.amber, label: "Bliže kraju karijere" },
              { upTo: 64, color: C.red,   label: "Kratko do mirovine" },
            ]}
            dynamicExplain={v => `Do mirovine (${mirDob} god.) preostaje ${Math.max(0,mirDob-v)} god.`}
            explain={null} />
        </Card>

        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {/* Gap vizualizacija */}
          <div style={{ background: gap_poz
              ? `linear-gradient(135deg, #3a0a0a, #6b1515)`
              : `linear-gradient(135deg, #0a2a1a, #1a5c35)`,
            borderRadius: 14, padding: "20px 24px", color: "#fff" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: gap_poz ? "#ffaaaa" : "#9fefc0",
              textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 8 }}>
              {gap_poz ? "⚠ Razlika do željene mirovine" : "✅ Cilj je dosegnut!"}
            </div>
            <div style={{ display: "flex", gap: 20, alignItems: "flex-end", flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)" }}>Projicirana ukupna mirovina</div>
                <div style={{ fontSize: 28, fontWeight: 900, color: gap_poz ? "#ff9090" : "#7bffb0",
                  fontFamily: "'Playfair Display',serif" }}>{fmt(ukupno)} EUR</div>
              </div>
              <div style={{ fontSize: 24, color: "rgba(255,255,255,0.4)" }}>{gap_poz ? "vs" : "vs"}</div>
              <div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)" }}>Željena mirovina</div>
                <div style={{ fontSize: 28, fontWeight: 900, color: "#fff",
                  fontFamily: "'Playfair Display',serif" }}>{fmt(zeljenaMirovina)} EUR</div>
              </div>
              <div style={{ background: (gap_poz ? "#ff9090" : "#7bffb0") + "22", borderRadius: 10,
                padding: "10px 16px", marginLeft: "auto" }}>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.7)" }}>
                  {gap_poz ? "Nedostaje" : "Višak"}
                </div>
                <div style={{ fontSize: 22, fontWeight: 900,
                  color: gap_poz ? "#ffaaaa" : "#9fefc0", fontFamily: "'Playfair Display',serif" }}>
                  {fmt(Math.abs(gap))} EUR
                </div>
              </div>
            </div>
          </div>

          {/* Decomposition */}
          <Card>
            <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 12 }}>Struktura projektirane mirovine</div>
            <ResponsiveContainer width="100%" height={160}>
              <BarChart data={progressData} layout="vertical" margin={{ left: 80, right: 20 }}>
                <XAxis type="number" tick={{ fontSize: 10 }} />
                <YAxis dataKey="name" type="category" tick={{ fontSize: 11 }} width={80} />
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" horizontal={false} />
                <Tooltip formatter={v => [fmtEur(v), ""]} />
                <Bar dataKey="value" radius={[0,4,4,0]}>
                  {progressData.map((d,i) => <Cell key={i} fill={d.fill} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </Card>

          {/* Akcijski plan */}
          {gap_poz && (
            <Card style={{ borderTop: `3px solid ${C.gold}` }}>
              <h3 style={{ fontSize: 14, marginBottom: 14, color: C.navy }}>💡 Akcijski plan — što možete učiniti</h3>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {[
                  { emoji: "🌱", l: "III. stup (minimalna uplata za cilj)",
                    v: `${stup3_mj > 0 ? fmt(stup3_mj) : "Nije potrebno"} EUR/mj.`,
                    opis: "Dopunska štednja uz DPS poticaj — najefikasniji put za popunjavanje razlike", color: C.green },
                  { emoji: "💼", l: "Potrebna plaća za željenu mirovinu",
                    v: `${fmt(potrebnaPlaca)} EUR bruto`,
                    opis: `PVB ${potrebnaPvb.toFixed(2)} — ${((potrebnaPvb-1)*100).toFixed(0)}% ${potrebnaPvb > 1 ? "iznad" : "ispod"} prosjeka RH`, color: C.blue },
                  { emoji: "📅", l: "Potrebni staž (uz sadašnju plaću)",
                    v: `${Math.max(0,Math.ceil(potrebniStaz))} god.`,
                    opis: potrebniStaz > buduciStaz ? `Trebate ${Math.ceil(potrebniStaz - buduciStaz)} god. više od planirane mirovine` : "Ostvarivo u planiranom roku", color: C.purple },
                  { emoji: "⏳", l: "Razmislite o odgodi mirovine",
                    v: godDoMir > 2 ? `+${(Math.min(24,godDoMir)*0.45).toFixed(1)}% za 2 god. odgode` : "N/A",
                    opis: "Svaki mjesec odgode donosi +0,45% trajnog povećanja PF-a", color: C.gold },
                ].map((it, i) => (
                  <div key={i} style={{ display:"flex", gap:12, padding:"10px 12px",
                    background: it.color+"0a", border:`1px solid ${it.color}22`, borderRadius:9 }}>
                    <span style={{ fontSize:20, flexShrink:0 }}>{it.emoji}</span>
                    <div style={{ flex:1 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                        <span style={{ fontSize:12, fontWeight:600, color:C.navy }}>{it.l}</span>
                        <span style={{ fontSize:13, fontWeight:800, color:it.color }}>{it.v}</span>
                      </div>
                      <div style={{ fontSize:11, color:C.slateL, marginTop:2 }}>{it.opis}</div>
                    </div>
                  </div>
                ))}
              </div>
            </Card>
          )}

          {!gap_poz && (
            <Card style={{ borderTop: `3px solid ${C.green}`, textAlign: "center", padding: "24px 20px" }}>
              <div style={{ fontSize: 36, marginBottom: 8 }}>🎉</div>
              <h3 style={{ fontSize: 15, marginBottom: 6, color: C.green }}>Cilj je dosegnut!</h3>
              <p style={{ fontSize: 12, color: C.slateL, lineHeight: 1.7 }}>
                Projicirana mirovina od <strong>{fmt(ukupno)} EUR/mj.</strong> premašuje
                željenih <strong>{fmt(zeljenaMirovina)} EUR/mj.</strong> za <strong>{fmt(-gap)} EUR</strong>.<br/>
                Razmotrite III. stup za daljnji rast ili raniji odlazak u mirovinu.
              </p>
            </Card>
          )}
        </div>
      </div>
    </div>
  );
}

// ── USPOREDBA S NAJBOLJIM SUSTAVIMA ────────────────────────────────────────
function UsporedbaSustaviView() {
  const [placha, setPlacha] = useState(1300);
  const [stazGod, setStazGod] = useState(40);

  const hr_m = calcStup1({ pvb: placha/1300, staz: stazGod, pf:1, mf:1, avm:AVM }).efektivna;
  const hr_dod = hr_m * 0.27;
  const hr_2 = rentaDozivotna(calcStup2FV({bruto_placa:placha,god_stednje:stazGod,neto_prinos:0.055})*0.8);
  const hr_uk = hr_m + hr_dod + hr_2;
  const hr_stopa = hr_uk / placha;

  const sustavi = [
    {
      zemlja: "🇭🇷 Hrvatska", kod: "HR", mercer: "C+", stopa: hr_stopa, opis_stopa: `${pct(hr_stopa)} (I.+II.+Dod.)`,
      boja: C.red, score: 55,
      opis: "Miješani sustav: I. stup (generacijska solidarnost) + II. stup (individualna kapitalizacija). Prosječna stopa zamjene ~38% od I. stupa samo; s II. stupu i dodatkom doseže ~50%.",
      doprinos: "20% (15% I. + 5% II.)", dob: "65 M / ~64 Ž",
      prednosti: ["II. stup osobna imovina", "Nasljedivo (uz uvjete)", "DPS za III. stup"],
      slabosti: ["Niska stopa zamjene I. stupa", "Demografski pritisak", "Nema garancije kapitala u doživotnoj renti"],
    },
    {
      zemlja: "🇳🇱 Nizozemska", kod: "NL", mercer: "A (85.4)", stopa: 0.96, opis_stopa: "~96% (neto)",
      boja: C.blue, score: 85.4,
      opis: "Tri stupa: AOW (osnovna državna, flat-rate ~12.000 EUR/god./par), obavezne strukovne mirovine (DB→DC do 2028.) i osobna štednja. Cjelokupni sustav vodi k zamjeni 70–100% dohotka.",
      doprinos: "~28% (poslodavac dominira)", dob: "67 (→ 67g.3mj. 2028.)",
      prednosti: ["Najviša stopa zamjene na svijetu", "Obavezne strukovne mirovine za gotovo sve", "Indeksacija uz inflaciju"],
      slabosti: ["Visoki doprinosi", "Prelazak DB→DC uvodi investicijski rizik", "Kompleksni sustav"],
    },
    {
      zemlja: "🇩🇰 Danska", kod: "DK", mercer: "A (82.3)", stopa: 0.80, opis_stopa: "~80% (neto)",
      boja: "#c41e3a", score: 82.3,
      opis: "Tri stupa: Folkepension (osnovna), ATP (obavezna DC za sve zaposlene), LifF / LD (privatne obavezne DC mirovine). Kombinacija socijalnog osiguranja i kapitalzirane štednje.",
      doprinos: "~12-15% + ATP", dob: "67 (vezana uz očekivani životni vijek)",
      prednosti: ["Visoka pokrivenost svih radnika", "Transparentan DC sustav", "Sigurna osnovna mirovina za sve"],
      slabosti: ["Visoki porezi za financiranje", "Dob umirovljenja raste s očekivanim vijekom", "Složeni sustav ATP/LifF"],
    },
    {
      zemlja: "🇮🇸 Island", kod: "IS", mercer: "A (84.0)", stopa: 0.95, opis_stopa: "~95%+ (neto)",
      boja: "#003897", score: 84.0,
      opis: "Obavezna trofazna mirovinska štednja: državna (Tryggingastofnun), obavezna strukovne (min. 15,5% plaće), privatna dobrovoljno. Visoke obavezne stope doprinosa jamče visoke mirovine.",
      doprinos: "12% radnik + 11,5% poslodavac (obavezno)", dob: "67",
      prednosti: ["Najviša stopa zamjene OECD", "Obavezna visoka kapitalizacija", "Nasljedive mirovine"],
      slabosti: ["Visoki doprinosi (23.5%!)", "Mali tržišni fond podložan vanjskim šokovima", "Geografska izoliranost"],
    },
    {
      zemlja: "🇦🇺 Australija", kod: "AU", mercer: "B+ (76.7)", stopa: 0.52, opis_stopa: "~52% (bruto, s Superannuation)",
      boja: "#00843D", score: 76.7,
      opis: "Superannuation — obavezna osobna mirovinska štednja (DC). Trenutno 11,5% plaće od poslodavca → raste na 12% do 2025. Vlada pruža minimalnu Age Pension za one s niskim kapitalom.",
      doprinos: "11.5% poslodavac (→12%)", dob: "67",
      prednosti: ["Vlastiti kapital, potpuno nasljediv", "Porezne olakšice na prinos", "Iznosi rastu s plaćom"],
      slabosti: ["Niska osnovna mirovina (samo za siromašne)", "Tržišni rizik", "Nedovoljna pokrivenost za niske plaće"],
    },
    {
      zemlja: "🇸🇬 Singapur", kod: "SG", mercer: "A (78.7)", stopa: 0.65, opis_stopa: "~65% (CPF)",
      boja: "#EF3340", score: 78.7,
      opis: "CPF (Central Provident Fund) — obavezni vlastiti račun. Dio CPF-a ide za mirovinu (CPF-RA), dio za zdravstvo (Medisave) i stanovanje. Zajamčeni minimalni prinos od 4% godišnje!",
      doprinos: "37% ukupno (20% radnik + 17% poslodavac)", dob: "65",
      prednosti: ["Zajamčeni minimalni prinos 4%", "Nasljedivo", "Jedan račun za mirovinu, zdravstvo i stan"],
      slabosti: ["Visoki doprinosi (37%!)", "Ograničena fleksibilnost isplate", "Ovisi o fiskalnoj politici države"],
    },
  ];

  const chartData = sustavi.map(s => ({
    name: s.zemlja.split(" ")[1],
    score: s.score,
    stopa: Math.round(s.stopa * 100),
    fill: s.boja,
  }));

  const [selected, setSelected] = useState("HR");

  const sel = sustavi.find(s => s.kod === selected) || sustavi[0];

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>Svjetski najuspješniji mirovinski sustavi</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>Mercer CFA Global Pension Index 2025 · OECD Pensions at a Glance 2025</p>
      </div>

      {/* Kontekst */}
      <div style={{ background: `linear-gradient(135deg, ${C.navy}, #1a3566)`, borderRadius: 14,
        padding: "20px 24px", marginBottom: 20, color: "#fff" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 16 }}>
          {[
            { v: "85.4", l: "Nizozemska — svjetski #1 (Mercer 2025)", c: C.goldL },
            { v: "63%", l: "OECD prosjek neto stope zamjene", c: "#7fbcff" },
            { v: "38%", l: "Hrvatska stopa zamjene (samo I. stup)", c: "#ff9090" },
            { v: "~50%", l: "HR s II. stupom i dodatkom 27%", c: "#9fefc0" },
          ].map((it,i) => (
            <div key={i} style={{ textAlign: "center" }}>
              <div style={{ fontSize: 24, fontWeight: 900, color: it.c, fontFamily: "'Playfair Display',serif", lineHeight: 1 }}>{it.v}</div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.6)", marginTop: 6, lineHeight: 1.4 }}>{it.l}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
        <Card>
          <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 12 }}>Mercer skor i stopa zamjene po sustavima</div>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={chartData} margin={{ top: 5, right: 10, bottom: 5, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
              <XAxis dataKey="name" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} domain={[0,100]} />
              <Tooltip formatter={(v,n) => [n==="score" ? `${v} bod.` : `${v}%`, n==="score" ? "Mercer skor" : "Stopa zamjene"]} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Bar dataKey="score" name="Mercer skor" radius={[4,4,0,0]}>
                {chartData.map((d,i) => <Cell key={i} fill={d.fill} opacity={0.9} />)}
              </Bar>
              <Bar dataKey="stopa" name="Stopa zamjene %" fill={C.navy+"44"} radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 12 }}>Odaberite sustav za detalje</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 14 }}>
            {sustavi.map(s => (
              <button key={s.kod} onClick={() => setSelected(s.kod)} style={{
                padding: "5px 12px", border: `2px solid ${selected===s.kod ? s.boja : C.border}`,
                borderRadius: 20, background: selected===s.kod ? s.boja : C.white,
                color: selected===s.kod ? "#fff" : C.navy, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
                {s.zemlja}
              </button>
            ))}
          </div>
          {sel && (
            <div style={{ background: sel.boja+"0a", border: `1.5px solid ${sel.boja}22`, borderRadius: 10, padding: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                <span style={{ fontSize: 14, fontWeight: 700 }}>{sel.zemlja}</span>
                <span style={{ fontSize: 11, fontWeight: 700, color: sel.boja,
                  background: sel.boja+"18", padding: "2px 9px", borderRadius: 10 }}>Mercer {sel.mercer}</span>
              </div>
              <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.7, marginBottom: 10 }}>{sel.opis}</p>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 10 }}>
                <div style={{ background: C.bg, borderRadius: 7, padding: "7px 10px" }}>
                  <div style={{ fontSize: 10, color: C.slateL }}>Stopa zamjene</div>
                  <div style={{ fontSize: 14, fontWeight: 700, color: sel.boja }}>{sel.opis_stopa}</div>
                </div>
                <div style={{ background: C.bg, borderRadius: 7, padding: "7px 10px" }}>
                  <div style={{ fontSize: 10, color: C.slateL }}>Doprinos (ukupno)</div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.navy }}>{sel.doprinos}</div>
                </div>
                <div style={{ background: C.bg, borderRadius: 7, padding: "7px 10px" }}>
                  <div style={{ fontSize: 10, color: C.slateL }}>Dob umirovljenja</div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.navy }}>{sel.dob}</div>
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.green, marginBottom: 4 }}>✅ Prednosti</div>
                  {sel.prednosti.map((p,i) => <div key={i} style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>• {p}</div>)}
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.red, marginBottom: 4 }}>⚠ Slabosti</div>
                  {sel.slabosti.map((s,i) => <div key={i} style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>• {s}</div>)}
                </div>
              </div>
            </div>
          )}
        </Card>
      </div>

      {/* Usporedba za vaše podatke */}
      <Card style={{ marginBottom: 20 }}>
        <h3 style={{ fontSize: 14, marginBottom: 14 }}>Kako bi izgledala vaša mirovina u različitim sustavima</h3>
        <div style={{ display: "flex", gap: 16, marginBottom: 16, flexWrap: "wrap" }}>
          <div style={{ flex: 1 }}>
            <MiniSlider label="Bruto plaća" value={placha} min={800} max={4000} step={50}
              format={v => `${v.toLocaleString("hr-HR")} EUR/mj.`} onChange={setPlacha} color={C.navy} />
          </div>
          <div style={{ flex: 1 }}>
            <MiniSlider label="Godina rada (staž)" value={stazGod} min={15} max={45} step={1}
              format={v => `${v} god.`} onChange={setStazGod} color={C.navy} />
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }}>
          {[
            { l: "🇭🇷 HR (sadašnji sustav)", v: hr_uk, s: hr_stopa, c: C.red, note: "I. + Dod. + II." },
            { l: "🇳🇱 NL model (96% neto)", v: placha * 0.96, s: 0.96, c: C.blue, note: "AOW + strukovne" },
            { l: "🇩🇰 DK model (~80% neto)", v: placha * 0.80, s: 0.80, c: "#c41e3a", note: "Folkepension + ATP + LifF" },
            { l: "🇮🇸 IS model (~95%)", v: placha * 0.95, s: 0.95, c: "#003897", note: "Državna + obavezna kapital." },
            { l: "🇦🇺 AU model (~52%)", v: placha * 0.52, s: 0.52, c: "#00843D", note: "Age Pension + Super" },
            { l: "🇸🇬 SG model (~65%)", v: placha * 0.65, s: 0.65, c: "#EF3340", note: "CPF (zajamčeni prinos 4%)" },
          ].map((it,i) => (
            <div key={i} style={{ background: it.c+"0a", border: `1.5px solid ${it.c}22`,
              borderRadius: 10, padding: "12px 14px", textAlign: "center" }}>
              <div style={{ fontSize: 11, color: C.slate, marginBottom: 4, lineHeight: 1.4 }}>{it.l}</div>
              <div style={{ fontSize: 20, fontWeight: 800, color: it.c, fontFamily: "'Playfair Display',serif" }}>{fmt(it.v)} EUR</div>
              <div style={{ fontSize: 10, color: C.slateL, marginTop: 2 }}>{pct(it.s)} bruto plaće</div>
              <div style={{ fontSize: 9, color: C.slateL, opacity: 0.7, marginTop: 2 }}>{it.note}</div>
            </div>
          ))}
        </div>
        <div style={{ marginTop: 14, padding: "10px 14px", background: C.amberL, borderRadius: 9,
          fontSize: 11, color: C.amber, lineHeight: 1.6 }}>
          ⚠ Ovo su aproksimacije na temelju OECD podataka i objavljenih stopa zamjene. Stvarni iznosi
          u svakom sustavu ovise o konkretnim pravilima, godinama rada, visini doprinosa i tržišnim prinosima.
          Izvor: Mercer CFA Global Pension Index 2025 · OECD Pensions at a Glance 2025.
        </div>
      </Card>

      {/* Ključne lekcije */}
      <Card>
        <h3 style={{ fontSize: 14, marginBottom: 14 }}>🎓 Što možemo naučiti od boljih sustava?</h3>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          {[
            { emoji: "💰", t: "Visoke obavezne stope doprinosa", opis: "NL: 28%, IS: 23.5%, SG: 37%. HR ukupno 20%. Viši doprinosi = više mirovine. Pitanje je tko ih snosi i koliko to usporava ekonomiju.", boja: C.blue },
            { emoji: "🏗", t: "Obavezne strukovne mirovine", opis: "Nizozemska i Danska osiguravaju gotovo 100% radnika kroz obavezne kolektivne ugovore. HR ima samo I. i II. stup — bez kolektivnih strukovnih planova.", boja: C.purple },
            { emoji: "📈", t: "Zajamčeni minimalni prinos", opis: "Singapurski CPF jamči 4% prinos godišnje čak i u lošim godinama. HR fondovi nemaju takvo jamstvo — prinos može biti negativan.", boja: C.green },
            { emoji: "⚖️", t: "Automatska usklađenost s demografijom", opis: "Danska i Islanda pomiču dob umirovljenja automatski s rastom očekivanog vijeka. HR ima fiksnu dob — demografska bomba tihuje.", boja: C.gold },
            { emoji: "👨‍👩‍👧", t: "Nasljedivost kapitala", opis: "Australija (Super) i Singapur (CPF) — kapital je u potpunosti nasljediv. HR II. stup — nasljediv samo uz uvjete. I. stup — nenasljediv.", boja: C.red },
            { emoji: "🔄", t: "Kombinacija solidarnosti + kapitalizacije", opis: "Svi vodeći sustavi imaju i solidarnu komponentu (osnovna mirovina za sve) i kapitalziranu (osobna štednja). HR ima oba, ali oba su slabija od nordijskih modela.", boja: C.navy },
          ].map((it,i) => (
            <div key={i} style={{ background: it.boja+"08", border:`1px solid ${it.boja}22`, borderRadius:9, padding:"12px 14px" }}>
              <div style={{ fontSize:18, marginBottom:6 }}>{it.emoji}</div>
              <div style={{ fontSize:12, fontWeight:700, color:C.navy, marginBottom:4 }}>{it.t}</div>
              <div style={{ fontSize:11, color:C.slate, lineHeight:1.6 }}>{it.opis}</div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

// ── PRAKTIČNI VODIČ ────────────────────────────────────────────────────────
function VodicView() {
  const [dob, setDob] = useState(35);

  const profili = [
    { range: [18,29], label: "18–29 god.", emoji: "🌱", color: C.green,
      naslov: "Mladi radnik — zlatno doba za početak",
      subtitle: "Svaka godina štednje sada vrijedi 3× više nego u 40-im",
      koraci: [
        { emoji: "✅", t: "Provjeri je li poslodavac prijavio u II. stup", opis: "Provjeri e-Građani / REGOS portal da li se svaki mjesec uplaćuje 5% bruto na tvoj osobni račun. Ako ne — poslodavac krši zakon.", hitno: true },
        { emoji: "📂", t: "Odaberi OMF kategoriju A", opis: "Fond kategorije A (visoki udio dionica) je optimalan za mlade — imaš 35+ godina do mirovine da prebrodeš sve tržišne padove. Ako te REGOS rasporedio u A automatski, ostani.", hitno: false },
        { emoji: "🌱", t: "Otvori III. stup ODMAH", opis: `Uplati samo 55,30 EUR/mj. (663,61 EUR/god.) i dobij maksimalni DPS od 99,54 EUR/god. Ukupno s prinosom: na 40 god. to je ~${fmt(calcStup2FV({bruto_placa:55.3*20,god_stednje:40,neto_prinos:0.055}))} EUR.`, hitno: true },
        { emoji: "📊", t: "Razumi PVB — pregovaraj o plaći!", opis: "Svaki % povećanja plaće povećava tvoj PVB i buduću mirovinu. Plaća od 2.000 EUR vs 1.300 EUR = PVB 1,54 vs 1,00 = gotovo 54% veća mirovina.", hitno: false },
        { emoji: "🔍", t: "Provjeri stanje računa jednom godišnje", opis: "e-Građani → REGOS → Moj mirovinski račun. Provjeri ukupan kapital, uplate, prinos i historiju.", hitno: false },
      ],
    },
    { range: [30,44], label: "30–44 god.", emoji: "💼", color: C.blue,
      naslov: "Sredina karijere — optimizacija i planiranje",
      subtitle: "Rast karijere i plaće direktno povećavaju mirovinu",
      koraci: [
        { emoji: "📈", t: "Analiziraj PVB kretanje", opis: "HZMO može dati informativni izračun mirovine s trenutnim PVB-om. Ako si napredovao u karijeri i plaća ti raste, PVB raste — ali provjeri.", hitno: false },
        { emoji: "🏠", t: "Razmotri balans: mirovina vs. nekretnina", opis: "Nekretnina nije mirovinska štednja — ali može biti izvor prihoda (najam) u mirovini. Nemoj zanemariti financijsku mirovinu dok kupuješ nekretninu.", hitno: false },
        { emoji: "🌱", t: "Maksimiziran III. stup", opis: "Ako imaš poslodavca koji uplaćuje u III. stup — to je povećanje plaće bez poreza. Do 804 EUR/god. od poslodavca neoporezivo.", hitno: true },
        { emoji: "📊", t: "Kalkulator: koliko mi nedostaje?", opis: "Koristi tab 'Planer' u ovoj aplikaciji. Unesi željenu mirovinu i vidi točno koliko treba prilagoditi štednju, plaću ili datum umirovljenja.", hitno: false },
        { emoji: "🔄", t: "Provjeri kategoriju OMF", opis: "REGOS automatski prebacuje iz A u B 15 god. nakon prvog zaposlenja. Provjeri jesi li u pravoj kategoriji za svoju dob i procijeni.", hitno: false },
        { emoji: "👶", t: "Porodiljni dopust = mirovinski staž!", opis: "Od 1.7.2025. svako dijete donosi 12 mj. dodanog staža. Ovo se automatski uračunava — ali provjeri da je HZMO evidentirao.", hitno: false },
      ],
    },
    { range: [45,54], label: "45–54 god.", emoji: "⚙️", color: C.gold,
      naslov: "Pred-mirovinski period — zaštita i prijelaz",
      subtitle: "Počnite smanjivati investicijski rizik i planirati isplatu",
      koraci: [
        { emoji: "🔄", t: "Prelazak OMF kategorije A → B", opis: "Ako još imaš visok udio dionica (kategorija A), razmotri prelazak u B. REGOS to automatski radi za tko do mirovine ima <5 god., ali možeš i ranije.", hitno: true },
        { emoji: "📋", t: "Informativni izračun mirovine na HZMO", opis: "Podnesi zahtjev za informativni izračun na mojamirovina.hr (e-Građani). Vidjet ćeš točan PVB, staž i projiciranu mirovinu.", hitno: true },
        { emoji: "⚖️", t: "Analiziraj: istup ili kombinirana mirovina?", opis: "Ako imaš manje od 10 god. u II. stupu ili si ušao dobrovoljno 2002. (bio si 40–50 god.), procijeni isplati li se istup. Razlika može biti značajna.", hitno: false },
        { emoji: "🌱", t: "Intenzivni III. stup ove dekade", opis: "10 godina aktivne štednje uz DPS sada, uz prinos, može dodati 20.000–40.000 EUR kapitala koji je dostupan od 55. god.", hitno: false },
        { emoji: "💼", t: "Planiranje dohotka od imovine u mirovini", opis: "Nekretnine, dividende, umirovljenički fond — planirajte sve izvore prihoda, ne samo državnu mirovinu.", hitno: false },
      ],
    },
    { range: [55,64], label: "55–64 god.", emoji: "🏁", color: C.purple,
      naslov: "Finalna ravnina — maksimizacija i odluke",
      subtitle: "Svaka godina rada sada vrijedi više nego ikad",
      koraci: [
        { emoji: "⏳", t: "Razmotri odgodu — +0,45% po mjesecu!", opis: "Svaki dodatni mjesec rada = trajno +0,45% mirovine. 24 mj. odgode = +10,8% mirovine zauvijek. Promijenite datum samo ako zdravlje i financije to dopuštaju.", hitno: false },
        { emoji: "🏦", t: "Istraži MOD-ove i vrstu mirovine", opis: "Usporedite 4 registrirana MOD-a. Odaberite vrstu mirovine PAŽLJIVO — ona je nepromijenljiva. Ako imate supružnika, razmotrite zajedničku mirovinu ili zajamčeno razdoblje", hitno: true },
        { emoji: "💵", t: "Isplativa jednokratna isplata 20%?", opis: "Provjeri uvjet: I. stup mirovina mora biti > najniža mirovina. 20% kapitala sada vs. viša renta — ovisi o vašim potrebama i zdravlju.", hitno: false },
        { emoji: "💰", t: "Iskoristite III. stup od 55. god.", opis: "III. stup je dostupan od 55. bez obzira radite li. Možete povući dio, sav ili ga prebaciti u redovnu isplatu.", hitno: true },
        { emoji: "📋", t: "Prijavite se HZMO-u najmanje 3 mj. unaprijed", opis: "Postupak ostvarivanja prava traje — nemojte čekati zadnji tren. Zahtjev za starosnu mirovinu podnosi se u nadležnoj ispostavi HZMO-a ili e-Građani.", hitno: true },
        { emoji: "📄", t: "Provjerite sve radne knjižice i staž", opis: "Svaki neevidentirani staž (stari podaci iz SFR), vojska, bolovanje — treba biti u HZMO evidenciji. Greška sada košta zauvijek.", hitno: true },
      ],
    },
  ];

  const profil = profili.find(p => dob >= p.range[0] && dob <= p.range[1]) || profili[1];
  const dobDo = Math.max(0, (profil === profili[3] ? 65 : profil.range[1] + 1) - dob);

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>Praktični vodič po životnoj dobi</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>Što učiniti sada, ovisno o vašoj dobi i fazi karijere</p>
      </div>

      {/* Dob slider */}
      <Card style={{ marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ fontSize: 15 }}>Vaša dob</h3>
          <span style={{ fontSize: 22, fontWeight: 800, color: profil.color,
            fontFamily: "'Playfair Display',serif" }}>{dob} god.</span>
        </div>
        <RangeTrack value={dob} min={18} max={64} step={1} color={profil.color}
          onChange={setDob} />
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 10, gap: 6 }}>
          {profili.map(p => (
            <button key={p.label} onClick={() => setDob(Math.round((p.range[0]+p.range[1])/2))}
              style={{ flex:1, padding:"6px 4px", border:`2px solid ${profil.label===p.label ? p.color : C.border}`,
                borderRadius:8, background: profil.label===p.label ? p.color+"11" : C.white,
                color: profil.label===p.label ? p.color : C.slateL, fontSize:10, fontWeight:600, cursor:"pointer" }}>
              {p.emoji} {p.label}
            </button>
          ))}
        </div>
      </Card>

      {/* Profil */}
      <div style={{ background: `linear-gradient(135deg, ${profil.color}, ${profil.color}bb)`,
        borderRadius: 14, padding: "20px 24px", marginBottom: 20, color: "#fff" }}>
        <div style={{ fontSize: 32, marginBottom: 8 }}>{profil.emoji}</div>
        <h2 style={{ fontSize: 20, marginBottom: 4 }}>{profil.naslov}</h2>
        <p style={{ fontSize: 13, color: "rgba(255,255,255,0.8)", marginBottom: 0 }}>{profil.subtitle}</p>
      </div>

      {/* Koraci */}
      <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 20 }}>
        {profil.koraci.map((k, i) => (
          <div key={i} style={{ display: "flex", gap: 14, padding: "14px 16px",
            background: k.hitno ? profil.color+"0d" : C.white,
            border: `1.5px solid ${k.hitno ? profil.color+"44" : C.border}`,
            borderRadius: 12 }}>
            <span style={{ fontSize: 22, flexShrink: 0, lineHeight: 1 }}>{k.emoji}</span>
            <div style={{ flex: 1 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
                <strong style={{ fontSize: 13, color: C.navy }}>{k.t}</strong>
                {k.hitno && (
                  <span style={{ fontSize: 9, fontWeight: 800, color: profil.color,
                    background: profil.color+"18", padding: "2px 7px", borderRadius: 8,
                    whiteSpace: "nowrap", marginLeft: 8 }}>PRIORITET</span>
                )}
              </div>
              <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.6, margin: 0 }}>{k.opis}</p>
            </div>
          </div>
        ))}
      </div>

      {/* Brzi checklist */}
      <Card style={{ background: C.navy }}>
        <div style={{ color: C.goldL, fontSize: 11, fontWeight: 700, marginBottom: 12,
          textTransform: "uppercase", letterSpacing: "0.08em" }}>✅ Opći checklist za svaku dob</div>
        {[
          ["Provjeri stanje II. stupa na REGOS portalu (e-Građani)", "Jednom godišnje"],
          ["Ažuriraj podatke u HZMO o eventualnom neevidentiranom stažu", "Jednom u životu"],
          ["Otvori ili aktiviraj III. stup (DPS = besplatni novac od države)", "Odmah ako nijesi"],
          ["Provjeri kategoriju OMF — A, B ili C je li primjerena tvojoj dobi", "Jednom godišnje"],
          ["Koristiti 'Moja mirovina' (mojamirovina.hr) za informativni izračun", "Jednom godišnje"],
          ["Planiraj mirovinu kao projekt — ne čekaj zadnji tren", "Sada"],
        ].map(([t,k],i) => (
          <div key={i} style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
            padding:"7px 10px", background: i%2===0 ? "rgba(255,255,255,0.06)" : "transparent",
            borderRadius:6, marginBottom:3 }}>
            <span style={{ fontSize:12, color:"rgba(255,255,255,0.8)" }}>› {t}</span>
            <span style={{ fontSize:10, color:C.goldL, fontWeight:700, whiteSpace:"nowrap", marginLeft:8 }}>{k}</span>
          </div>
        ))}
        <div style={{ marginTop:12, padding:"10px 12px", background:"rgba(255,255,255,0.07)", borderRadius:8,
          fontSize:11, color:"rgba(255,255,255,0.6)" }}>
          Korisni linkovi: <strong style={{ color:C.goldL }}>mojamirovina.hr</strong> (e-Građani) ·
          <strong style={{ color:C.goldL }}> regos.hr</strong> · <strong style={{ color:C.goldL }}>mirovinsko.hr</strong> ·
          <strong style={{ color:C.goldL }}> hanfa.hr</strong>
        </div>
      </Card>
    </div>
  );
}

// ─── MAIN APP ────────────────────────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(null);         // null = guest
  const [tab, setTab] = useState("stup1");
  const [authModal, setAuthModal] = useState(null); // null | reason string
  const [scenariiRefresh, setScenariiRefresh] = useState(0);
  const [appReady, setAppReady] = useState(false);

  // Restore session on mount (non-blocking — app renders immediately)
  useEffect(() => {
    stGet("session:current").then(u => {
      if (u) setUser(u);
      setAppReady(true);
    });
    // Seed demo account
    stGet("user:testuser@hr.com").then(u => {
      if (!u) stSet("user:testuser@hr.com", { name: "Test Korisnik", email: "testuser@hr.com", password: "test123", created: Date.now() });
    });
  }, []);

  async function handleLogin(u) {
    await stSet("session:current", u);
    setUser(u);
    setAuthModal(null);
  }

  async function handleLogout() {
    await stDel("session:current");
    setUser(null);
  }

  function requireAuth(reason) {
    setAuthModal(reason || "Potrebna je prijava.");
  }

  const tabs = [
    { id: "stup1",    icon: <Calculator size={15} />, label: "I. stup" },
    { id: "stup2",    icon: <PiggyBank size={15} />,  label: "II. stup" },
    { id: "stup3",    icon: <TrendingUp size={15} />, label: "III. stup" },
    { id: "ukupno",   icon: <BarChart2 size={15} />,  label: "Ukupno" },
    { id: "invalid",  icon: <Shield size={15} />,     label: "Invalidska" },
    { id: "planer",   icon: <DollarSign size={15} />, label: "Planer" },
    { id: "svijet",   icon: <Users size={15} />,      label: "Svijet" },
    { id: "vodic",    icon: <Clock size={15} />,      label: "Vodič" },
    { id: "edu",      icon: <BookOpen size={15} />,   label: "Edukacija" },
    { id: "blog",     icon: <Star size={15} />,       label: "Blog i AI" },
    { id: "scenariji",icon: <FileText size={15} />,   label: "Scenariji" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.bg }}>
      <style>{css}{SLIDER_CSS}
        {`.blog-card:hover { transform: translateY(-2px); box-shadow: 0 8px 32px rgba(26,46,82,0.12); }`}
      </style>

      {/* Auth modal — renders on top of everything, only when needed */}
      {authModal !== null && (
        <AuthModal
          onLogin={handleLogin}
          onClose={() => setAuthModal(null)}
          reason={authModal}
        />
      )}

      {/* NAV */}
      <nav style={{
        background: C.navy, color: "#fff", padding: "0 24px",
        display: "flex", alignItems: "center", gap: 0,
        boxShadow: "0 2px 16px rgba(0,0,0,0.18)", position: "sticky", top: 0, zIndex: 50,
        height: 58,
      }}>
        <div style={{ fontFamily: "'Playfair Display', serif", fontSize: 18, fontWeight: 700, color: "#fff", marginRight: 32, whiteSpace: "nowrap" }}>
          🏛 Mirovina HR
        </div>
        <div style={{ display: "flex", gap: 2, flex: 1 }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              padding: "0 16px", height: 58, border: "none", background: "none",
              color: tab === t.id ? C.goldL : "rgba(255,255,255,0.65)",
              fontWeight: tab === t.id ? 700 : 400, fontSize: 13,
              borderBottom: tab === t.id ? `2.5px solid ${C.gold}` : "2.5px solid transparent",
              display: "flex", alignItems: "center", gap: 7,
              transition: "all 0.15s", cursor: "pointer", whiteSpace: "nowrap",
            }}>
              {t.icon} {t.label}
            </button>
          ))}
        </div>
        {/* Right: guest or logged-in user */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginLeft: "auto" }}>
          {user ? (
            <>
              <div style={{ fontSize: 13, color: "rgba(255,255,255,0.7)" }}>
                <User size={14} style={{ verticalAlign: "middle", marginRight: 5 }} />
                {user.name}
              </div>
              <button onClick={handleLogout} style={{
                background: "rgba(255,255,255,0.1)", border: "none", borderRadius: 7,
                color: "#fff", padding: "6px 12px", fontSize: 12, cursor: "pointer",
                display: "flex", alignItems: "center", gap: 5,
              }}>
                <LogOut size={13} /> Odjava
              </button>
            </>
          ) : (
            <button onClick={() => requireAuth("")} style={{
              background: `linear-gradient(135deg, ${C.gold}, ${C.goldL})`,
              border: "none", borderRadius: 7, color: C.navy,
              padding: "7px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer",
              display: "flex", alignItems: "center", gap: 5,
            }}>
              <User size={13} /> Prijava / Registracija
            </button>
          )}
        </div>
      </nav>

      {/* AVM Banner */}
      <div style={{ background: `linear-gradient(135deg, ${C.gold}cc, ${C.goldL}aa)`, padding: "8px 24px", fontSize: 12, color: C.navy, display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
        <strong>📊 Aktualni podaci (HZMO):</strong>
        <span>AVM od 1.7.2025. = <strong>14,45 EUR</strong></span>
        <span>|</span>
        <span>Doprinos II. stupa = <strong>5% bruto</strong></span>
        <span>|</span>
        <span>Najniža mirovina (2026.) = <strong>15,73 EUR × staž</strong></span>
        <span>|</span>
        <span>Naknada OMF (2025.) = <strong>0,24% god.</strong></span>
        {!user && (
          <>
            <span>|</span>
            <span style={{ color: C.navy, opacity: 0.75 }}>
              💡 Kalkulator je slobodan · <strong style={{ cursor: "pointer", textDecoration: "underline" }} onClick={() => requireAuth("")}>Prijavi se</strong> samo za spremanje scenarija
            </span>
          </>
        )}
      </div>

      {/* CONTENT */}
      <main style={{ maxWidth: 1100, margin: "0 auto", padding: "28px 20px 60px" }}>
        {tab === "stup1" && <Stup1Calculator user={user} onAuthRequired={requireAuth} onSave={() => setScenariiRefresh(n => n + 1)} />}
        {tab === "stup2" && <Stup2Calculator user={user} onAuthRequired={requireAuth} onSave={() => setScenariiRefresh(n => n + 1)} />}
        {tab === "stup3" && <Stup3View user={user} onAuthRequired={requireAuth} />}
        {tab === "ukupno" && <UkupnoView />}
        {tab === "invalid" && <InvalidskaView />}
        {tab === "planer" && <PlanerView />}
        {tab === "svijet" && <UsporedbaSustaviView />}
        {tab === "vodic" && <VodicView />}
        {tab === "edu" && <EduView />}
        {tab === "scenariji" && <MojiScenarijiView user={user} onAuthRequired={requireAuth} key={scenariiRefresh} />}
        {tab === "blog" && <BlogView />}
      </main>

      {/* Footer */}
      <footer style={{ background: C.navy, color: "rgba(255,255,255,0.5)", padding: "18px 24px", textAlign: "center", fontSize: 11, lineHeight: 1.8 }}>
        <strong style={{ color: "rgba(255,255,255,0.8)" }}>Mirovina HR</strong> — isključivo informativna platforma<br />
        Izvor: HZMO (mirovinsko.hr) · HANFA · mss.gov.hr · ZOMO · ZOMF · AVM 14,45 EUR od 1.7.2025.<br />
        ⚠ Kalkulatori su aproksimativni. Za točan izračun koristite mojamirovina.hr ili se obratite HZMO-u.
      </footer>
    </div>
  );
}
