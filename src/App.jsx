// Mirovina HR – React SPA
// Deps: react, recharts, lucide-react, xlsx (SheetJS)
// Storage: persistent (window.storage API) + in-memory fallback
import { useState, useEffect, useCallback, useRef } from "react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, AreaChart, Area, PieChart, Pie, Cell
} from "recharts";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
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
const AVM = 15.64; // EUR, od 1.7.2026. (HZMO)
const AVM_DATUM = "1.7.2026.";
const PODACI_AZURIRANI = "listopad 2026.";
const MIN_MIROVINA_FAKTOR = 1.06; // najniža mirovina = 106 % AVM po godini staža (ZOMO čl. 80.)
const MAX_VB = 3.8; // najviši prosječni vrijednosni bod (Zakon o najvišoj mirovini)
const PROSJ_BRUTO = 2016; // EUR, prosječna mjesečna bruto plaća 2025. (DZS)
const MIN_PLACA = 1050; // EUR bruto, minimalna plaća 2026.
const FAKTOR_OSNOVNE = 0.75; // faktor osnovne mirovine za staž od 2002. u oba stupa (ZOMO čl. 90.)
const DODATAK = 0.27; // dodatak na mirovinu iz I. stupa (NN 156/23)
const OSOBNI_ODBITAK = 600; // EUR/mj. (ZPD, NN 152/2024)
const POREZNI_PRAG_MJ = 5000; // EUR/mj. — niža stopa do 60.000 EUR/god.

// Porez na mirovinu: stopu određuje grad/općina (zadano 20 % / 30 %), obračunati porez umanjuje se za 50 %.
// Prijedlog za 2027. (u Saboru): umanjenje 100 % za dio oporeziv nižom stopom.
function calcNetoMirovina(bruto, { stopa = 0.20, visaStopa = 0.30, godina = 2026 } = {}) {
  const osnovica = Math.max(0, bruto - OSOBNI_ODBITAK);
  const nizi = Math.min(osnovica, POREZNI_PRAG_MJ) * stopa;
  const visi = Math.max(0, osnovica - POREZNI_PRAG_MJ) * visaStopa;
  const porezBruto = nizi + visi;
  const porezEfektivni = godina >= 2027 ? visi * 0.50 : porezBruto * 0.50;
  return { osnovica, porezBruto, porezEfektivni, neto: bruto - porezEfektivni };
}

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

// Usporedba izbora pri umirovljenju za osiguranike u oba stupa (od 2019. bira svatko):
// "samo I. stup": mirovina kao da je bio samo u I. stupu + 27 %, sredstva II. stupa idu u proračun.
// "I. + II.": osnovna mirovina (staž od 2002. × faktor 0,75) + 27 %, plus renta iz II. stupa.
// Dodatak od 27 % računa se na mirovinu iz formule; ako je s dodatkom niža od najniže, isplaćuje se najniža (bez dodatka).
function saDodatkom(r) {
  const izFormule = Math.min(r.najvisa, r.bruto);
  const s = izFormule * (1 + DODATAK);
  return s >= r.najniza
    ? { ukupno: s, mirovina: izFormule, dodatak: izFormule * DODATAK, najniza: false }
    : { ukupno: r.najniza, mirovina: r.najniza, dodatak: 0, najniza: true };
}

function calcIzbor({ pvb, staz, stazOd2002, pf = 1, mf = 1, avm = AVM, rentaII = 0 }) {
  const post = Math.min(Math.max(0, stazOd2002), staz);
  const samoI = calcStup1({ pvb, staz, pf, mf, avm });
  const osnovna = calcStup1({ pvb, staz: staz - post + post * FAKTOR_OSNOVNE, pf, mf, avm });
  const dI = saDodatkom(samoI);
  const dO = saDodatkom(osnovna);
  const kombinirana = dO.ukupno + rentaII;
  return { samoI, osnovna, dI, dO, samoIUkupno: dI.ukupno, osnovnaSDodatkom: dO.ukupno, kombinirana, razlika: kombinirana - dI.ukupno };
}

// Kombinirana mirovina (I. + II.) kad je cijeli staž od 2002. ili zadani broj godina u II. stupu.
function kombiniranaMirovina({ pvb, staz, stazOd2002 = staz, pf = 1, mf = 1, avm = AVM, rentaII = 0 }) {
  const iz = calcIzbor({ pvb, staz, stazOd2002, pf, mf, avm, rentaII });
  return { osnovna: iz.dO.mirovina, dodatak: iz.dO.dodatak, renta: rentaII, ukupno: iz.kombinirana, samoI: iz.samoIUkupno };
}

const TEKUCA_GODINA = 2026;
const DPS_STOPA = 0.15;
const DPS_MAX_UPLATA_GOD = 663.61;

// Dob za starosnu mirovinu (ZOMO čl. 38. i 170.): muškarci 65; žene 64 god. u 2026., +3 mj. godišnje do 65 u 2030.
function dobZaStarosnu(spol, godRodjenja) {
  if (spol !== "Ž") return 65;
  const prag = { 2026: 64, 2027: 64.25, 2028: 64.5, 2029: 64.75 };
  for (let y = TEKUCA_GODINA; y <= 2030; y++) {
    const d = prag[y] ?? 65;
    if (y - godRodjenja >= d) return d;
  }
  return 65;
}
function dobZaPrijevremenu(spol, godRodjenja) {
  if (spol !== "Ž") return { dob: 60, staz: 35 };
  const prag = { 2026: [59, 34], 2027: [59.25, 34.25], 2028: [59.5, 34.5], 2029: [59.75, 34.75] };
  for (let y = TEKUCA_GODINA; y <= 2030; y++) {
    const [d, st] = prag[y] ?? [60, 35];
    if (y - godRodjenja >= d) return { dob: d, staz: st };
  }
  return { dob: 60, staz: 35 };
}
const fmtDob = (d) => {
  const g = Math.floor(d); const m = Math.round((d - g) * 12);
  return m ? `${g} god. i ${m} mj.` : `${g} god.`;
};

function godineOd2002(godRodjenja, staz, dobUmirovljenja) {
  const godinaUmirovljenja = godRodjenja + dobUmirovljenja;
  return Math.max(0, Math.min(staz, godinaUmirovljenja - 2002));
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
    const post2002 = s.stup1 && s.stup2 ? Math.min(s.stup1.staz, s.stup2.god_stednje) : 0;
    const r1osn = s.stup1 ? calcStup1({ ...s.stup1, staz: s.stup1.staz - post2002 + post2002 * FAKTOR_OSNOVNE }) : null;
    const dod27 = r1osn ? r1osn.efektivna * DODATAK : 0;

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

    const uk = (r1osn ? r1osn.efektivna + dod27 : 0) + renta2;
    const ref_placa = s.stup2 ? s.stup2.bruto_placa : PROSJ_BRUTO;

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
    [`Mirovinski Sustav HR — Izvoz scenarija | ${username} | ${datum}`],
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
      [`Formula: M = PVB × Staž × PF × MF × AVM | ZOMO | AVM = ${fmt(AVM)} EUR (${AVM_DATUM})`],
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
      ["Formula: FV = PMT × [((1+r_mj)^n_mj − 1)/r_mj] | ZOMF | Naknada OMF: 0,23% god. (2026.)"],
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
    [`Polazni AVM: ${fmt(AVM)} EUR (${AVM_DATUM}) | usklađivanje 1.1. i 1.7. (ZOMO čl. 91.)`],
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
    ["Najniža mirovina", "M_min = 1,06 × AVM × Staž × PF × MF", "ZOMO čl. 80.", `Od ${AVM_DATUM}`],
    ["Najviša mirovina", "M_max = Staž × 3,8 × PF × MF × AVM", "Zakon o najvišoj mirovini", ""],
    ["Prijevremena — PF", "PF = 1,00 − 0,002 × broj_mj", "ZOMO čl. 83.", "Smanjenje se ukida sa 70 god."],
    ["Odgoda mirovine — PF", "PF = 1,00 + 0,0045 × broj_mj", "ZOMO čl. 77.", "Trajno povećanje"],
    ["Dodatak na mirovinu", "Dod = 0,27 × osnovna mirovina", "NN 156/2023", "Od 1.1.2024."],
    ["Invalidska mirovina — MF (potpuna)", "MF = 1,10", "ZOMO čl. 48.", "Od 1.1.2026."],
    ["Invalidska mirovina — MF (djelomična)", "MF = 0,90", "ZOMO čl. 48.", "Od 1.1.2026."],
    ["AVM usklađivanje", "Prema rastu plaća i potrošačkih cijena", "ZOMO čl. 91.", "1.1. i 1.7."],
    ["II. stup — akumulacija", "FV = PMT × [((1+r_mj)^n_mj − 1) / r_mj]", "ZOMF", "Geometrijska renta"],
    ["II. stup — doprinos", "5% bruto plaće na osobni račun", "Zakon o doprinosima", "Uplaćuje poslodavac"],
    ["II. stup — naknade OMF", "0,23% god. imovine fonda (2026.)", "ZOMF", "Pada do 0,20% od 2029."],
    ["III. stup — DPS", "DPS = 15% × god. uplate, max 99,54 EUR/god.", "ZDOBF čl. 135.", ""],
    ["III. stup — poslodavac", "Max 804 EUR/god. neoporezivo", "Zakon o porezu na dohodak", "Beneficija za zaposlenika"],
    [],
    ["AVM POVIJEST"],
    ["Datum", "AVM (EUR)", "Promjena"],
    ["1.7.2025.", xlCell(14.45, { fmt: "#,##0.00" }), "+6,48%"],
    ["1.1.2026.", xlCell(14.84, { fmt: "#,##0.00" }), "+2,68%"],
    ["1.7.2026.", xlCell(15.64, { fmt: "#,##0.00" }), "+5,40%"],
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

// ─── PDF EXPORT (bez registracije) ──────────────────────────────────────────
const PDF_DISCLAIMER =
  "Ovo je informativni izračun generiran neovisnim alatom Mirovinski Sustav HR i ne predstavlja " +
  "službeni izračun mirovine niti obvezujući dokument. Za točan i pravno valjan izračun obratite se " +
  "HZMO-u (mirovinsko.hr) odnosno za II./III. stup HANFA-i (hanfa.hr) ili na mojamirovina.hr. " +
  "Korišteni podaci i formule temelje se na ZOMO/ZOMF/ZDOBF te AVM = 15,64 EUR (od 1.7.2026.) i podložni su promjenama.";

// Roboto (latin + latin-ext) — uključuje sva hrvatska slova (č, ć, š, ž, đ)
let _pdfFontsPromise = null;
async function loadPdfFonts(doc) {
  if (!_pdfFontsPromise) {
    _pdfFontsPromise = (async () => {
      const toBase64 = buf => {
        let binary = "";
        const bytes = new Uint8Array(buf);
        for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
        return btoa(binary);
      };
      const [regular, bold] = await Promise.all([
        fetch("/fonts/Roboto-Regular.ttf").then(r => r.arrayBuffer()),
        fetch("/fonts/Roboto-Bold.ttf").then(r => r.arrayBuffer()),
      ]);
      return { regular: toBase64(regular), bold: toBase64(bold) };
    })();
  }
  const fonts = await _pdfFontsPromise;
  doc.addFileToVFS("Roboto-Regular.ttf", fonts.regular);
  doc.addFileToVFS("Roboto-Bold.ttf", fonts.bold);
  doc.addFont("Roboto-Regular.ttf", "Roboto", "normal");
  doc.addFont("Roboto-Bold.ttf", "Roboto", "bold");
}

async function exportSummaryToPDF({ title, subtitle, rows, formule }) {
  const doc = new jsPDF();
  await loadPdfFonts(doc);
  doc.setFont("Roboto", "normal");
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  let y = 18;

  // Header
  doc.setFillColor(26, 46, 82); // C.navy
  doc.rect(0, 0, pageW, 30, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("Roboto", "bold");
  doc.setFontSize(15);
  doc.text("Mirovinski Sustav HR", 14, 13);
  doc.setFont("Roboto", "normal");
  doc.setFontSize(11);
  doc.text(title, 14, 22);
  if (subtitle) {
    doc.setFontSize(9);
    doc.setTextColor(220, 226, 240);
    doc.text(subtitle, 14, 28);
  }

  y = 40;
  doc.setTextColor(120, 120, 120);
  doc.setFontSize(9);
  doc.text(`Izračun generiran: ${new Date().toLocaleString("hr-HR")}`, 14, y);
  y += 9;

  // Rows table
  doc.setFontSize(11);
  rows.forEach(([label, value]) => {
    doc.setFont("Roboto", "normal");
    doc.setTextColor(80, 80, 80);
    doc.text(String(label), 14, y);
    doc.setFont("Roboto", "bold");
    doc.setTextColor(20, 20, 20);
    doc.text(String(value), pageW - 14, y, { align: "right" });
    doc.setDrawColor(230, 230, 230);
    doc.line(14, y + 2.5, pageW - 14, y + 2.5);
    y += 9;
  });

  // Formule i parametri
  if (formule && formule.length) {
    y += 6;
    doc.setFont("Roboto", "bold");
    doc.setFontSize(11);
    doc.setTextColor(26, 46, 82);
    doc.text("Formule i parametri", 14, y);
    y += 7;
    doc.setFont("Roboto", "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(70, 70, 70);
    formule.forEach(line => {
      const split = doc.splitTextToSize(line, pageW - 28);
      if (y + split.length * 4.5 > pageH - 30) { doc.addPage(); y = 18; }
      doc.text(split, 14, y);
      y += split.length * 4.5 + 3;
    });
  }

  // Disclaimer footer na svakoj stranici
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFillColor(245, 246, 250);
    doc.rect(0, pageH - 24, pageW, 24, "F");
    doc.setFont("Roboto", "normal");
    doc.setFontSize(7);
    doc.setTextColor(110, 110, 110);
    doc.text(doc.splitTextToSize(PDF_DISCLAIMER, pageW - 28), 14, pageH - 18);
    doc.setFontSize(7);
    doc.text(`Stranica ${i} / ${pageCount}`, pageW - 14, pageH - 4, { align: "right" });
  }

  doc.save(`${title.replace(/[^a-zA-Z0-9čćžšđČĆŽŠĐ]+/g, "_")}_${new Date().toISOString().slice(0,10)}.pdf`);
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

Od 2002. godine zaposleni koji su u II. stupu (obvezno svi koji su tada bili mlađi od 40 i svi kasnije zaposleni) uplaćuju **5% bruto plaće** na osobni račun u Obveznom mirovinskom fondu (OMF). Ova sredstva nisu zajednička — vaša su osobna imovina, naslijediva i odvojena od državnog proračuna.

## Kako raste vaš kapital?

OMF ulaže vaše uplate u financijske instrumente (obveznice, dionice, nekretninski fondovi). Prinos nije zajamčen; dugoročni prosječni prinos OMF-ova kreće se oko **5% godišnje** (Hanfa).

Formula akumulacije:
> **FV = PMT × [((1 + r_mj)^n − 1) / r_mj]**

gdje je PMT = mjesečna uplata, r_mj = neto prinos/12, n = broj mjeseci.

**Primjer:** Osoba s prosječnom bruto plaćom 2.016 EUR, koja štedi 33 godine uz 5,5% neto prinos, akumulira **~112.500 EUR**.

## Naknade koje ne vidite

Svake godine OMF naplaćuje naknadu za upravljanje. HANFA propisuje:
- 2026.: **0,23%** godišnje od ukupne imovine fonda (2027.: 0,22%)
- 2029. i dalje: **0,20%**

Na 33 godine štednje, razlika između 3% i 7% prinosa uz plaću od 2.016 EUR/mj. = razlika **~87.000 EUR** u akumuliranom kapitalu.

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

Trenutno nijedna hrvatska opcija ne vraća neisplaćeni kapital nasljednicima ako živite dulje od aktuarskog prosjeka. **Ovo je ključna razlika od osiguranja s povratom premije.**`
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
3. Malo tržišne konkurencije — samo 2 MOD-a`
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
    excerpt: "Za svaki mjesec ranijeg odlaska u mirovinu polazni faktor se smanjuje za 0,2 %. Na 36 mjeseci to je 7,2 % manja mirovina sve do 70. godine.",
    category: "Strategija", readMin: 7, date: "2025-09-25",
    tags: ["prijevremena mirovina", "polazni faktor", "strategija"],
    content: `## Kazna za rani odlazak

Za svaki **mjesec** ranijeg odlaska u mirovinu (u odnosu na zakonsku dob), polazni faktor se smanjuje za **0,2%**.

> **PF = 1,00 − 0,002 × broj_mj_ranijeg_odlaska**

## Primjeri

| Rani odlazak | PF | Smanjenje mirovine (do 70. god.) |
|---|---|---|
| 6 mj. | 0,988 | −1,2% |
| 12 mj. | 0,976 | −2,4% |
| 24 mj. | 0,952 | −4,8% |
| 36 mj. | 0,928 | −7,2% |
| 60 mj. | 0,880 | −12,0% |

## Puna mirovinska dob 2025.

Muškarci: **65 godina**
Žene: **64 god.** u 2026., zatim +3 mj. godišnje do 65 god. u 2030.

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
  @media (max-width: 760px) {
    main [style*="grid-template-columns: 1fr 1fr"],
    main [style*="grid-template-columns: repeat(3"],
    main [style*="grid-template-columns: repeat(4"],
    main [style*="grid-template-columns: 2fr 3fr 1fr"] { grid-template-columns: 1fr !important; }
    main { padding: 18px 12px 60px !important; }
    .main-nav, .sub-nav { position: static !important; padding: 0 12px !important; }
    .grp-btn { height: 40px !important; padding: 0 9px !important; font-size: 12px !important; }
    .data-banner { display: none !important; }
  }
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

function Card({ children, style = {}, className = "", ...rest }) {
  return (
    <div className={className} style={{
      background: C.white, borderRadius: 14, padding: 24,
      boxShadow: "0 2px 16px rgba(26,46,82,0.07)", border: `1px solid ${C.border}`,
      ...style
    }} {...rest}>{children}</div>
  );
}

function Ukratko({ children, color = C.blue }) {
  return (
    <div style={{ marginBottom: 18, padding: "14px 18px", background: C.white, borderRadius: 12,
      borderLeft: `5px solid ${color}`, boxShadow: "0 2px 10px rgba(26,46,82,0.06)" }}>
      <div style={{ fontSize: 11, fontWeight: 800, color, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>Ukratko</div>
      <div style={{ fontSize: 13, color: C.navy, lineHeight: 1.65 }}>{children}</div>
    </div>
  );
}

function Detalji({ naslov = "Detaljnije objašnjenje", children, otvoreno = false, style = {} }) {
  const [open, setOpen] = useState(otvoreno);
  return (
    <div style={{ marginTop: 14, border: `1px solid ${C.border}`, borderRadius: 10, background: C.white, ...style }}>
      <button onClick={() => setOpen(o => !o)} aria-expanded={open}
        style={{ width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center",
          padding: "10px 14px", background: "none", border: "none", fontSize: 13, fontWeight: 600, color: C.navy, textAlign: "left" }}>
        <span>{naslov}</span>
        {open ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
      </button>
      {open && <div style={{ padding: "0 14px 14px" }}>{children}</div>}
    </div>
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

// ─── FEEDBACK MODAL ──────────────────────────────────────────────────────────
function FeedbackModal({ onClose }) {
  const [tip, setTip] = useState("greska");
  const [poruka, setPoruka] = useState("");
  const [email, setEmail] = useState("");
  const [sent, setSent] = useState(false);

  const inp = { width: "100%", padding: "11px 14px", border: `1.5px solid ${C.border}`,
    borderRadius: 9, fontSize: 14, outline: "none", background: "#fafbff", fontFamily: "inherit" };

  function handleSubmit() {
    if (!poruka.trim()) return;
    const tipLabel = { greska: "Greška", prijedlog: "Prijedlog", komentar: "Komentar" }[tip];
    const subject = `[Mirovina HR] ${tipLabel}`;
    const body = `${poruka}${email ? `\n\nKontakt e-adresa: ${email}` : ""}`;
    window.location.href = `mailto:hjelic@gmail.com?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    setSent(true);
  }

  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 200,
      background: "rgba(15,25,50,0.55)", backdropFilter: "blur(4px)",
      display: "flex", alignItems: "center", justifyContent: "center", padding: 20,
    }} onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="fadeIn" style={{ width: "100%", maxWidth: 440, background: C.white, borderRadius: 18, padding: 32, boxShadow: "0 24px 80px rgba(0,0,0,0.3)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <div style={{ fontSize: 24, marginBottom: 4 }}>💬</div>
            <h2 style={{ fontSize: 20, marginBottom: 4 }}>Prijavi grešku ili komentar</h2>
            <p style={{ fontSize: 12, color: C.slateL, lineHeight: 1.5 }}>
              Otvorit će se tvoj e-mail klijent s pripremljenom porukom.
            </p>
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: C.slateL, padding: 4, cursor: "pointer" }}>
            <X size={20} />
          </button>
        </div>

        {sent ? (
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 32, marginBottom: 10 }}>✅</div>
            <p style={{ fontSize: 14, color: C.slate, lineHeight: 1.6, marginBottom: 18 }}>
              Hvala! Provjeri da se e-mail otvorio u tvojoj aplikaciji i pošalji ga.
            </p>
            <button onClick={onClose} style={{
              padding: "10px 24px", background: C.navy, border: "none", borderRadius: 9,
              color: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer",
            }}>Zatvori</button>
          </div>
        ) : (
          <>
            <div style={{ display: "flex", marginBottom: 16, background: C.bg, borderRadius: 9, padding: 3, gap: 0 }}>
              {[["greska","Greška"],["prijedlog","Prijedlog"],["komentar","Komentar"]].map(([v, l]) => (
                <button key={v} onClick={() => setTip(v)}
                  style={{ flex: 1, padding: "8px 0", border: "none", borderRadius: 7,
                    background: tip === v ? C.navy : "transparent",
                    color: tip === v ? "#fff" : C.slateL, fontWeight: 600, fontSize: 12,
                    transition: "all 0.2s", cursor: "pointer" }}>
                  {l}
                </button>
              ))}
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 11 }}>
              <textarea style={{ ...inp, minHeight: 110, resize: "vertical" }}
                placeholder="Opiši grešku, prijedlog ili komentar…"
                value={poruka} onChange={e => setPoruka(e.target.value)} />
              <input style={inp} placeholder="Tvoja e-adresa (neobavezno, za odgovor)"
                type="email" value={email} onChange={e => setEmail(e.target.value)} />
            </div>

            <button onClick={handleSubmit} disabled={!poruka.trim()} style={{
              marginTop: 18, width: "100%", padding: "12px 0",
              background: !poruka.trim() ? C.slateL : `linear-gradient(135deg, ${C.gold}, ${C.goldL})`,
              border: "none", borderRadius: 10, color: C.navy, fontWeight: 700, fontSize: 14,
              boxShadow: "0 4px 14px rgba(201,146,43,0.3)", cursor: !poruka.trim() ? "default" : "pointer"
            }}>
              Pošalji
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function NetoMirovinaBox({ bruto, napomena }) {
  const [stopa, setStopa] = useState(0.20);
  const t26 = calcNetoMirovina(bruto, { stopa });
  const t27 = calcNetoMirovina(bruto, { stopa, godina: 2027 });
  return (
    <div style={{ marginTop: 12, padding: "12px 14px", background: C.greenL, borderRadius: 10, border: `1.5px solid ${C.green}33` }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.green }}>💶 Koliko ostaje na računu (neto)</div>
        <label style={{ fontSize: 10, color: C.slate }}>
          Stopa poreza vašeg grada:{" "}
          <select value={stopa} onChange={e => setStopa(Number(e.target.value))} style={{ fontSize: 11, padding: "2px 4px", borderRadius: 4, border: `1px solid ${C.border}` }}>
            {[0.15, 0.17, 0.18, 0.20, 0.21, 0.22, 0.23].map(v => <option key={v} value={v}>{Math.round(v * 100)} %</option>)}
          </select>
        </label>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 10, color: C.slateL, marginBottom: 2 }}>Bruto</div>
          <div style={{ fontSize: 14, fontWeight: 700, color: C.navy }}>{fmt(bruto)} EUR</div>
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 10, color: C.slateL, marginBottom: 2 }}>Neto 2026.</div>
          <div style={{ fontSize: 16, fontWeight: 800, color: C.green }}>{fmt(t26.neto)} EUR</div>
          <div style={{ fontSize: 9, color: C.slateL }}>porez {fmt(t26.porezEfektivni)} EUR</div>
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 10, color: C.slateL, marginBottom: 2 }}>Neto od 2027.*</div>
          <div style={{ fontSize: 16, fontWeight: 800, color: C.green }}>{fmt(t27.neto)} EUR</div>
          <div style={{ fontSize: 9, color: C.slateL }}>porez {fmt(t27.porezEfektivni)} EUR</div>
        </div>
      </div>
      <div style={{ fontSize: 10, color: C.slate, lineHeight: 1.6 }}>
        Do {OSOBNI_ODBITAK} EUR mirovine porez se ne plaća (osobni odbitak). Na ostatak se obračunava porez po stopi vašeg
        grada ili općine, a umirovljenici plaćaju samo pola tog iznosa. Prirez ne postoji od 2024.
        {napomena}
        <span style={{ display: "block", marginTop: 6, padding: "4px 8px", background: C.amberL, borderRadius: 6, color: C.amber, fontWeight: 600 }}>
          * Vlada je u rujnu 2026. uputila Saboru prijedlog da se od 1.1.2027. ukine porez na mirovine (do 5.000 EUR osnovice). Iznos vrijedi ako prijedlog bude usvojen.
        </span>
      </div>
    </div>
  );
}

// ─── STUP 1 KALKULATOR ───────────────────────────────────────────────────────
function Stup1Calculator({ onSave, user, onAuthRequired }) {
  const [params, setParams] = useState({
    staz: 40, pvb: 1.0, vrsta: "starosna",
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
    : params.vrsta === "invalidska_potpuna" ? 1.1
    : params.vrsta === "invalidska_djelomicna" ? 0.9 : 1.0;

  const res = calcStup1({ pvb: params.pvb, staz: params.staz, pf, mf, avm: params.avm });
  const sd = saDodatkom(res);

  const cmp_data = [
    { name: "Prijevremena (-36 mj.)", value: saDodatkom(calcStup1({ pvb: params.pvb, staz: Math.max(1, params.staz - 3), pf: 1 - 0.002 * 36, mf: 1.0, avm: params.avm })).ukupno },
    { name: "Standardna (puna dob)", value: saDodatkom(calcStup1({ pvb: params.pvb, staz: params.staz, pf: 1.0, mf: 1.0, avm: params.avm })).ukupno },
    { name: "Vaš scenarij", value: sd.ukupno },
    { name: "Odgoda (+24 mj.)", value: saDodatkom(calcStup1({ pvb: params.pvb, staz: params.staz + 2, pf: 1 + 0.0045 * 24, mf: 1.0, avm: params.avm })).ukupno },
    { name: "Odgoda (+48 mj.)", value: saDodatkom(calcStup1({ pvb: params.pvb, staz: params.staz + 4, pf: 1 + 0.0045 * 48, mf: 1.0, avm: params.avm })).ukupno },
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

  async function handleExportPDF() {
    await exportSummaryToPDF({
      title: "I. mirovinski stup — izračun mirovine",
      subtitle: naziv,
      rows: [
        ["Vrsta mirovine", vrsteOptions.find(o => o.v === params.vrsta)?.l || params.vrsta],
        ["Mirovinski staž", `${params.staz} god.`],
        ["Prosječni vrijednosni bodovi (PVB)", params.pvb.toFixed(2)],
        ["Polazni faktor (PF)", pf.toFixed(3)],
        ["Mirovinski faktor (MF)", mf.toFixed(2)],
        ["AVM", `${params.avm.toFixed(2)} EUR`],
        ["Osobni bodovi (OB)", res.ob.toFixed(3)],
        ["Bruto mirovina", `${fmt(res.bruto)} EUR`],
        ["Najniža mirovina", `${fmt(res.najniza)} EUR`],
        ["Najviša mirovina", `${fmt(res.najvisa)} EUR`],
        ["Mirovina iz formule (u granicama najniže/najviše)", `${fmt(res.efektivna)} EUR`],
        ["Mirovina s dodatkom 27 % / mj.", `${fmt(sd.ukupno)} EUR`],
        ["Mirovina s dodatkom 27 % / god.", fmtEur(sd.ukupno * 12)],
      ],
      formule: [
        `OB = PVB × Staž × PF = ${params.pvb.toFixed(2)} × ${params.staz} × ${pf.toFixed(3)} = ${res.ob.toFixed(3)}`,
        `Bruto = OB × MF × AVM = ${res.ob.toFixed(3)} × ${mf.toFixed(2)} × ${params.avm.toFixed(2)} = ${fmt(res.bruto)} EUR`,
        `Najniža = ${MIN_MIROVINA_FAKTOR} × AVM × Staž × PF × MF = ${fmt(res.najniza)} EUR`,
        `Najviša = Staž × ${MAX_VB} × PF × MF × AVM = ${fmt(res.najvisa)} EUR`,
        `Efektivna = MAX(Najniža, MIN(Najviša, Bruto)) = ${fmt(res.efektivna)} EUR`,
      ],
    });
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
                  { v: "invalidska_potpuna", emoji: "⚕️", l: "Invalidska — potpuna", sub: "Potpuni gubitak radne sposobnosti (MF 1,1)" },
                  { v: "invalidska_djelomicna", emoji: "🩺", l: "Invalidska — djelomična", sub: "Djelomični gubitak radne sposobnosti (MF 0,9)" },
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
              const est = (v * AVM).toFixed(0);
              return `S ${v} godina staža i prosječnom plaćom (PVB 1,0) okvirna bruto mirovina iz I. stupa je oko ${est} EUR/mj. Svaka dodatna godina staža donosi oko ${AVM.toFixed(0)} EUR/mj. više.`;
            }}
            explain={
              <div>
                <p><strong>Staž</strong> je ukupan broj godina koje ste proveli radeći i uplaćujući doprinose u mirovinski sustav.</p>
                <p style={{ marginTop: 6 }}>Svaka godina staža direktno povećava mirovinu — u formuli je <em>M = PVB × <strong>Staž</strong> × PF × MF × AVM</em>, dakle što je staž veći, to je mirovina veća, linearno.</p>
                <p style={{ marginTop: 6 }}>Računa se i vojni rok, bolovanje i rodiljni dopust. Staž se zbraja iz svih radnih odnosa u životu, a majkama se dodaje 12 mjeseci za svako dijete (od 1.7.2025.).</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.blueL, borderRadius: 6 }}>
                  💡 <strong>Praktično:</strong> 5 dodatnih godina rada uz prosječnu plaću = oko {(5 * AVM).toFixed(0)} EUR bruto više mirovine svaki mjesec.
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
              const avgHR = PROSJ_BRUTO;
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
                  💡 <strong>Primjer:</strong> Zarađivali ste 3.024 EUR bruto, a prosjek RH bio 2.016 EUR → bod te godine = 1,50. Ako je to bio prosjek vaše karijere, PVB = 1,50.
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
                return `Odlaskom ${v} mj. ranije mirovina je manja za ${kazna} % sve do 70. godine života, kada se smanjenje ukida. Polazni faktor PF = ${pf_calc.toFixed(3)}.`;
              }}
              explain={
                <div>
                  <p>Za svaki <strong>mjesec ranijeg odlaska</strong> u mirovinu, mirovina se trajno smanjuje za <strong>0,2%</strong>.</p>
                  <p style={{ marginTop: 6 }}>Smanjenje vrijedi do navršene 70. godine života. Formula: <em>PF = 1,00 − 0,002 × broj mjeseci</em>. Uvjet za prijevremenu mirovinu u 2026.: muškarci 60 god. i 35 god. staža, žene 59 god. i 34 god. staža.</p>
                  <div style={{ marginTop: 8, padding: "6px 10px", background: C.redL, borderRadius: 6 }}>
                    ⚠ <strong>Pažnja:</strong> 3 godine ranije = −7,2 % mirovine sve do 70. godine (uz to i 3 godine staža manje).
                  </div>
                  <div style={{ marginTop: 8, padding: "6px 10px", background: C.greenL, borderRadius: 6 }}>
                    ✅ <strong>Od 1.1.2026. (ZOMO čl. 83.):</strong> kad korisnik prijevremene mirovine navrši <strong>70 godina</strong>, HZMO po službenoj dužnosti ukida smanjenje od sljedećeg mjeseca.
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
              { upTo: 15.64,color: C.green, label: `Sadašnja razina (${AVM_DATUM})` },
              { upTo: 17,   color: C.blue,  label: "Projicirana buduća razina" },
              { upTo: 20,   color: C.purple,label: "Optimistična projekcija" },
            ]}
            dynamicExplain={v => {
              if (Math.abs(v - AVM) < 0.02) return `AVM ${v.toFixed(2)} EUR — ovo je sadašnja vrijednost (od ${AVM_DATUM}). Svaki vaš osobni bod vrijedi ${v.toFixed(2)} EUR mirovine.`;
              const razlika = ((v - AVM) / AVM * 100).toFixed(1);
              return `AVM ${v.toFixed(2)} EUR = ${razlika > 0 ? '+' : ''}${razlika}% vs. sadašnja vrijednost. Ovim simulirate buduće usklađivanje AVM.`;
            }}
            explain={
              <div>
                <p><strong>AVM (aktualna vrijednost mirovine)</strong> je "cijena jednog boda". HZMO je usklađuje dva puta godišnje (1.1. i 1.7.) prema rastu plaća i cijena (ZOMO čl. 91.).</p>
                <p style={{ marginTop: 6 }}>Vaša mirovina = <em>Osobni bodovi × AVM</em>. Ako AVM poraste za 5%, vaša mirovina poraste za 5%.</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.blueL, borderRadius: 6 }}>
                  💡 Sadašnja vrijednost: <strong>{fmt(AVM)} EUR/bod</strong> (od {AVM_DATUM}). Promijenite ovaj slider da vidite kako bi mirovina izgledala u budućnosti s drugačijim AVM-om.
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
                {mf === 1.0 ? "Puna mirovina" : mf > 1 ? "Uvećana (invalidska, potpuni gubitak)" : "Umanjena (invalidska, djelomični gubitak)"}
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
              <ResultBox label="✅ Mirovina s dodatkom od 27 % / mj." value={`${fmt(sd.ukupno)} EUR`} color={C.green} big
                sub={sd.najniza ? "isplaćuje se najniža mirovina" : `${fmt(sd.mirovina)} EUR + dodatak ${fmt(sd.dodatak)} EUR`} />
            </div>
            <div style={{ marginTop: 10, textAlign: "center", fontSize: 12, color: C.slateL }}>
              Godišnje: <strong>{fmtEur(sd.ukupno * 12)}</strong>
            </div>
            <div style={{ marginTop: 10, padding: "8px 12px", background: C.purpleL, borderRadius: 8, fontSize: 11, color: C.slate, lineHeight: 1.6 }}>
              Ovo je mirovina ako ste samo u I. stupu (ili pri umirovljenju izaberete samo I. stup). Ako ste u II. stupu i izaberete
              mirovinu iz oba stupa, dio iz I. stupa je niži, ali dobivate i mirovinu iz II. stupa. Usporedbu radi kalkulator „Samo I. ili I. + II.“.
            </div>
            <div style={{ marginTop: 12, padding: "10px 12px", background: C.bg, borderRadius: 8, fontSize: 11, color: C.slate, lineHeight: 1.6 }}>
              <strong>📐 Najniža i najviša mirovina:</strong> zakon jamči najmanje {fmt(res.najniza)} EUR za vaš staž
              ({fmt(MIN_MIROVINA_FAKTOR * params.avm)} EUR po godini staža) i ograničava mirovinu na najviše {fmt(res.najvisa)} EUR
              (kao da ste cijeli radni vijek zarađivali {String(MAX_VB).replace(".", ",")} prosječne plaće).{" "}
              {res.bruto >= res.najvisa
                ? <>Vaš izračun premašuje gornju granicu, pa se isplaćuje <strong>najviša mirovina</strong>.</>
                : res.bruto < res.najniza
                  ? <>Vaš izračun je ispod donje granice, pa se isplaćuje <strong>najniža mirovina</strong>.</>
                  : <>Vaš izračun je između tih granica, pa se isplaćuje točno izračunani iznos.</>}
            </div>

            <NetoMirovinaBox bruto={sd.ukupno} />
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
              <button onClick={handleExportPDF} title="Izvezi izračun kao PDF" style={{
                padding: "8px 14px", background: C.white, color: C.navy,
                border: `1.5px solid ${C.border}`, borderRadius: 8, fontWeight: 600, fontSize: 13,
                display: "flex", alignItems: "center", gap: 6,
              }}>
                <FileText size={15} />
                PDF
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
            ["Bruto mirovina", `${res.ob.toFixed(3)} × ${mf.toFixed(2)} × ${params.avm.toFixed(2)} = ${fmt(res.bruto)} EUR`],
            ["Najniža mirovina", `${MIN_MIROVINA_FAKTOR} × ${params.avm.toFixed(2)} × ${params.staz} × ${pf.toFixed(3)} × ${mf.toFixed(2)} = ${fmt(res.najniza)} EUR`],
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
    bruto_placa: PROSJ_BRUTO, god_stednje: 33,
    neto_prinos: 0.055, e65: 18,
    jednokratni_pct: 0.20,
    zomf_varijanta: "standard",
  });
  const [naziv, setNaziv] = useState("Moj II. stup scenarij");
  const [saved, setSaved] = useState(false);

  const kapital = calcStup2FV(params);
  const jednokratni_iznos = kapital * params.jednokratni_pct;
  const preostali = kapital * (1 - params.jednokratni_pct);
  const renta_standard = rentaDozivotna(preostali, params.e65);

  // Projekcija akumulacije
  const proj_data = Array.from({ length: Math.ceil(params.god_stednje / 5) + 1 }, (_, i) => {
    const god = i * 5;
    if (god > params.god_stednje) return null;
    const k = calcStup2FV({ ...params, god_stednje: god });
    const uplata = params.bruto_placa * 0.05 * god * 12;
    return { god: `${god}g`, kapital: Math.round(k), uplate: Math.round(uplata), prinos: Math.round(k - uplata) };
  }).filter(Boolean);

  // Financijska analiza opcija isplate — koliko se ukupno "izvuče" iz preostalog kapitala
  const ukupno_standard = renta_standard * 12 * params.e65;
  const nasljedstvo_standard = 0;

  // Stvarne opcije isplate prema ZOMF — 2 osnovna modela (pojedinačna / obiteljska) ×
  // 3 razine jamstva (bez jamstva, 5 god., 10 god.) = 6 kombinacija koje MOD mora ponuditi.
  const G5 = 0.985;  // procjena: smanjenje rente za zajamčeno razdoblje isplate od 5 god.
  const G10 = 0.95;  // procjena: smanjenje rente za zajamčeno razdoblje isplate od 10 god.
  const OBIT = 0.88; // procjena: smanjenje rente za doživotni nastavak isplate bračnom drugu

  const renta_pojedinacna = renta_standard;
  const renta_pojedinacna_g5 = renta_standard * G5;
  const renta_pojedinacna_g10 = renta_standard * G10;
  const renta_obiteljska = renta_standard * OBIT;
  const renta_obiteljska_g5 = renta_obiteljska * G5;
  const renta_obiteljska_g10 = renta_obiteljska * G10;

  // DCF — diskontiranje budućih isplata na sadašnju vrijednost (stopa = pretpostavljeni neto prinos fonda)
  const DISCOUNT_R = params.neto_prinos;
  const annuityFactor = (r, n) => r > 0 ? (1 - Math.pow(1 + r, -n)) / r : n;
  const AF = annuityFactor(DISCOUNT_R, params.e65);

  const varijante_stvarne = [
    { key: "standard", l: "Pojedinačna doživotna mirovina", short: "Pojedinačna", renta: renta_pojedinacna, garancija_god: null,
      sub: "Isplaćuje se dok je korisnik živ. Prestaje njegovom smrću — najviša mjesečna renta.",
      nasljedstvo_detalj: "Smrću korisnika isplata trajno prestaje. Preostala sredstva ostaju u zajedničkom riziko-pulu MOD-a (koriste se za isplate drugim korisnicima koji žive dulje od prosjeka). Nasljednici ne primaju ništa od sredstava II. stupa." },
    { key: "pojedinacna_g5", l: "Pojedinačna doživotna mirovina sa zajamčenim razdobljem 5 god.", short: "Pojedinačna + 5g jamstvo", renta: renta_pojedinacna_g5, garancija_god: 5,
      sub: "Doživotna renta uz jamstvo isplate prvih 5 god. nasljednicima ako korisnik umre ranije. Renta ~1,5% niža.",
      nasljedstvo_detalj: "Ako korisnik umre u prvih 5 godina primanja mirovine, imenovani korisnik (nasljednik) prima preostale mjesečne iznose do kraja 5. godine. Nakon isteka tog razdoblja nasljednici ne primaju ništa, bez obzira koliko dugo bi korisnik inače živio." },
    { key: "pojedinacna_g10", l: "Pojedinačna doživotna mirovina sa zajamčenim razdobljem 10 god.", short: "Pojedinačna + 10g jamstvo", renta: renta_pojedinacna_g10, garancija_god: 10,
      sub: "Doživotna renta uz jamstvo isplate prvih 10 god. nasljednicima ako korisnik umre ranije. Renta ~5% niža.",
      nasljedstvo_detalj: "Ako korisnik umre u prvih 10 godina primanja mirovine, imenovani korisnik (nasljednik) prima preostale mjesečne iznose do kraja 10. godine. Nakon isteka tog razdoblja nasljednici ne primaju ništa." },
    { key: "obiteljska", l: "Obiteljska (zajednička) doživotna mirovina", short: "Obiteljska", renta: renta_obiteljska, garancija_god: null,
      sub: "Nakon smrti korisnika, mirovina se nastavlja doživotno isplaćivati bračnom drugu. Renta ~12% niža.",
      nasljedstvo_detalj: "Nakon smrti korisnika, isplata se nastavlja doživotno bračnom drugu (ako je živ), u istom ili ugovorenom iznosu. Kada i bračni drug umre, isplata trajno prestaje — drugi nasljednici (npr. djeca) ne primaju ništa od sredstava II. stupa." },
    { key: "obiteljska_g5", l: "Obiteljska doživotna mirovina sa zajamčenim razdobljem 5 god.", short: "Obiteljska + 5g jamstvo", renta: renta_obiteljska_g5, garancija_god: 5,
      sub: "Nastavak isplate bračnom drugu doživotno + jamstvo prvih 5 god. nasljednicima ako oboje umru ranije.",
      nasljedstvo_detalj: "Isplata se nakon smrti korisnika nastavlja bračnom drugu doživotno. Ako oboje (korisnik i bračni drug) umru u prvih 5 godina primanja mirovine, preostale mjesečne iznose do kraja 5. godine prima imenovani nasljednik. Nakon isteka razdoblja, nasljednici ne primaju ništa." },
    { key: "obiteljska_g10", l: "Obiteljska doživotna mirovina sa zajamčenim razdobljem 10 god.", short: "Obiteljska + 10g jamstvo", renta: renta_obiteljska_g10, garancija_god: 10,
      sub: "Nastavak isplate bračnom drugu doživotno + jamstvo prvih 10 god. nasljednicima ako oboje umru ranije.",
      nasljedstvo_detalj: "Isplata se nakon smrti korisnika nastavlja bračnom drugu doživotno. Ako oboje (korisnik i bračni drug) umru u prvih 10 godina primanja mirovine, preostale mjesečne iznose do kraja 10. godine prima imenovani nasljednik. Nakon isteka razdoblja, nasljednici ne primaju ništa." },
  ].map(v => {
    const nasljedstvo = v.garancija_god ? v.renta * 12 * Math.max(0, v.garancija_god - params.e65) : 0;
    const ukupno_zivot = v.renta * 12 * params.e65;
    const ukupna_vrijednost = ukupno_zivot + nasljedstvo + jednokratni_iznos;
    const pv_renta = v.renta * 12 * AF;
    const pv_nasljedstvo = v.garancija_god ? nasljedstvo / Math.pow(1 + DISCOUNT_R, v.garancija_god) : 0;
    const dcf_ukupna_vrijednost = jednokratni_iznos + pv_renta + pv_nasljedstvo;
    const obiteljska = v.key.startsWith("obiteljska");
    const nasljedstvo_kratko = v.garancija_god
      ? (obiteljska
          ? `Bračnom drugu doživotno; ako i on umre u prvih ${v.garancija_god} god., ostatak prima imenovana osoba`
          : `Ako umrete u prvih ${v.garancija_god} god., ostatak tog razdoblja prima imenovana osoba`)
      : (obiteljska ? "Nastavlja se bračnom drugu doživotno" : "Isplata prestaje");
    return {
      ...v, ukupno_zivot, nasljedstvo, ukupna_vrijednost, nasljedstvo_kratko,
      razlika_kapital: ukupna_vrijednost - kapital,
      dcf_ukupna_vrijednost,
      dcf_razlika_kapital: dcf_ukupna_vrijednost - kapital,
    };
  });

  const varijanta_odabrana = varijante_stvarne.find(v => v.key === params.zomf_varijanta) || varijante_stvarne[0];

  // Život od kamata (5% pravilo) — koliko kapitala treba da se živi samo od prinosa, bez trošenja glavnice
  const POVRAT_5 = 0.05;
  const povrat_5_god = kapital * POVRAT_5;
  const povrat_5_mj = povrat_5_god / 12;
  const tablica_kamate = [200, 300, 400, 500, 600, 800, 1000, 1250, 1500, 2000, 2500, 3000].map(prihod_mj => ({
    prihod_mj,
    prihod_god: prihod_mj * 12,
    potreban_kapital: (prihod_mj * 12) / POVRAT_5,
  }));
  const dostupan_prihod_idx = tablica_kamate.reduce((best, row, i) =>
    row.potreban_kapital <= kapital ? i : best, -1);

  async function handleSave() {
    if (!user) { onAuthRequired("Za spremanje scenarija potrebna je prijava ili registracija. Besplatno je i traje 30 sekundi."); return; }
    const scenario = {
      id: Date.now(), naziv, created: Date.now(),
      stup2: { ...params },
      stup1: { pvb: params.bruto_placa / PROSJ_BRUTO, staz: 40, pf: 1.0, mf: 1.0, avm: AVM },
    };
    const existing = (await stGet(`scenarios:${user.email}`)) || [];
    await stSet(`scenarios:${user.email}`, [...existing, scenario]);
    setSaved(true); setTimeout(() => setSaved(false), 2500);
    if (onSave) onSave();
  }

  async function handleExportPDF() {
    await exportSummaryToPDF({
      title: "II. mirovinski stup — izračun kapitala i rente",
      subtitle: naziv,
      rows: [
        ["Prosječna bruto plaća", `${params.bruto_placa.toLocaleString("hr-HR")} EUR/mj.`],
        ["Godine štednje", `${params.god_stednje} god.`],
        ["Pretpostavljeni neto prinos", `${(params.neto_prinos*100).toFixed(1)}% god.`],
        ["Akumulirani kapital", `${fmt(kapital)} EUR`],
        [`Jednokratna isplata (${Math.round(params.jednokratni_pct*100)}%)`, `${fmt(jednokratni_iznos)} EUR`],
        ["Preostali kapital za rentu", `${fmt(preostali)} EUR`],
        [`Doživotna renta — ${varijanta_odabrana.l}`, `${fmt(varijanta_odabrana.renta)} EUR`],
      ],
      formule: [
        `Kapital (FV) = PMT × [((1+r)^n − 1) / r], PMT = mjesečna bruto plaća × 5%, r = ${(params.neto_prinos*100).toFixed(1)}%/12, n = ${params.god_stednje}×12 = ${fmt(kapital)} EUR`,
        `Jednokratna isplata = Kapital × ${Math.round(params.jednokratni_pct*100)}% = ${fmt(jednokratni_iznos)} EUR`,
        `Doživotna renta ≈ Preostali kapital / (e₆₅ × 12) = ${fmt(preostali)} / (${params.e65} × 12) = ${fmt(renta_standard)} EUR/mj.`,
        varijanta_odabrana.key !== "standard"
          ? `${varijanta_odabrana.l} = ${fmt(varijanta_odabrana.renta)} EUR/mj. — MOD utvrđuje stvarni iznos aktuarskim izračunom.`
          : "MOD koristi vlastite aktuarske tablice — ovo je aproksimacija.",
      ],
    });
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
              { upTo: 1049, color: C.red,    label: "Ispod minimalne plaće (1.050 EUR)" },
              { upTo: 1699, color: C.amber,  label: "Ispod prosjeka" },
              { upTo: 2299, color: C.green,  label: "Oko prosjeka RH (2.016 EUR)" },
              { upTo: 3499, color: C.blue,   label: "Iznad prosjeka" },
              { upTo: 5000, color: C.purple, label: "Visoka plaća" },
            ]}
            dynamicExplain={v => {
              const mj_uplata = v * 0.05;
              return `Od vaše bruto plaće ${v.toLocaleString("hr-HR")} EUR, svaki mjesec ide ${mj_uplata.toFixed(0)} EUR na vaš osobni račun u mirovinskom fondu (5%). Godišnje: ${(mj_uplata*12).toFixed(0)} EUR.`;
            }}
            explain={
              <div>
                <p>Svaki mjesec, od vaše bruto plaće, <strong>5% ide na vaš osobni račun</strong> u Obveznom mirovinskom fondu (OMF). To su vaša osobna sredstva — nitko ih ne može uzeti.</p>
                <p style={{ marginTop: 6 }}>Taj doprinos uplaćuje poslodavac iz vaše bruto plaće (vidljiv je na platnoj listi). Stanje računa možete provjeriti na regos.hr ili u e-Građanima.</p>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.purpleL, borderRadius: 6 }}>
                  💡 <strong>Praktično:</strong> Pri plaći 2.016 EUR bruto = ~101 EUR/mj. → ~1.210 EUR/god. → na 33 godine uz 5,5% prinos = ~112.500 EUR kapitala.
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
              return `Uz ${(v*100).toFixed(1)}% prinos, kamate bi donijele ${fmt(prinos_eur)} EUR dodatno — to je ${Math.round(prinos_eur/k_0*100)}% više od pukih uplata. Naknada OMF (0,23 %/god. u 2026.) već je oduzeta.`;
            }}
            explain={
              <div>
                <p>OMF ulaže vaša sredstva u obveznice, dionice i druge oblike imovine. Prinos <strong>nije zajamčen</strong>. Pravila ulaganja propisuje Zakon o obveznim mirovinskim fondovima, a fondove nadzire Hanfa.</p>
                <ul style={{ marginTop: 8, paddingLeft: 16, display: "flex", flexDirection: "column", gap: 4 }}>
                  <li><strong>3,5%</strong> — konzervativna pretpostavka (loši uvjeti na tržištu)</li>
                  <li><strong>5,5%</strong> — umjerena pretpostavka, blizu dugoročnog prosjeka OMF-ova</li>
                  <li><strong>7,5%</strong> — optimistično (dobra tržišna kretanja)</li>
                </ul>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.amberL, borderRadius: 6 }}>
                  ⚠ Razlika između 3,5% i 7,5% na 33 god. štednje = <strong>oko 99.000 EUR</strong> razlike u kapitalu pri plaći 2.016 EUR.
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
              return `Jednokratno primite ${fmt(Number(jednokr))} EUR odmah. Preostalih ${fmt(Number(preost))} EUR pretvara se u doživotnu mirovinu.`;
            }}
            explain={
              <div>
                <p>Zakonom je dopušteno podići <strong>do 20% akumuliranog kapitala</strong> jednokratno pri umirovljenju — ali samo ako je vaša mirovina iz I. stupa veća od 115% najniže mirovine.</p>
                <p style={{ marginTop: 6 }}>Novac možete iskoristiti kako želite (npr. otplata kredita). Manje podignuto jednokratno znači višu doživotnu mirovinu.</p>
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
                  <li><strong>Muškarci</strong> u dobi 65: prosječno još oko 16 godina (DZS)</li>
                  <li><strong>Žene</strong> u dobi 65: prosječno još oko 20 godina (DZS)</li>
                </ul>
                <div style={{ marginTop: 8, padding: "6px 10px", background: C.blueL, borderRadius: 6 }}>
                  💡 Izračun ne uključuje prinos tijekom isplate, pa je okvirna mirovina nešto niža od stvarne. MOD koristi vlastite aktuarske tablice.
                </div>
              </div>
            }
          />

          {/* ZOMF opcije isplate */}
          <div style={{ borderRadius: 12, border: `2px solid ${C.green}44`, background: C.greenL, overflow: "hidden" }}>
            <div style={{ padding: "12px 14px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                <span style={{ fontSize: 18 }}>🏛</span>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: C.green }}>Vrsta mirovine u MOD-u</div>
                  <div style={{ fontSize: 10, color: C.green, opacity: 0.8 }}>Odaberite da vidite okvirni iznos</div>
                </div>
              </div>
              {[
                { label: "Pojedinačna mirovina", keys: ["standard", "pojedinacna_g5", "pojedinacna_g10"] },
                { label: "Obiteljska (zajednička) mirovina", keys: ["obiteljska", "obiteljska_g5", "obiteljska_g10"] },
              ].map(group => (
                <div key={group.label} style={{ marginBottom: 6 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.green, opacity: 0.7, margin: "4px 0 4px 2px", textTransform: "uppercase" }}>
                    {group.label}
                  </div>
                  {group.keys.map(key => {
                    const v = varijante_stvarne.find(x => x.key === key);
                    return (
                      <button key={v.key} onClick={() => setParams(p => ({ ...p, zomf_varijanta: v.key }))}
                        style={{ width: "100%", marginBottom: 4, padding: "6px 10px",
                          border: `2px solid ${params.zomf_varijanta === v.key ? C.green : C.border}`,
                          borderRadius: 9, background: params.zomf_varijanta === v.key ? C.green : C.white,
                          color: params.zomf_varijanta === v.key ? "#fff" : C.navy,
                          cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>
                        <div style={{ fontSize: 12, fontWeight: 700 }}>{fmt(v.renta)} EUR/mj. — {v.l}</div>
                      </button>
                    );
                  })}
                </div>
              ))}
              <div style={{ fontSize: 10, color: C.green, opacity: 0.8, marginTop: 4 }}>
                {varijanta_odabrana.sub}
              </div>
            </div>
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Akumulirani kapital</h3>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
              <ResultBox label="Ukupni kapital" value={`${fmt(kapital)} EUR`} color={C.purple} big />
              <ResultBox label="Ukupno uplaćeno (bez prinosa)" value={`${fmt(params.bruto_placa * 0.05 * params.god_stednje * 12)} EUR`} color={C.slate} />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <ResultBox label={`Jednokratna isplata (${Math.round(params.jednokratni_pct*100)}%)`}
                value={`${fmt(jednokratni_iznos)} EUR`} color={C.gold} />
              <ResultBox label="Preostali kapital za rentu" value={`${fmt(preostali)} EUR`} color={C.blue} />
            </div>
          </Card>

          <Card style={{ border: `2px solid ${C.green}33` }}>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Doživotna renta — {varijanta_odabrana.l}</h3>
            <ResultBox label="Renta / mj." value={`${fmt(varijanta_odabrana.renta)} EUR`} color={C.green} big />
            <div style={{ marginTop: 10, fontSize: 11, color: C.slateL }}>
              {varijanta_odabrana.key === "standard"
                ? "Okvirni izračun: kapital podijeljen s očekivanim brojem mjeseci isplate. Stvarni iznos izračunava MOD."
                : `${varijanta_odabrana.sub} Stvarni iznos utvrđuje MOD aktuarskim izračunom prema dobi, spolu i (za obiteljske varijante) dobi bračnog druga.`}
            </div>
            <div style={{ marginTop: 8, fontSize: 11, color: C.green, background: C.greenL, padding: "8px 12px", borderRadius: 8 }}>
              <strong>Nakon smrti korisnika:</strong> {varijanta_odabrana.nasljedstvo_detalj}
            </div>
          </Card>

          <Card>
            <h3 style={{ fontSize: 13, marginBottom: 12 }}>Usporedba vrsta mirovine (okvirno)</h3>
            <ResponsiveContainer width="100%" height={190}>
              <BarChart data={varijante_stvarne.map(v => ({ name: v.short, renta: v.renta, key: v.key }))} margin={{ top: 5, right: 10, bottom: 50, left: 0 }}>
                <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-30} textAnchor="end" interval={0} />
                <YAxis tick={{ fontSize: 10 }} tickFormatter={v => `${Math.round(v)}`} />
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <Tooltip formatter={v => [`${fmt(v)} EUR/mj.`, "Renta"]} />
                <Bar dataKey="renta" radius={[4,4,0,0]}>
                  {varijante_stvarne.map((d, i) => <Cell key={i} fill={d.key === params.zomf_varijanta ? C.green : C.blue} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
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
              <button onClick={handleExportPDF} title="Izvezi izračun kao PDF" style={{
                padding: "8px 14px", background: C.white, color: C.navy,
                border: `1.5px solid ${C.border}`, borderRadius: 8, fontWeight: 600, fontSize: 13,
                display: "flex", alignItems: "center", gap: 6,
              }}>
                <FileText size={15} />
                PDF
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

      {/* Kako odabrati vrstu mirovine u MOD-u */}
      <Card style={{ marginTop: 20 }}>
        <h3 style={{ fontSize: 15, marginBottom: 6 }}>Koju vrstu mirovine odabrati u MOD-u?</h3>
        <p style={{ fontSize: 12, color: C.slateL, marginBottom: 12, lineHeight: 1.6 }}>
          Pri umirovljenju kapital ({fmt(preostali)} EUR nakon jednokratne isplate) prelazi u mirovinsko osiguravajuće društvo (MOD),
          koje vam doživotno isplaćuje mirovinu. U Hrvatskoj trenutno posluju dva MOD-a (Hrvatski MOD i Raiffeisen MOD). Odluka je trajna.
        </p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 10, marginBottom: 14 }}>
          {[
            { t: "Pojedinačna", c: C.blue, d: "Najviša mjesečna isplata. Isplata prestaje vašom smrću.", kada: "Kad nemate bračnog druga koji ovisi o vašoj mirovini." },
            { t: "Obiteljska (zajednička)", c: C.green, d: "Niža isplata, ali se nakon vaše smrti nastavlja bračnom drugu doživotno.", kada: "Kad bi bračni drug bez vaše mirovine teško živio." },
            { t: "Sa zajamčenim razdobljem", c: C.gold, d: "Ako umrete u prvih 5 ili 10 godina, ostatak tog razdoblja prima osoba koju odredite.", kada: "Kad želite zaštitu obitelji u slučaju rane smrti." },
          ].map(o => (
            <div key={o.t} style={{ padding: "12px 14px", borderRadius: 10, background: o.c + "10", border: `1.5px solid ${o.c}33` }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: o.c, marginBottom: 4 }}>{o.t}</div>
              <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.55, marginBottom: 6 }}>{o.d}</div>
              <div style={{ fontSize: 11, color: C.navy }}><strong>Kada:</strong> {o.kada}</div>
            </div>
          ))}
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 520 }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${C.border}` }}>
                <th style={{ textAlign: "left", padding: "8px 10px" }}>Vrsta mirovine</th>
                <th style={{ textAlign: "right", padding: "8px 10px" }}>Okvirno / mj.</th>
                <th style={{ textAlign: "left", padding: "8px 10px" }}>Nakon vaše smrti</th>
              </tr>
            </thead>
            <tbody>
              {varijante_stvarne.map(v => (
                <tr key={v.key} style={{ borderBottom: `1px solid ${C.border}`, background: v.key === params.zomf_varijanta ? C.greenL : "transparent" }}>
                  <td style={{ padding: "8px 10px", fontWeight: 600 }}>{v.short}</td>
                  <td style={{ padding: "8px 10px", textAlign: "right", fontWeight: 700 }}>{fmt(v.renta)} EUR</td>
                  <td style={{ padding: "8px 10px", color: C.slate }}>{v.nasljedstvo_kratko}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ fontSize: 10, color: C.slateL, marginTop: 8, lineHeight: 1.5 }}>
          Iznosi su okvirni: obiteljska i zajamčena varijanta procijenjene su kao 1,5–12 % niže od pojedinačne. Stvarni iznos
          izračunava MOD prema vašoj dobi, spolu i dobi bračnog druga. Usporedite ponude oba MOD-a prije potpisa ugovora.
        </div>
        <div style={{ marginTop: 12, padding: "12px 16px", background: C.blueL, borderRadius: 10, border: `1.5px solid ${C.blue}33` }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.blue, marginBottom: 6 }}>Važna pravila</div>
          <ul style={{ fontSize: 12, color: C.slate, lineHeight: 1.7, paddingLeft: 18, margin: 0 }}>
            <li>Mirovina iz oba stupa moguća je samo ako je osnovna mirovina iz I. stupa veća od najniže mirovine.
              Inače se mirovina određuje samo iz I. stupa.</li>
            <li>Jednokratno možete podići do 20 % kapitala samo ako je osnovna mirovina iz I. stupa veća od 115 % najniže mirovine.</li>
            <li>Umjesto mirovine iz oba stupa možete izabrati mirovinu samo iz I. stupa (sredstva II. stupa tada idu u državni proračun).
              Usporedbu radi kalkulator „Samo I. ili I. + II.“.</li>
          </ul>
        </div>
      </Card>

      {/* Nasljeđivanje sredstava II. stupa */}
      <Card style={{ marginTop: 20 }}>
        <h3 style={{ fontSize: 15, marginBottom: 8 }}>Što nasljeđuje obitelj?</h3>
        <ul style={{ fontSize: 12, color: C.slate, lineHeight: 1.75, paddingLeft: 18, margin: 0 }}>
          <li><strong>Prije umirovljenja:</strong> sredstva na vašem osobnom računu u II. stupu nasljeđuju se kao i druga imovina.
            Ako član obitelji ostvaruje obiteljsku mirovinu, može izabrati nasljeđivanje sredstava (tada mu se dio obiteljske
            mirovine za staž od 2002. određuje kao osnovna mirovina) ili prijenos sredstava u I. stup i punu obiteljsku mirovinu.</li>
          <li><strong>Nakon umirovljenja:</strong> sredstva se više ne nasljeđuju. Isplata se nastavlja samo bračnom drugu
            (obiteljska varijanta) ili do kraja zajamčenog razdoblja (varijante s jamstvom).</li>
          <li><strong>Jednokratna isplata</strong> (do 20 %) postaje vaš novac i nasljeđuje se normalno.</li>
        </ul>
      </Card>

      <Detalji naslov="Za napredne: život samo od prinosa (ilustracija)" style={{ marginTop: 20 }}>
        <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.6, marginBottom: 10 }}>
          II. stup se u Hrvatskoj ne može ostaviti uloženim i trošiti samo prinos: kapital se obavezno pretvara u doživotnu mirovinu.
          Tablica ispod samo ilustrira koliko bi kapitala trebalo za određeni mjesečni prihod uz 5 % godišnjeg prinosa
          (npr. za vlastitu štednju ili III. stup).
        </p>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
          <ResultBox label="5% godišnje (ne dira glavnicu)" value={`${fmt(povrat_5_god)} EUR/god.`} color={C.gold} big />
          <ResultBox label="= mjesečno" value={`${fmt(povrat_5_mj)} EUR/mj.`} color={C.gold} big />
        </div>
        <p style={{ fontSize: 12, color: C.slateL, marginBottom: 10, lineHeight: 1.6 }}>
          Obrnuto — koliko bi kapitala trebalo da uz 5% godišnji prinos isplata pokrije željeni mjesečni iznos,
          bez trošenja glavnice:
        </p>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 480 }}>
            <thead>
              <tr style={{ borderBottom: `2px solid ${C.border}` }}>
                <th style={{ textAlign: "left", padding: "8px 10px" }}>Željeni prihod / mj.</th>
                <th style={{ textAlign: "right", padding: "8px 10px" }}>Željeni prihod / god.</th>
                <th style={{ textAlign: "right", padding: "8px 10px" }}>Potreban kapital (uz 5% god.)</th>
              </tr>
            </thead>
            <tbody>
              {tablica_kamate.map((row, i) => (
                <tr key={row.prihod_mj} style={{ borderBottom: `1px solid ${C.border}`,
                  background: i === dostupan_prihod_idx ? C.goldL : "transparent" }}>
                  <td style={{ padding: "8px 10px", fontWeight: 600 }}>
                    {i === dostupan_prihod_idx ? "🏆 " : ""}{fmt(row.prihod_mj)} EUR
                  </td>
                  <td style={{ padding: "8px 10px", textAlign: "right" }}>{fmt(row.prihod_god)} EUR</td>
                  <td style={{ padding: "8px 10px", textAlign: "right", fontWeight: 700 }}>{fmt(row.potreban_kapital)} EUR</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ fontSize: 10, color: C.slateL, marginTop: 8, lineHeight: 1.5 }}>
          {dostupan_prihod_idx >= 0
            ? `🏆 = najviši mjesečni iznos koji vaš trenutni kapital (${fmt(kapital)} EUR) već pokriva uz 5% god. prinos, bez trošenja glavnice.`
            : `Vaš trenutni kapital (${fmt(kapital)} EUR) uz 5% god. prinos ne pokriva ni najniži prag od ${fmt(tablica_kamate[0].prihod_mj)} EUR/mj. iz tablice.`}
          {" "}Pravilo "5%" je pojednostavljena procjena (ne uzima u obzir inflaciju, poreze ni stvarnu volatilnost prinosa) — služi za ilustraciju, ne kao financijski savjet.
        </div>
      </Detalji>
    </div>
  );
}

// ─── UKUPNO ──────────────────────────────────────────────────────────────────
function UkupnoView() {
  const [p, setP] = useState({ pvb: 1.0, staz: 40, stazOd2002: 25, kapital: 60000, jednokratni_pct: 0, e65: 18 });
  const set = (k) => (v) => setP(prev => ({ ...prev, [k]: v }));

  const renta = rentaDozivotna(p.kapital * (1 - p.jednokratni_pct), p.e65);
  const iz = calcIzbor({ pvb: p.pvb, staz: p.staz, stazOd2002: p.stazOd2002, rentaII: renta });
  const najniza = iz.samoI.najniza;
  const kombiniranaMoguca = iz.osnovnaSDodatkom > najniza;
  const jednokratnaMoguca = iz.osnovnaSDodatkom > najniza * 1.15;
  const boljeKombinirana = kombiniranaMoguca && iz.razlika > 0;

  async function handleExportPDF() {
    await exportSummaryToPDF({
      title: "Samo I. stup ili I. + II. stup — usporedba",
      subtitle: `PVB ${p.pvb.toFixed(2)}, staž ${p.staz} god. (od 2002.: ${p.stazOd2002} god.)`,
      rows: [
        ["Samo I. stup (s dodatkom 27 %)", `${fmt(iz.samoIUkupno)} EUR/mj.`],
        ["Osnovna mirovina iz I. stupa (s dodatkom 27 %)", `${fmt(iz.osnovnaSDodatkom)} EUR/mj.`],
        ["Mirovina iz II. stupa (okvirno)", `${fmt(renta)} EUR/mj.`],
        ["I. + II. stup ukupno", `${fmt(iz.kombinirana)} EUR/mj.`],
        ["Razlika (I. + II. minus samo I.)", `${fmt(iz.razlika)} EUR/mj.`],
        ["Kapital u II. stupu", `${fmt(p.kapital)} EUR`],
      ],
      formule: [
        `Samo I. stup = PVB × staž × AVM × 1,27 = ${fmt(iz.samoIUkupno)} EUR`,
        `Osnovna mirovina = PVB × (staž do 2002. + staž od 2002. × ${FAKTOR_OSNOVNE}) × AVM × 1,27 = ${fmt(iz.osnovnaSDodatkom)} EUR`,
        `Mirovina iz II. stupa ≈ kapital × (1 − jednokratno) / (${p.e65} × 12) = ${fmt(renta)} EUR`,
      ],
    });
  }

  return (
    <div className="fadeIn">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4, gap: 12 }}>
        <h2 style={{ fontSize: 22 }}>Samo I. stup ili I. + II. stup?</h2>
        <button onClick={handleExportPDF} title="Izvezi izračun kao PDF" style={{
          padding: "8px 14px", background: C.white, color: C.navy,
          border: `1.5px solid ${C.border}`, borderRadius: 8, fontWeight: 600, fontSize: 13,
          display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap",
        }}>
          <FileText size={15} />
          PDF
        </button>
      </div>
      <p style={{ color: C.slateL, fontSize: 13, marginBottom: 16 }}>Najvažnija odluka za sve koji su u II. stupu, a bliže se mirovini</p>

      <Ukratko color={C.purple}>
        Kad odlazite u mirovinu, birate između dvije mogućnosti. <strong>Samo I. stup:</strong> mirovina se računa kao da
        nikad niste bili u II. stupu, uz dodatak od 27 %, a novac iz II. stupa prelazi u državni proračun.{" "}
        <strong>I. + II. stup:</strong> dobivate nešto nižu „osnovnu“ mirovinu iz I. stupa (također s dodatkom od 27 %) i
        uz nju doživotnu mirovinu iz II. stupa. Isplativije je ono što daje veći mjesečni iznos. HZMO vam prije odluke
        šalje informativni izračun obje mogućnosti.
      </Ukratko>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 20, marginBottom: 20 }}>
        <Card>
          <h3 style={{ fontSize: 14, marginBottom: 14, color: C.navy }}>Vaši podaci</h3>
          <Slider label="Prosječna plaća u odnosu na prosjek RH (PVB)" value={p.pvb} min={0.3} max={3.8} step={0.01}
            format={v => `${v.toFixed(2)} (≈ ${fmt(v * PROSJ_BRUTO, 0)} EUR danas)`} onChange={set("pvb")} />
          <Slider label="Ukupni mirovinski staž (god.)" value={p.staz} min={15} max={50} step={1}
            format={v => `${v}`} onChange={v => setP(prev => ({ ...prev, staz: v, stazOd2002: Math.min(prev.stazOd2002, v) }))} />
          <Slider label="Od toga staž od 2002. (godine u II. stupu)" value={p.stazOd2002} min={0} max={Math.min(p.staz, 45)} step={1}
            format={v => `${v}`} onChange={set("stazOd2002")} />
        </Card>
        <Card>
          <h3 style={{ fontSize: 14, marginBottom: 14, color: C.purple }}>II. stup</h3>
          <Slider label="Stanje računa u II. stupu (EUR, vidi regos.hr ili e-Građani)" value={p.kapital} min={0} max={300000} step={1000}
            format={v => `${fmt(v, 0)} EUR`} onChange={set("kapital")} />
          <Slider label="Jednokratna isplata" value={p.jednokratni_pct} min={0} max={0.2} step={0.05}
            format={v => v === 0 ? "ne uzimam" : `${Math.round(v * 100)} %`} onChange={set("jednokratni_pct")} />
          <Slider label="Očekivane godine primanja mirovine" value={p.e65} min={12} max={26} step={1}
            format={v => `${v}`} onChange={set("e65")} />
        </Card>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16, marginBottom: 20 }}>
        <Card style={{ border: `2px solid ${!boljeKombinirana ? C.green : C.border}` }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, marginBottom: 8 }}>A) Samo I. stup</div>
          <div style={{ fontSize: 28, fontWeight: 800, color: C.navy, fontFamily: "'Playfair Display', serif" }}>{fmt(iz.samoIUkupno)} EUR</div>
          <div style={{ fontSize: 11, color: C.slateL, marginTop: 6, lineHeight: 1.6 }}>
            {iz.dI.najniza ? <>Isplaćuje se najniža mirovina ({fmt(iz.dI.ukupno)} EUR), jer je viša od izračuna s dodatkom.</> : <>{fmt(iz.dI.mirovina)} EUR + dodatak 27 % ({fmt(iz.dI.dodatak)} EUR).</>}
            Kapital iz II. stupa ({fmt(p.kapital, 0)} EUR) prelazi u državni proračun.
          </div>
          {!boljeKombinirana && <div style={{ marginTop: 8, fontSize: 12, fontWeight: 700, color: C.green }}>✓ Viša mjesečna mirovina</div>}
        </Card>
        <Card style={{ border: `2px solid ${boljeKombinirana ? C.green : C.border}` }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: C.purple, marginBottom: 8 }}>B) I. + II. stup</div>
          <div style={{ fontSize: 28, fontWeight: 800, color: C.purple, fontFamily: "'Playfair Display', serif" }}>{fmt(iz.kombinirana)} EUR</div>
          <div style={{ fontSize: 11, color: C.slateL, marginTop: 6, lineHeight: 1.6 }}>
            Osnovna mirovina {fmt(iz.osnovnaSDodatkom)} EUR ({iz.dO.najniza ? "najniža osnovna mirovina" : "s dodatkom 27 %"}) + mirovina iz II. stupa okvirno {fmt(renta)} EUR.
            {p.jednokratni_pct > 0 && <> Uz to jednokratno {fmt(p.kapital * p.jednokratni_pct)} EUR.</>}
          </div>
          {boljeKombinirana && <div style={{ marginTop: 8, fontSize: 12, fontWeight: 700, color: C.green }}>✓ Viša mjesečna mirovina</div>}
          {!kombiniranaMoguca && <div style={{ marginTop: 8, fontSize: 12, fontWeight: 700, color: C.red }}>Nije moguće uz ove podatke</div>}
        </Card>
      </div>

      <Card style={{ marginBottom: 20, background: boljeKombinirana ? C.greenL : C.blueL }}>
        <div style={{ fontSize: 13, color: C.navy, lineHeight: 1.7 }}>
          <strong>Što to znači:</strong>{" "}
          {!kombiniranaMoguca
            ? <>Vaša osnovna mirovina s dodatkom ({fmt(iz.osnovnaSDodatkom)} EUR) nije veća od najniže mirovine za vaš staž ({fmt(najniza)} EUR), pa mirovina iz oba stupa nije moguća. Mirovina se određuje samo iz I. stupa.</>
            : boljeKombinirana
              ? <>Uz ove pretpostavke mirovina iz oba stupa daje oko <strong>{fmt(iz.razlika)} EUR više</strong> mjesečno. Pri odabiru vrste mirovine u MOD-u (npr. obiteljske) iznos iz II. stupa bit će nešto niži.</>
              : <>Uz ove pretpostavke mirovina samo iz I. stupa daje oko <strong>{fmt(-iz.razlika)} EUR više</strong> mjesečno. To je čest ishod kad je kapital u II. stupu manji ili kad ste mnogo godina staža ostvarili prije 2002.</>}
          {kombiniranaMoguca && !jednokratnaMoguca && p.jednokratni_pct > 0 && <> Jednokratna isplata nije moguća jer osnovna mirovina nije veća od 115 % najniže ({fmt(najniza * 1.15)} EUR).</>}
        </div>
        <div style={{ fontSize: 11, color: C.slate, marginTop: 8 }}>
          Ovo je okvirna procjena. Konačnu odluku donesite na temelju informativnog izračuna HZMO-a i ponude MOD-a.
        </div>
      </Card>

      <NetoMirovinaBox bruto={Math.max(iz.samoIUkupno, iz.kombinirana)} />

      <Detalji naslov="Kako se računa osnovna mirovina i dodatak od 27 %?">
        <ul style={{ fontSize: 12, color: C.slate, lineHeight: 1.75, paddingLeft: 18, margin: 0 }}>
          <li>Od 2002. osiguranici u oba stupa plaćaju 15 % plaće u I. stup i 5 % u II. stup. Zato se dio mirovine iz I. stupa
            za staž od 2002. množi faktorom osnovne mirovine od oko {FAKTOR_OSNOVNE.toString().replace(".", ",")} (15 od 20 postotnih bodova, ZOMO čl. 90.).
            Staž prije 2002. računa se u punom iznosu.</li>
          <li>Dodatak od 27 % pripada i mirovini samo iz I. stupa i osnovnoj mirovini kod mirovine iz oba stupa. Od 1.1.2024.
            kod kombinirane mirovine računa se na cijelu osnovnu mirovinu (NN 156/23). HZMO ga obračunava sam, zahtjev nije potreban.</li>
          <li>Od 1.1.2019. izbor između ove dvije mogućnosti ima svatko tko je u oba stupa, u postupku ostvarivanja starosne ili
            prijevremene starosne mirovine. Izjava se daje u Mirovinskim informativnim centrima, u FINA-i ili putem e-usluge „Odabir mirovine“.</li>
          <li>Mirovina iz II. stupa ovdje je okvirna (kapital podijeljen s brojem mjeseci isplate, bez prinosa tijekom isplate).</li>
        </ul>
      </Detalji>
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
                      <span style={{ fontSize: 10, color: "rgba(255,255,255,0.6)" }}>{k2 !== null ? "Ukupno (osnovna I. + 27 % + II.)" : "Ukupno (I. + 27 %)"}</span>
                      <span style={{ fontSize: 15, fontWeight: 800, color: "#7bffb0" }}>
                        {fmtEur(r1
                          ? kombiniranaMirovina({ ...s.stup1, stazOd2002: k2 !== null ? Math.min(s.stup1.staz, s.stup2.god_stednje) : 0, rentaII: renta2 || 0 }).ukupno
                          : (renta2 || 0))}/mj.
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
  const [placa, setPlaca] = useState(PROSJ_BRUTO);
  const [staz, setStaz] = useState(40);
  const [doprinos_pct, setDoprinos] = useState(0.20);
  const [prinos_mode, setPrinosMode] = useState("40y_nominal"); // 40y_nominal | 40y_realni | long_nominal | custom
  const [custom_prinos, setCustomPrinos] = useState(0.07);
  const [showDetail, setShowDetail] = useState(false);

  const prinos_map = {
    "40y_nominal": SP500_40Y_NOMINAL,
    "40y_realni":  SP500_40Y_REALNI,
    "long_nominal":SP500_LONG_NOMINAL,
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
  const pvb_est = placa / PROSJ_BRUTO;
  const hr_ii = rentaDozivotna(calcStup2FV({ bruto_placa: placa, god_stednje: staz, neto_prinos: 0.055 }) * 0.8);
  const hr_k = kombiniranaMirovina({ pvb: pvb_est, staz, rentaII: hr_ii });
  const hr_i = hr_k.osnovna;
  const hr_dod = hr_k.dodatak;
  const hr_uk = hr_k.ukupno;

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
  const [filter, setFilter] = useState("Sve");

  const cats = ["Sve", ...new Set(BLOG_POSTS.map(p => p.category))];
  const filtered = filter === "Sve" ? BLOG_POSTS : BLOG_POSTS.filter(p => p.category === filter);

  function renderMd(text) {
    return text.split("\n").map((line, i, arr) => {
      if (line.startsWith("## ")) return <h2 key={i} style={{ fontSize: 18, marginTop: 24, marginBottom: 8, color: C.navy }}>{line.slice(3)}</h2>;
      if (line.startsWith("### ")) return <h3 key={i} style={{ fontSize: 15, marginTop: 16, marginBottom: 6, color: C.navyL }}>{line.slice(4)}</h3>;
      if (line.startsWith("> ")) return <blockquote key={i} style={{ borderLeft: `3px solid ${C.gold}`, paddingLeft: 12, color: C.slate, fontStyle: "italic", margin: "12px 0" }}>{line.slice(2)}</blockquote>;
      if (line.startsWith("|")) {
        if (/^\|[\s\-|]+\|$/.test(line.trim())) return null;
        const cells = line.trim().replace(/^\||\|$/g, "").split("|").map(c => c.trim());
        const header = i + 1 < arr.length && /^\|[\s\-|]+\|$/.test(arr[i + 1].trim());
        return (
          <div key={i} style={{ display: "grid", gridTemplateColumns: `repeat(${cells.length}, 1fr)`, gap: 8,
            padding: "6px 10px", fontSize: 13, background: header ? C.navy : (i % 2 ? C.bg : C.white),
            color: header ? "#fff" : C.slate, fontWeight: header ? 700 : 400, borderRadius: header ? "8px 8px 0 0" : 0 }}>
            {cells.map((c, j) => <span key={j}>{c}</span>)}
          </div>
        );
      }
      if (line.startsWith("- ")) return <p key={i} style={{ lineHeight: 1.7, marginBottom: 4, paddingLeft: 14 }} dangerouslySetInnerHTML={{ __html: "• " + line.slice(2).replace(/\*\*(.*?)\*\*/g, (_, t) => `<strong>${t}</strong>`) }} />;
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

    </div>
  );
}

// ─── EDUKACIJA ───────────────────────────────────────────────────────────────
const AVG_HR = PROSJ_BRUTO;

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
  const [plaça, setPlaça] = useState(PROSJ_BRUTO);
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
    <div className="fadeIn">

      {/* Horizontal nav, centered */}
      <div style={{ position: "sticky", top: 80, zIndex: 5, display: "flex", justifyContent: "center",
        flexWrap: "wrap", gap: 6, background: C.white, borderRadius: 14,
        border: `1px solid ${C.border}`, padding: "10px 12px", marginBottom: 24,
        boxShadow: "0 2px 12px rgba(26,46,82,0.07)" }}>
        {navItems.map(n => (
          <button key={n.id} onClick={() => scrollTo(n.id)} style={{
            padding: "8px 14px", border: "none", borderRadius: 8,
            background: section === n.id ? C.goldL : "none",
            textAlign: "center", fontSize: 12, fontWeight: section === n.id ? 700 : 400,
            color: section === n.id ? C.navy : C.slate,
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
          <h1 style={{ fontSize: 28, marginBottom: 14, letterSpacing: "-0.02em", lineHeight: 1.2 }}>
            Vaša mirovina se gradi danas — ne kad odete u mirovinu.
          </h1>
          <p style={{ fontSize: 17, fontWeight: 700, color: "rgba(255,255,255,0.95)", lineHeight: 1.6, maxWidth: 640, marginBottom: 14 }}>
            Većina ljudi ne razmišlja o mirovini dok je ne počne primati. To je kasno
            — previše kasno da se išta promijeni.
          </p>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, maxWidth: 640 }}>
            <p style={{ fontSize: 14, color: "rgba(255,255,255,0.75)", lineHeight: 1.7 }}>
              Svaka godina rada ostavlja trag u vašem mirovinskom dosjeu. Svaki postotak
              plaće, svaki valorizirani bod, svaki mjesec staža — sve se zbraja ili
              oduzima od iznosa koji ćete primati do kraja života.
            </p>
            <p style={{ fontSize: 14, color: "rgba(255,255,255,0.75)", lineHeight: 1.7 }}>
              Razlika između osobe koja razumije sustav i one koja ga ne razumije
              može biti i nekoliko stotina eura mirovine svaki mjesec — zauvijek.
            </p>
            <p style={{ fontSize: 14, color: "rgba(255,255,255,0.75)", lineHeight: 1.7 }}>
              Ovdje ćete naučiti što je PVB i zašto pregovaranje o plaći danas izravno
              povećava vašu mirovinu za 30 godina. Zašto jedan jedini mjesec ranijeg
              odlaska u mirovinu trajno smanjuje svaki vaš budući ček. Zašto 55 eura
              mjesečno u trećem stupu može biti vrijednije od iznosa koji u njega ulažete.
            </p>
            <p style={{ fontSize: 14, color: "rgba(255,255,255,0.75)", lineHeight: 1.7 }}>
              Nijedan od ovih podataka nije skriven — ali nitko vam ih nije objasnio
              jasno, s konkretnim brojevima i vašim primjerom. Do sad.
            </p>
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
                { sym: "MF*", desc: "Mirovinski faktor (vrsta)", color: "#ff9090" },
                { sym: "×", desc: "", color: "#fff" },
                { sym: "AVM", desc: `Aktualna vrijednost = ${fmt(AVM)} EUR`, color: C.goldL },
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
            <div style={{ marginTop: 10, fontSize: 11, color: "rgba(255,255,255,0.5)", lineHeight: 1.6 }}>
              * MF (mirovinski faktor) ovisi o vrsti mirovine: <strong style={{ color: "#fff" }}>1,00</strong> za
              starosnu i prijevremenu mirovinu, <strong style={{ color: "#fff" }}>1,10</strong> za potpunu
              invalidsku i <strong style={{ color: "#fff" }}>0,90</strong> za djelomičnu invalidsku mirovinu
              (od 1.1.2026., ZOMO čl. 48.).
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
                    <li>Starosna mirovina: <strong>min. 15 god.</strong> staža + dob 65 (muškarci) / 64 u 2026. (žene, do 2030. raste na 65)</li>
                    <li>Prijevremena: muškarci <strong>60 god. + 35 god.</strong> staža, žene 59 god. + 34 god. staža (2026.)</li>
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
                    { plaça: PROSJ_BRUTO, pvb: "1,00", boja: C.green,  label: "Prosjek RH" },
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
                  { pvb_val: 0.52, opis: "Minimalna plaća cijele karijere", plaça_eur: "~1.050 EUR bruto (2026.)" },
                  { pvb_val: 1.0, opis: "Uvijek prosječna plaća RH", plaça_eur: "~2.016 EUR bruto (2025.)" },
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
              i mirovina je niža. Odlazite li <span style={{ color: C.green, fontWeight: 700 }}>kasnije</span>{" "}
              — PF je veći od 1,0 i mirovina je trajno viša. Smanjenje zbog ranijeg odlaska ne prestaje kad navršite
              redovnu dob, nego tek s <strong>70 godina</strong> (od 1.1.2026.), kada ga HZMO sam ukida.
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
                  <div style={{ marginTop: 4, color: C.red }}>Smanjenje do 70. godine: <strong>−{((1-pf_rani)*100).toFixed(1)}%</strong></div>
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
                  { god: "2024.", avm: 13.57 }, { god: "1.7.2025.", avm: 14.45 }, { god: "1.7.2026.", avm: 15.64, curr: true },
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
                <strong>Kako se usklađuje (ZOMO čl. 91.):</strong>
                <ul style={{ paddingLeft: 16, marginTop: 6 }}>
                  <li>2× godišnje (1. siječnja i 1. srpnja)</li>
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
                const k2 = calcStup2FV({ bruto_placa: plaça, god_stednje: staz, neto_prinos: prinos });
                const r2 = rentaDozivotna(k2 * 0.8);
                const km = kombiniranaMirovina({ pvb: plaça / AVG_HR, staz, rentaII: r2 });
                const m1 = km.osnovna;
                const dod = km.dodatak;
                const uk = km.ukupno;
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
                { t: "Demografska vremenska bomba", d: "Omjer radnik/umirovljenik pada. 2050. bit će manje od 1 radnika po umirovljeniku. I. stup tada neće biti održiv bez drastičnog smanjenja mirovina ili povećanja poreza." },
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
                  obveze: ["Vodi vaš osobni račun", "Rasporedit će vas u fond kategorije B ako ne odaberete", "Prebacuje sredstva MOD-u pri umirovljenju"] },
                { emoji: "📈", naziv: "Mirovinsko društvo (OMF)", boja: C.purple,
                  opis: "Profesionalni upravitelj fonda. Ulaže vaša sredstva u financijske instrumente. Pod nadzorom HANFA-e.",
                  obveze: ["Ulaže novac prema strategiji fonda", "Naplaćuje naknadu za upravljanje (0,23 % god. u 2026.)", "Pod nadzorom je Hanfe"] },
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
                U Hrvatskoj trenutno posluju <strong>dva MOD-a</strong> (registar Hanfe, 2026.):
              </p>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
                {[
                  { n: "Hrvatsko mirovinsko osiguravajuće društvo d.d.", v: "Zagreb" },
                  { n: "Raiffeisen mirovinsko osiguravajuće društvo d.d.", v: "Zagreb" },
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
                  bez obzira koliko dugo živite. Mjesečni iznos zato ovisi o očekivanom trajanju života i kamatama.
                </div>
              </div>
            </Card>

            <Card style={{ borderLeft: `4px solid ${C.purple}` }}>
              <h3 style={{ fontSize: 14, color: C.purple, marginBottom: 14 }}>📈 Kategorije OMF fondova (A, B, C)</h3>
              <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.7, marginBottom: 12 }}>
                Svako mirovinsko društvo upravlja s <strong>tri kategorije fonda</strong> različitog rizika.
                Ako sami ne izaberete, REGOS vas raspoređuje u kategoriju B.
              </p>
              {[
                { kat: "A", boja: C.green, profil: "Agresivni rast",
                  dionice: "najviše dionica", za: "Više od 10 god. do mirovine",
                  opis: "Najviši očekivani prinos dugoročno, ali i najveće oscilacije. Kad do mirovine ostane manje od 10 godina, REGOS vas prebacuje u B (na zahtjev možete ostati u A do 5 godina prije mirovine)." },
                { kat: "B", boja: C.blue, profil: "Umjereni rast",
                  dionice: "umjereno dionica", za: "Zadana kategorija",
                  opis: "Ravnoteža rasta i sigurnosti. U nju se raspoređuju svi koji ne izaberu kategoriju." },
                { kat: "C", boja: C.slate, profil: "Konzervativni",
                  dionice: "malo ili bez dionica", za: "Zadnjih 5 god. do mirovine",
                  opis: "Naglasak na zaštiti kapitala: niži prinos, niži rizik gubitka. Namijenjen je osiguranicima kojima je do mirovine manje od 5 godina." },
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
                Fond i kategoriju možete promijeniti preko REGOS-a (FINA ili e-Građani). Točna pravila i rokove provjerite na regos.hr.
              </div>
            </Card>
          </div>

          {/* Oporezivanje mirovina */}
          <Card style={{ marginBottom: 18, border: `2px solid ${C.gold}33` }}>
            <h3 style={{ fontSize: 14, color: C.navy, marginBottom: 14 }}>💶 Oporezivanje mirovina — što ostaje na računu</h3>
            <div style={{ padding: "10px 14px", background: C.amberL, borderRadius: 10, border: `1.5px solid ${C.amber}55`, marginBottom: 14 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.amber, marginBottom: 4 }}>⏳ Od 1.1.2027. porez na mirovine trebao bi se ukinuti</div>
              <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.6 }}>
                Vlada je u rujnu 2026. uputila Saboru izmjene Zakona o porezu na dohodak: umanjenje poreza za umirovljenike raste s 50 %
                na 100 % za dio mirovine oporeziv nižom stopom (do 5.000 EUR osnovice). Ako ih Sabor usvoji, većina umirovljenika od 2027. neće plaćati porez.
                Do tada vrijede pravila u nastavku.
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
              <div style={{ padding: "12px 14px", background: C.bg, borderRadius: 10 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.navy, marginBottom: 8 }}>Kako se obračunava porez</div>
                <ol style={{ fontSize: 12, color: C.slate, lineHeight: 1.8, paddingLeft: 16, margin: 0 }}>
                  <li>Porezna osnovica = mirovina − <strong>{OSOBNI_ODBITAK} EUR</strong> (osobni odbitak/mj.)</li>
                  <li>Porez po stopi vašeg grada ili općine: niža stopa 15–23 % (najčešće 20 %) do 5.000 EUR osnovice, viša 25–33 % iznad</li>
                  <li>Umirovljenici plaćaju <strong>pola</strong> obračunatog poreza (umanjenje 50 %)</li>
                  <li>Prirez ne postoji od 2024. (uključen je u stope gradova)</li>
                </ol>
              </div>
              <div style={{ padding: "12px 14px", background: C.bg, borderRadius: 10 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.navy, marginBottom: 8 }}>Primjeri neto mirovine</div>
                {[600, 900, 1200, 1800].map(b => {
                  const t = calcNetoMirovina(b);
                  return (
                    <div key={b} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, borderBottom: `1px solid ${C.border}`, padding: "4px 0" }}>
                      <span style={{ color: C.slateL }}>{b} EUR bruto</span>
                      <span style={{ fontWeight: 700, color: C.green }}>≈ {fmt(t.neto)} EUR neto</span>
                    </div>
                  );
                })}
                <div style={{ fontSize: 10, color: C.slateL, marginTop: 6 }}>Stopa 20 % · osobni odbitak {OSOBNI_ODBITAK} EUR · 2026.</div>
              </div>
            </div>
            <div style={{ fontSize: 11, color: C.slateL, lineHeight: 1.6 }}>
              Mirovine ≤ {OSOBNI_ODBITAK} EUR/mj. ne plaćaju porez. Osobni odbitak može biti veći uz uzdržavane članove ili invaliditet.
              Točan obračun radi HZMO i isplaćuje neto iznos. Godišnji obračun (porezna prijava) nije obveza za umirovljenike s jednim izvorom dohotka.
            </div>
          </Card>

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
                  opis: "Birate mirovinu SAMO iz I. stupa (računa se kao da nikad niste bili u II. stupu, uz dodatak 27 %; sredstva II. stupa idu u državni proračun) ili mirovinu iz I. + II. stupa.",
                  detalj: "Od 2019. ovaj izbor ima svatko tko je u oba stupa. Prije odluke dobivate informativni izračun obje mogućnosti." },
                { n: "3", boja: C.blue, naslov: "Odabir MOD-a (za kombiniranu)", emoji: "🏦",
                  opis: "Nakon obavijesti REGOS-a birate MOD u FINA-i, Mirovinskom informativnom centru ili putem e-Građana. Ako ne odaberete u roku, REGOS vas raspoređuje po službenoj dužnosti.",
                  detalj: "Trenutno posluju dva MOD-a. Usporedite ponuđene iznose za istu vrstu mirovine." },
                { n: "4", boja: C.blue, naslov: "Jednokratna isplata do 20% (opcija)", emoji: "💵",
                  opis: "Možete zatražiti jednokratnu isplatu do 20 % kapitala prije početka doživotne mirovine. Uvjet: osnovna mirovina iz I. stupa mora biti veća od 115 % najniže mirovine.",
                  detalj: null },
                { n: "5", boja: C.purple, naslov: "Ugovor o mirovini s MOD-om", emoji: "📄",
                  opis: "MOD i vi potpisujete Ugovor o mirovini. U njemu se određuje: vrsta mirovine, iznos rente, uvjeti usklađivanja, eventualno zajamčeno razdoblje i ime korisnika.",
                  detalj: "Jednom sklopljeni ugovor nije moguće promijeniti. Zato je izbor vrste mirovine iznimno važna i TRAJNA odluka." },
                { n: "6", boja: C.green, naslov: "Isplata doživotne rente", emoji: "💰",
                  opis: "Mirovina iz II. stupa isplaćuje se svaki mjesec dok ste živi, bez obzira na duljinu života. Način usklađivanja s inflacijom navodi se u ugovoru.",
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

          {/* Izbor: samo I. stup ili I. + II. */}
          <Card style={{ marginBottom: 18, borderTop: `4px solid ${C.navy}` }}>
            <h3 style={{ fontSize: 14, color: C.navy, marginBottom: 10 }}>🔁 Izbor pri umirovljenju: samo I. stup ili I. + II. stup</h3>
            <ul style={{ fontSize: 12, color: C.slate, lineHeight: 1.75, paddingLeft: 18, margin: 0 }}>
              <li><strong>Tko bira:</strong> od 1.1.2019. svatko tko je u oba stupa, u postupku ostvarivanja starosne ili prijevremene
                starosne mirovine. Nije bitno kako ste ušli u II. stup.</li>
              <li><strong>Samo I. stup:</strong> mirovina se računa kao da ste uvijek bili samo u I. stupu, uz dodatak od 27 %.
                Sredstva s računa u II. stupu prenose se u državni proračun.</li>
              <li><strong>I. + II. stup:</strong> osnovna mirovina iz I. stupa (staž od 2002. računa se s faktorom oko 0,75) uz dodatak od
                27 %, plus doživotna mirovina iz II. stupa preko MOD-a. Moguće samo ako je osnovna mirovina veća od najniže.</li>
              <li><strong>Kako odlučiti:</strong> HZMO i REGOS prije odluke šalju informativni izračun obje mogućnosti. Izaberite onu s
                većim ukupnim iznosom; ako je razlika mala, razmislite i o vrsti mirovine u MOD-u (npr. obiteljskoj).</li>
              <li><strong>Odluka je trajna.</strong> Izjava se daje u Mirovinskom informativnom centru, u FINA-i ili putem e-usluge „Odabir mirovine“.</li>
            </ul>
            <div style={{ marginTop: 10, fontSize: 12, color: C.purple, fontWeight: 600 }}>
              → Okvirnu usporedbu za svoje podatke napravite u kalkulatoru „Samo I. ili I. + II.“ (izbornik Odluke pred mirovinu).
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
              prije odabira. Ako imate ovisne članove obitelji, razmatrajte zajedničku mirovinu ili zajamčeno razdoblje,
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
                { pravo: "Pravo izbora i promjene fonda i kategorije",
                  kako: "FINA, Mirovinski informativni centri ili e-Građani (pravila i rokovi na regos.hr)",
                  izvor: "ZOMF" },
                { pravo: "Pravo na informacije o ulaganjima fonda",
                  kako: "Statut fonda, godišnje i polugodišnje izvješće — javno objavljeno",
                  izvor: "ZOMF čl. 75." },
                { pravo: "Pravo na jednokratnu isplatu do 20 %",
                  kako: "Pri umirovljenju, ako je osnovna mirovina iz I. stupa veća od 115 % najniže mirovine",
                  izvor: "Zakon o MOD-ovima" },
                { pravo: "Pravo izbora mirovine samo iz I. stupa",
                  kako: "Svi u oba stupa, izjavom pri ostvarivanju starosne ili prijevremene mirovine (od 2019.)",
                  izvor: "ZOMO" },
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
            letterSpacing: "0.08em", marginBottom: 6 }}>Državni poticaj (DPS)</div>
          <p style={{ fontSize: 13, color: "rgba(255,255,255,0.85)", lineHeight: 1.7 }}>
            Hrvatska država daje <strong style={{ color: "#7bffb0" }}>15% od vaše godišnje uplate</strong> besplatno — do
            maksimalnih <strong style={{ color: "#7bffb0" }}>{DPS_MAX.toFixed(2)} EUR godišnje</strong>.
            Ostvaruje se uplatom <strong style={{ color: "#7bffb0" }}>{DPS_MAX_UPLATA} EUR godišnje</strong> ({(DPS_MAX_UPLATA/12).toFixed(2)} EUR/mj.).
            <br/>U fazi štednje <strong style={{ color: "#7bffb0" }}>nema poreza na prinos</strong>. Isplata rente oporezuje se porezom na dohodak uz primjenu osobnog odbitka; prijevremeni otkup prije 55. god. povlači i povrat DPS-a.
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
              explain={<div><p style={{ fontSize: 12, lineHeight: 1.7, color: C.slate }}>Sredstva su dostupna od 55. godine. Što ranije počnete, to više zarađuje složena kamata: kamata se računa i na prethodno zarađenu kamatu.</p></div>}
            />
            <SmartParam icon="📈" title="Godišnji prinos fonda" value={prinos} min={0.01} max={0.10} step={0.005}
              format={v => `${(v*100).toFixed(1)}%`} onChange={setPrinos}
              zones={[
                { upTo: 0.034, color: C.amber, label: "Konzervativno" },
                { upTo: 0.065, color: C.green, label: "Prosjek DMF-ova" },
                { upTo: 0.10,  color: C.blue,  label: "Optimistično" },
              ]}
              dynamicExplain={v => `Uz ${(v*100).toFixed(1)}% prinos, na ${god} god. štednje prinos generira ${fmt(kapital - uplata*12*god - dps_total)} EUR iznad direktnih uplata i DPS-a.`}
              explain={<div><p style={{ fontSize: 12, lineHeight: 1.7, color: C.slate }}>Dobrovoljni fondovi razlikuju se po strategiji ulaganja (udjelu dionica). Prinos nije zajamčen; povijesne prinose pojedinih fondova možete usporediti na hanfa.hr i stranicama društava.</p></div>}
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
              <ResultBox label="Prinos, DPS i uplate poslodavca" value={`${fmt(kapital - uplata*12*god)} EUR`} color={C.purple} />
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
              ["Bez poreza na prinos tijekom štednje", "Obična štednja: 12 % porez na kamate"],
              ["Dugoročno viši očekivani prinos od depozita", "Ali bez jamstva prinosa"],
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
  const [clanova, setClanova] = useState(1);

  // MF od 1.1.2026.: potpuna = 1.1, djelomična = 0.9
  const mf_map = { potpuna: 1.1, djelomicna: 0.9, profesionalna: 0.9 };
  const mf = mf_map[vrsta] || 1.0;
  const res = calcStup1({ pvb, staz, pf: 1.0, mf, avm: AVM });
  const resD = saDodatkom(res);

  // Obiteljska mirovina: mirovinski faktor ovisi o broju članova obitelji (ZOMO čl. 85.)
  const MF_OBITELJ = { 1: 0.77, 2: 0.88, 3: 1.0, 4: 1.1 };
  const mfObitelj = MF_OBITELJ[clanova];
  const obiteljna = saDodatkom(calcStup1({ pvb, staz, pf: 1.0, mf: mfObitelj, avm: AVM })).ukupno;
  const obitelj1 = saDodatkom(calcStup1({ pvb, staz, pf: 1.0, mf: 0.77, avm: AVM })).ukupno;
  const dioObiteljske = Math.min(Math.max(obitelj1 * 0.27, 3 * AVM), obitelj1 * 0.5);

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>Invalidska i obiteljska mirovina</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>Zaštita kad osiguranik ne može raditi ili umre</p>
      </div>

      <Ukratko color={C.red}>
        <strong>Invalidska mirovina</strong> pripada vam ako zbog bolesti ili ozljede izgubite radnu sposobnost. Računa se kao
        starosna, ali s mirovinskim faktorom 1,1 (potpuni gubitak) ili 0,9 (djelomični gubitak).{" "}
        <strong>Obiteljska mirovina</strong> pripada bračnom drugu, djeci i uzdržavanim roditeljima umrlog osiguranika ili
        umirovljenika; iznos ovisi o broju članova obitelji koji je primaju.
      </Ukratko>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        {/* Invalidska */}
        <div>
          <div style={{ background: `linear-gradient(135deg, ${C.red}cc, #8b0000)`, borderRadius: 14,
            padding: "16px 20px", marginBottom: 14, color: "#fff" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#ffaaaa", textTransform: "uppercase", marginBottom: 6 }}>Invalidska mirovina</div>
            <p style={{ fontSize: 12, color: "rgba(255,255,255,0.85)", lineHeight: 1.7 }}>
              Formula je ista kao za starosnu mirovinu, a razlikuje se samo <strong style={{ color: "#ffaaaa" }}>mirovinski faktor (MF)</strong>.
              Od 1.1.2026. povećan je na 1,1 (potpuni gubitak) i 0,9 (djelomični gubitak radne sposobnosti).
            </p>
          </div>
          <Card style={{ marginBottom: 14 }}>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Vrsta invalidske mirovine</h3>
            {[
              { v: "potpuna", emoji: "🚫", l: "Potpuni gubitak radne sposobnosti", mf: "MF = 1,1 (od 1.1.2026.)", boja: C.red,
                opis: "Osiguranik ne može raditi. Od 1.1.2026. smije raditi do 3,5 sata dnevno uz isplatu mirovine." },
              { v: "djelomicna", emoji: "⚕️", l: "Djelomični gubitak radne sposobnosti", mf: "MF = 0,9 (od 1.1.2026.)", boja: C.amber,
                opis: "Osiguranik može raditi na drugom, primjerenom poslu. Mirovina se može primati uz rad." },
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
            <SmartParam icon="💰" title="Plaća u odnosu na prosjek RH (PVB)" value={pvb} min={0.3} max={3.8} step={0.05}
              format={v => v.toFixed(2)} onChange={setPvb} zones={[
                { upTo: 0.99, color: C.amber, label: "Ispod prosjeka" },
                { upTo: 1.09, color: C.green, label: "Prosjek" },
                { upTo: 3.8,  color: C.blue,  label: "Iznad prosjeka" },
              ]}
              dynamicExplain={v => `PVB ${v.toFixed(2)} znači plaću oko ${fmt(v * PROSJ_BRUTO, 0)} EUR bruto (prema prosjeku 2025.).`}
              explain={null} />
            <SmartParam icon="📅" title="Mirovinski staž u trenutku invalidnosti" value={staz} min={1} max={45} step={1}
              format={v => `${v} god.`} onChange={setStaz} zones={[
                { upTo: 4,  color: C.red,   label: "⚠ Ispod minimuma za neke slučajeve" },
                { upTo: 14, color: C.amber, label: "Kratki staž" },
                { upTo: 45, color: C.green, label: "Dobar staž" },
              ]}
              dynamicExplain={v => `Sa ${v} god. staža najniža mirovina iznosi ${fmt(MIN_MIROVINA_FAKTOR * AVM * v * mf)} EUR.`}
              explain={<div><p style={{ fontSize:12, lineHeight:1.7, color:C.slate }}>Ako je invalidnost posljedica bolesti ili ozljede izvan rada, potreban je minimalni staž koji ovisi o dobi (otprilike trećina radnog vijeka). Kod ozljede na radu ili profesionalne bolesti staž nije uvjet, a najniža mirovina računa se kao za 40 godina staža.</p></div>} />
            <div style={{ padding:"10px 12px", background: vrsta==="potpuna" ? C.redL : C.amberL, borderRadius:9, textAlign:"center" }}>
              <div style={{ fontSize:11, color:C.slate }}>Invalidska mirovina / mj.</div>
              <div style={{ fontSize:26, fontWeight:800, color: vrsta==="potpuna" ? C.red : C.amber, fontFamily:"'Playfair Display',serif" }}>
                {fmt(resD.ukupno)} EUR
              </div>
              <div style={{ fontSize:11, color:C.slateL, marginTop:2 }}>
                MF = {mf.toFixed(1).replace(".", ",")} · {resD.najniza ? "najniža mirovina" : "uključen dodatak 27 %"}
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
              Pravo stječu članovi obitelji umrlog osiguranika ili korisnika mirovine. Iznos se računa iz mirovine
              umrlog, s mirovinskim faktorom koji ovisi o broju članova obitelji: 0,77 / 0,88 / 1,0 / 1,1.
            </p>
          </div>
          <Card style={{ marginBottom: 14 }}>
            <h3 style={{ fontSize: 14, marginBottom: 10 }}>Tko ima pravo?</h3>
            <ul style={{ fontSize: 12, color: C.slate, lineHeight: 1.7, paddingLeft: 18, margin: "0 0 12px" }}>
              <li><strong>Bračni ili izvanbračni drug:</strong> ako je do smrti navršio 50 godina (s navršenih 45 pravo stječe s 50),
                ako je potpuno nesposoban za rad ili ako skrbi o djeci koja imaju pravo na obiteljsku mirovinu.</li>
              <li><strong>Djeca:</strong> do 15. godine, a za vrijeme redovitog školovanja najdulje do 26. godine; trajno ako su
                potpuno nesposobna za rad.</li>
              <li><strong>Roditelji, braća i sestre:</strong> ako ih je umrli uzdržavao.</li>
            </ul>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.navy, marginBottom: 6 }}>Koliko članova obitelji prima mirovinu?</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 6 }}>
              {[1, 2, 3, 4].map(n => (
                <button key={n} onClick={() => setClanova(n)}
                  style={{ padding: "8px 4px", border: `2px solid ${clanova === n ? C.blue : C.border}`, borderRadius: 9,
                    background: clanova === n ? C.blueL : C.white, cursor: "pointer" }}>
                  <div style={{ fontSize: 13, fontWeight: 800, color: C.navy }}>{n === 4 ? "4+" : n}</div>
                  <div style={{ fontSize: 10, color: C.slateL }}>MF {MF_OBITELJ[n].toFixed(2).replace(".", ",")}</div>
                </button>
              ))}
            </div>
          </Card>
          <Card>
            <h3 style={{ fontSize: 14, marginBottom: 14 }}>Iznos obiteljske mirovine</h3>
            <SmartParam icon="💰" title="Plaća umrlog u odnosu na prosjek RH (PVB)" value={pvb} min={0.3} max={3.8} step={0.05}
              format={v => v.toFixed(2)} onChange={setPvb} zones={[
                { upTo: 0.99, color: C.amber }, { upTo: 1.09, color: C.green }, { upTo: 3.8, color: C.blue }
              ]}
              dynamicExplain={v => `PVB ${v.toFixed(2)} znači plaću oko ${fmt(v * PROSJ_BRUTO, 0)} EUR bruto (prema prosjeku 2025.).`}
              explain={null} />
            <SmartParam icon="📅" title="Staž umrlog osiguranika" value={staz} min={1} max={45} step={1}
              format={v => `${v} god.`} onChange={setStaz} zones={[
                { upTo: 14, color: C.amber }, { upTo: 45, color: C.green }
              ]}
              dynamicExplain={v => `Uvjet: umrli je imao najmanje 5 godina staža osiguranja ili 10 godina mirovinskog staža, ili je bio umirovljenik.`}
              explain={null} />
            <div style={{ padding:"10px 12px", background: C.blueL, borderRadius:9, textAlign:"center", marginBottom:10 }}>
              <div style={{ fontSize:11, color:C.slate }}>Obiteljska mirovina / mj. (ukupno za sve članove)</div>
              <div style={{ fontSize:26, fontWeight:800, color:C.blue, fontFamily:"'Playfair Display',serif" }}>
                {fmt(obiteljna)} EUR
              </div>
              <div style={{ fontSize:11, color:C.slateL, marginTop:2 }}>
                MF {String(mfObitelj).replace(".", ",")} · uključen dodatak 27 % (osim kad se isplaćuje najniža)
              </div>
            </div>
            <div style={{ fontSize:11, color:C.slate, lineHeight:1.6, padding:"8px 10px", background:C.greenL, borderRadius:8, marginBottom: 10 }}>
              <strong>Imate vlastitu mirovinu?</strong> Udovica ili udovac s navršenih 65 godina, koji je jedini korisnik, može uz svoju
              mirovinu primati i <strong>dio obiteljske mirovine</strong>: 27 % obiteljske mirovine za jednog člana
              (ovdje oko {fmt(dioObiteljske)} EUR), ako ukupne mirovine ne prelaze 80 AVM ({fmt(80 * AVM, 0)} EUR). Druga mogućnost je
              izabrati veću od dviju mirovina.
            </div>
            <div style={{ fontSize:11, color:C.slateL, lineHeight:1.6, padding:"8px 10px", background:C.bg, borderRadius:8 }}>
              ⚠ Ovi iznosi su aproksimativni. Stvarni uvjeti i iznosi ovisni su o konkretnim okolnostima.
              Za točne podatke obratite se HZMO-u ili zatražite informativni izračun u e-Građanima (HZMO e-usluge).
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
  const [trenutnaPlacha, setTrenutnaPlaca] = useState(PROSJ_BRUTO);
  const [trenutniStaz, setTrenutniStaz] = useState(20);
  const [dob, setDob] = useState(42);
  const [spol, setSpol] = useState("M");
  // Realni prinos (iznad rasta plaća), da projekcija bude u današnjem novcu kao i I. stup
  const [prinos2] = useState(0.03);

  const mirDob = dobZaStarosnu(spol, TEKUCA_GODINA - dob);
  const godDoMir = Math.max(0, Math.ceil(mirDob - dob));
  const buduciStaz = trenutniStaz + godDoMir;
  const pvb_est = trenutnaPlacha / PROSJ_BRUTO;
  const godUStupu2 = Math.min(trenutniStaz, TEKUCA_GODINA - 2002) + godDoMir;
  const projekcija = (pvb, staz, godII, pf = 1) => {
    const k = calcStup2FV({ bruto_placa: pvb * PROSJ_BRUTO, god_stednje: godII, neto_prinos: prinos2 });
    const renta = rentaDozivotna(k * 0.8);
    return { k, ...kombiniranaMirovina({ pvb, staz, stazOd2002: Math.min(staz, godII), pf, rentaII: renta }) };
  };
  const proj = projekcija(pvb_est, buduciStaz, godUStupu2);
  const r1 = { efektivna: proj.osnovna };
  const k2 = proj.k;
  const renta2 = proj.renta;
  const dod = proj.dodatak;
  const ukupno = proj.ukupno;
  const gap = zeljenaMirovina - ukupno;
  const gap_poz = gap > 0;

  // Što treba promijeniti za doći do željene mirovine (mirovina raste približno razmjerno plaći)
  const potrebnaPvb = pvb_est * zeljenaMirovina / ukupno;
  const potrebnaPlaca = potrebnaPvb * PROSJ_BRUTO;
  const dodatneGodine = (() => {
    for (let g = 0; g <= 15; g++) {
      if (projekcija(pvb_est, buduciStaz + g, godUStupu2 + g, 1 + 0.054 * g).ukupno >= zeljenaMirovina) return g;
    }
    return null;
  })();
  const potrebniStup3_god = () => {
    if (!gap_poz) return 0;
    const r = 0.05 / 12; const n = Math.max(1, godDoMir) * 12;
    const k_pot = gap * 18 * 12;
    const ukupnaUplata = k_pot * r / (Math.pow(1 + r, n) - 1);
    const maxSPoticajem = (DPS_MAX_UPLATA_GOD / 12) * (1 + DPS_STOPA);
    return ukupnaUplata <= maxSPoticajem ? ukupnaUplata / (1 + DPS_STOPA) : ukupnaUplata - (DPS_MAX_UPLATA_GOD / 12) * DPS_STOPA;
  };
  const stup3_mj = potrebniStup3_god();

  const progressData = [
    { name: "I. stup", value: Math.round(r1.efektivna), fill: C.navy },
    { name: "Dodatak 27 %", value: Math.round(dod), fill: C.gold },
    { name: "II. stup", value: Math.round(renta2), fill: C.purple },
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
            {[["M","Muškarac","Starosna mirovina s 65 god."],["Ž","Žena",`Za vas: s ${fmtDob(dobZaStarosnu("Ž", TEKUCA_GODINA - dob))}`]].map(([v,l,s]) => (
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
              { upTo: 799,  color: C.amber,  label: "Skromna mirovina" },
              { upTo: 1199, color: C.gold,   label: "Prosječna" },
              { upTo: 1799, color: C.green,  label: "Udobna mirovina" },
              { upTo: 3000, color: C.blue,   label: "Visoka mirovina" },
            ]}
            dynamicExplain={v => `To je ${pct(v/trenutnaPlacha)} vaše sadašnje bruto plaće. Prosječna mirovina u Hrvatskoj znatno je niža od plaće, pa je razlika uobičajena.`}
            explain={null} />
          <SmartParam icon="💼" title="Trenutna bruto plaća" value={trenutnaPlacha} min={700} max={5000} step={50}
            format={v => `${v.toLocaleString("hr-HR")} EUR`} onChange={setTrenutnaPlaca}
            zones={[
              { upTo: 1099, color: C.amber }, { upTo: 1499, color: C.green }, { upTo: 5000, color: C.blue }
            ]}
            dynamicExplain={v => `Vaša plaća je ${(v/PROSJ_BRUTO).toFixed(2)} prosječne bruto plaće u RH (2.016 EUR u 2025.). Kalkulator pretpostavlja da je tako bilo kroz cijelu karijeru.`}
            explain={null} />
          <SmartParam icon="📅" title="Trenutni mirovinski staž" value={trenutniStaz} min={0} max={45} step={1}
            format={v => `${v} god.`} onChange={setTrenutniStaz}
            zones={[
              { upTo: 14, color: C.amber }, { upTo: 29, color: C.gold }, { upTo: 45, color: C.green }
            ]}
            dynamicExplain={v => `Do starosne mirovine ostaje još ${godDoMir} god. Na kraju karijere imat ćete oko ${v + godDoMir} god. staža.`}
            explain={null} />
          <SmartParam icon="🎂" title="Trenutna dob" value={dob} min={18} max={67} step={1}
            format={v => `${v} god.`} onChange={setDob}
            zones={[
              { upTo: 34, color: C.green, label: "Puno vremena za planiranje" },
              { upTo: 49, color: C.gold,  label: "Sredina karijere" },
              { upTo: 59, color: C.amber, label: "Bliže kraju karijere" },
              { upTo: 67, color: C.red,   label: "Kratko do mirovine" },
            ]}
            dynamicExplain={v => `Starosna mirovina za vas: s ${fmtDob(mirDob)}. Preostaje još ${godDoMir} god.`}
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
            <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 4 }}>Struktura projektirane mirovine</div>
            <div style={{ fontSize: 10, color: C.slateL, marginBottom: 10, lineHeight: 1.5 }}>
              Iznosi su u današnjem novcu. Pretpostavke: plaća prati prosjek RH, II. stup od 2002. (ili od početka rada) s realnim
              prinosom 3 % godišnje iznad rasta plaća, 20 % kapitala podignuto jednokratno, uključen dodatak 27 % na osnovnu mirovinu.
            </div>
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
                  { emoji: "🌱", l: "Mjesečna uplata u III. stup",
                    v: `${stup3_mj > 0 ? fmt(stup3_mj) : "Nije potrebno"} EUR/mj.`,
                    opis: "Uz državni poticaj (15 % uplate, do 99,54 EUR godišnje) i 5 % prinosa pokriva razliku kroz 18 godina mirovine", color: C.green },
                  { emoji: "💼", l: "Potrebna plaća za željenu mirovinu",
                    v: `${fmt(potrebnaPlaca)} EUR bruto`,
                    opis: `Prosječno kroz karijeru — ${Math.abs((potrebnaPvb-1)*100).toFixed(0)} % ${potrebnaPvb > 1 ? "iznad" : "ispod"} prosjeka RH`, color: C.blue },
                  { emoji: "⏳", l: "Rad nakon starosne dobi",
                    v: dodatneGodine === null ? "više od 15 god." : dodatneGodine === 0 ? "nije potrebno" : `${dodatneGodine} god. dulje`,
                    opis: "Svaki mjesec rada nakon zakonske dobi trajno povećava mirovinu za 0,45 % (2 godine = +10,8 %), uz dodatni staž", color: C.gold },
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
  const [placha, setPlacha] = useState(PROSJ_BRUTO);
  const [stazGod, setStazGod] = useState(40);

  const hr_2 = rentaDozivotna(calcStup2FV({ bruto_placa: placha, god_stednje: stazGod, neto_prinos: 0.055 }) * 0.8);
  const hr_k = kombiniranaMirovina({ pvb: placha / PROSJ_BRUTO, staz: stazGod, rentaII: hr_2 });
  const hr_neto = calcNetoMirovina(hr_k.ukupno).neto;
  const hr_stopa = hr_k.ukupno / placha;

  const sustavi = [
    { zemlja: "🇭🇷 Hrvatska", kod: "HR", mercer: "B (68,7)", score: 68.7, boja: C.red,
      opis: "Miješani sustav: I. stup (generacijska solidarnost, 15 % plaće) i II. stup (osobna štednja u fondu, 5 % plaće), uz dobrovoljni III. stup s državnim poticajem.",
      doprinos: "20 % (15 % I. + 5 % II.)", dob: "65 M / 64 Ž (2026.), izjednačavanje do 2030.",
      prednosti: ["Dobar nadzor i transparentnost (Mercer: integritet A)", "II. stup je osobna imovina", "Državni poticaj za III. stup"],
      slabosti: ["Mirovine niske u odnosu na plaće", "Nepovoljan omjer radnika i umirovljenika", "Malo strukovnih (poslodavačkih) mirovinskih planova"] },
    { zemlja: "🇳🇱 Nizozemska", kod: "NL", mercer: "A (85,4)", score: 85.4, boja: C.blue,
      opis: "Državna osnovna mirovina (AOW) za sve, kvazi-obvezne strukovne mirovine preko poslodavaca i dobrovoljna štednja. Strukovni fondovi prelaze na nove ugovore do 2028.",
      doprinos: "Visok, većinom poslodavac", dob: "67 (67 i 3 mj. od 2028.)",
      prednosti: ["Gotovo svi zaposleni u strukovnim planovima", "Visoke mirovine u odnosu na plaće", "Najbolji ukupni rezultat"],
      slabosti: ["Visoki troškovi doprinosa", "Složena reforma sustava"] },
    { zemlja: "🇮🇸 Island", kod: "IS", mercer: "A (84,0)", score: 84.0, boja: "#003897",
      opis: "Državna mirovina uz obvezne strukovne fondove u koje se uplaćuje najmanje 15,5 % plaće.",
      doprinos: "min. 15,5 % (4 % radnik + 11,5 % poslodavac)", dob: "67",
      prednosti: ["Visoka obvezna kapitalizirana štednja", "Dobra održivost"],
      slabosti: ["Malo tržište osjetljivo na vanjske šokove"] },
    { zemlja: "🇩🇰 Danska", kod: "DK", mercer: "A (82,3)", score: 82.3, boja: "#c41e3a",
      opis: "Državna osnovna mirovina (Folkepension), obvezni ATP za sve zaposlene i strukovne mirovine iz kolektivnih ugovora.",
      doprinos: "Strukovne ~12–15 % + ATP", dob: "67, raste s očekivanim trajanjem života (68 od 2030.)",
      prednosti: ["Visoka pokrivenost zaposlenih", "Dob umirovljenja vezana uz trajanje života"],
      slabosti: ["Visoki porezi za financiranje", "Sve kasnije umirovljenje"] },
    { zemlja: "🇸🇬 Singapur", kod: "SG", mercer: "A (80,8)", score: 80.8, boja: "#EF3340",
      opis: "CPF: obvezni osobni račun za mirovinu, zdravstvo i stanovanje. Za račune namijenjene mirovini država jamči minimalnu kamatu oko 4 %.",
      doprinos: "do 37 % (20 % radnik + 17 % poslodavac, do 55 god.)", dob: "63 (isplata CPF-a od 65)",
      prednosti: ["Zajamčena minimalna kamata na mirovinske račune", "Kapital je nasljediv"],
      slabosti: ["Vrlo visoki doprinosi", "Ograničena fleksibilnost"] },
    { zemlja: "🇦🇺 Australija", kod: "AU", mercer: "B+ (77,6)", score: 77.6, boja: "#00843D",
      opis: "Superannuation: obvezna osobna štednja koju uplaćuje poslodavac, uz državnu Age Pension za one s manje imovine.",
      doprinos: "12 % plaće (poslodavac, od 1.7.2025.)", dob: "67",
      prednosti: ["Kapital je osobna i nasljediva imovina", "Porezne olakšice"],
      slabosti: ["Tržišni rizik snosi pojedinac", "Niže mirovine za niske plaće"] },
  ];

  const chartData = sustavi.map(s => ({ name: s.zemlja.split(" ")[1], score: s.score, fill: s.boja }));
  const [selected, setSelected] = useState("HR");
  const sel = sustavi.find(s => s.kod === selected) || sustavi[0];

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 16 }}>
        <h2 style={{ fontSize: 22, marginBottom: 4 }}>Hrvatska i najbolji mirovinski sustavi</h2>
        <p style={{ color: C.slateL, fontSize: 13 }}>Mercer CFA Institute Global Pension Index 2025 (52 sustava)</p>
      </div>

      <Ukratko color={C.blue}>
        Mercer svake godine ocjenjuje mirovinske sustave prema primjerenosti mirovina, održivosti i nadzoru. Hrvatska je 2025.
        dobila ocjenu <strong>B (68,7 bodova)</strong>. Najbolje ocijenjeni sustavi imaju više obvezne štednje: Nizozemska,
        Island i Danska kroz strukovne mirovine preko poslodavaca. Pouka za vas: uz I. i II. stup dodatna štednja (III. stup)
        bitno podiže mirovinu.
      </Ukratko>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 20, marginBottom: 20 }}>
        <Card>
          <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 12 }}>Mercer ocjena 2025. (bodovi)</div>
          <ResponsiveContainer width="100%" height={210}>
            <BarChart data={chartData} margin={{ top: 5, right: 10, bottom: 5, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
              <XAxis dataKey="name" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} domain={[0, 100]} />
              <Tooltip formatter={v => [`${String(v).replace(".", ",")} bodova`, "Mercer"]} />
              <Bar dataKey="score" radius={[4, 4, 0, 0]}>
                {chartData.map((d, i) => <Cell key={i} fill={d.fill} opacity={0.9} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 12 }}>Odaberite sustav</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 14 }}>
            {sustavi.map(s => (
              <button key={s.kod} onClick={() => setSelected(s.kod)} style={{
                padding: "5px 12px", border: `2px solid ${selected === s.kod ? s.boja : C.border}`,
                borderRadius: 20, background: selected === s.kod ? s.boja : C.white,
                color: selected === s.kod ? "#fff" : C.navy, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
                {s.zemlja}
              </button>
            ))}
          </div>
          <div style={{ background: sel.boja + "0a", border: `1.5px solid ${sel.boja}22`, borderRadius: 10, padding: 14 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
              <span style={{ fontSize: 14, fontWeight: 700 }}>{sel.zemlja}</span>
              <span style={{ fontSize: 11, fontWeight: 700, color: sel.boja, background: sel.boja + "18", padding: "2px 9px", borderRadius: 10 }}>Mercer {sel.mercer}</span>
            </div>
            <p style={{ fontSize: 12, color: C.slate, lineHeight: 1.7, marginBottom: 10 }}>{sel.opis}</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 10 }}>
              <div style={{ background: C.bg, borderRadius: 7, padding: "7px 10px" }}>
                <div style={{ fontSize: 10, color: C.slateL }}>Doprinosi</div>
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
                {sel.prednosti.map((p, i) => <div key={i} style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>• {p}</div>)}
              </div>
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: C.red, marginBottom: 4 }}>⚠ Slabosti</div>
                {sel.slabosti.map((s, i) => <div key={i} style={{ fontSize: 11, color: C.slate, lineHeight: 1.5 }}>• {s}</div>)}
              </div>
            </div>
          </div>
        </Card>
      </div>

      <Card style={{ marginBottom: 20 }}>
        <h3 style={{ fontSize: 14, marginBottom: 14 }}>Koliko je mirovina u odnosu na plaću u Hrvatskoj (vaš primjer)</h3>
        <div style={{ display: "flex", gap: 16, marginBottom: 16, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 220 }}>
            <MiniSlider label="Bruto plaća" value={placha} min={MIN_PLACA} max={5000} step={50}
              format={v => `${v.toLocaleString("hr-HR")} EUR/mj.`} onChange={setPlacha} color={C.navy} />
          </div>
          <div style={{ flex: 1, minWidth: 220 }}>
            <MiniSlider label="Godine staža (sve u oba stupa)" value={stazGod} min={15} max={45} step={1}
              format={v => `${v} god.`} onChange={setStazGod} color={C.navy} />
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 10 }}>
          {[
            { l: "Mirovina bruto (I. + II.)", v: `${fmt(hr_k.ukupno)} EUR` },
            { l: "Neto 2026.", v: `${fmt(hr_neto)} EUR` },
            { l: "U odnosu na bruto plaću", v: pct(hr_stopa) },
          ].map(it => (
            <div key={it.l} style={{ background: C.bg, borderRadius: 10, padding: "12px 14px", textAlign: "center" }}>
              <div style={{ fontSize: 11, color: C.slateL, marginBottom: 4 }}>{it.l}</div>
              <div style={{ fontSize: 20, fontWeight: 800, color: C.navy, fontFamily: "'Playfair Display',serif" }}>{it.v}</div>
            </div>
          ))}
        </div>
        <div style={{ marginTop: 10, fontSize: 11, color: C.slateL, lineHeight: 1.6 }}>
          Pretpostavke: plaća jednaka cijeli radni vijek, cijeli staž u oba stupa, prinos II. stupa 5,5 %, 20 % kapitala podignuto jednokratno.
        </div>
      </Card>

      <Detalji naslov="Što možemo naučiti od boljih sustava?">
        <ul style={{ fontSize: 12, color: C.slate, lineHeight: 1.75, paddingLeft: 18, margin: 0 }}>
          <li><strong>Više obvezne štednje:</strong> Island najmanje 15,5 % plaće, Singapur do 37 %, Australija 12 % (poslodavac), Hrvatska 5 % u II. stupu.</li>
          <li><strong>Strukovne mirovine:</strong> u Nizozemskoj i Danskoj gotovo svi zaposleni štede preko poslodavca. U Hrvatskoj postoje zatvoreni
            dobrovoljni fondovi poslodavaca, ali ih koristi malo ljudi. Poslodavac vam može uplaćivati do 804 EUR godišnje bez poreza.</li>
          <li><strong>Dob umirovljenja i trajanje života:</strong> Danska povezuje dob umirovljenja s očekivanim trajanjem života; u Hrvatskoj je dob propisana zakonom.</li>
          <li><strong>Zajamčeni prinos:</strong> Singapur jamči minimalnu kamatu na mirovinske račune; hrvatski fondovi nemaju jamstvo prinosa.</li>
        </ul>
      </Detalji>
    </div>
  );
}

// ── PRAKTIČNI VODIČ ────────────────────────────────────────────────────────
function VodicView() {
  const [dob, setDob] = useState(35);

  const profili = [
    { range: [18,29], label: "18–29 god.", emoji: "🌱", color: C.green,
      naslov: "Mladi radnik — zlatno doba za početak",
      subtitle: "Što ranije počnete štedjeti, to više za vas radi složena kamata",
      koraci: [
        { emoji: "✅", t: "Provjeri je li poslodavac prijavio u II. stup", opis: "Provjeri e-Građani / REGOS portal da li se svaki mjesec uplaćuje 5% bruto na tvoj osobni račun. Ako ne — poslodavac krši zakon.", hitno: true },
        { emoji: "📂", t: "Razmisli o kategoriji A", opis: "Ako ne izabereš, REGOS te raspoređuje u kategoriju B. Kategorija A ima više dionica: veće oscilacije, ali dugoročno viši očekivani prinos. S 30+ godina do mirovine imaš vremena prebroditi padove.", hitno: false },
        { emoji: "🌱", t: "Otvori III. stup ODMAH", opis: `Uplati samo 55,30 EUR/mj. (663,61 EUR/god.) i dobij maksimalni DPS od 99,54 EUR/god. Ukupno s prinosom: na 40 god. to je ~${fmt(calcStup2FV({bruto_placa:55.3*20,god_stednje:40,neto_prinos:0.055}))} EUR.`, hitno: true },
        { emoji: "📊", t: "Razumi PVB — pregovaraj o plaći!", opis: "Svaki % povećanja plaće povećava tvoj PVB i buduću mirovinu. Plaća od 3.000 EUR vs prosjek 2.016 EUR = PVB 1,49 vs 1,00 = oko 49 % veća mirovina iz I. stupa.", hitno: false },
        { emoji: "🔍", t: "Provjeri stanje računa jednom godišnje", opis: "e-Građani → REGOS → Moj mirovinski račun. Provjeri ukupan kapital, uplate, prinos i historiju.", hitno: false },
      ],
    },
    { range: [30,44], label: "30–44 god.", emoji: "💼", color: C.blue,
      naslov: "Sredina karijere — optimizacija i planiranje",
      subtitle: "Rast karijere i plaće direktno povećavaju mirovinu",
      koraci: [
        { emoji: "📈", t: "Analiziraj PVB kretanje", opis: "HZMO može dati informativni izračun mirovine s trenutnim PVB-om. Ako si napredovao u karijeri i plaća ti raste, PVB raste — ali provjeri.", hitno: false },
        { emoji: "🏠", t: "Razmotri balans: mirovina vs. nekretnina vs. investicije", opis: "Nekretnina, ETF-ovi i dionice nisu mirovinska štednja — ali mogu biti dodatni izvor prihoda (najam, dividende, prodaja udjela) u mirovini. Dugoročno ulaganje u globalne ETF-ove (npr. preko brokera) nudi veću likvidnost i diverzifikaciju od nekretnine, uz slično ili veće dugoročno povećanje. Nemoj zanemariti financijsku mirovinu (II. i III. stup) dok ulažeš u nekretninu ili portfelj dionica/ETF-ova.", hitno: false },
        { emoji: "🌱", t: "Iskoristi III. stup preko poslodavca", opis: "Ako imaš poslodavca koji uplaćuje u III. stup — to je povećanje plaće bez poreza. Do 804 EUR/god. od poslodavca neoporezivo.", hitno: true },
        { emoji: "📊", t: "Kalkulator: koliko mi nedostaje?", opis: "Koristi tab 'Planer' u ovoj aplikaciji. Unesi željenu mirovinu i vidi točno koliko treba prilagoditi štednju, plaću ili datum umirovljenja.", hitno: false },
        { emoji: "🔄", t: "Provjeri kategoriju OMF", opis: "Ako ne izabereš, REGOS te raspoređuje u kategoriju B. Uz više od 10 godina do mirovine kategorija A dugoročno nosi veći očekivani prinos, ali i veće oscilacije. Provjeri na regos.hr.", hitno: false },
        { emoji: "👶", t: "Porodiljni dopust = mirovinski staž!", opis: "Od 1.7.2025. majci (ili posvojiteljici) se za svako dijete automatski dodaje 12 mj. mirovinskog staža. Provjeri da je HZMO ispravno evidentirao.", hitno: false },
      ],
    },
    { range: [45,54], label: "45–54 god.", emoji: "⚙️", color: C.gold,
      naslov: "Pred-mirovinski period — zaštita i prijelaz",
      subtitle: "Počnite smanjivati investicijski rizik i planirati isplatu",
      koraci: [
        { emoji: "🔄", t: "Prelazak OMF kategorije A → B", opis: "Kad do mirovine ostane manje od 10 godina, REGOS te iz kategorije A prebacuje u B (na zahtjev možeš ostati u A do 5 godina prije mirovine). U zadnjih 5 godina razmisli o kategoriji C.", hitno: true },
        { emoji: "📋", t: "Informativni izračun mirovine na HZMO", opis: "Podnesi zahtjev za informativni izračun na mojamirovina.hr (e-Građani). Vidjet ćeš točan PVB, staž i projiciranu mirovinu.", hitno: true },
        { emoji: "⚖️", t: "Samo I. stup ili I. + II.?", opis: "Pri umirovljenju svatko u oba stupa bira između mirovine samo iz I. stupa i mirovine iz oba stupa. Isprobaj kalkulator „Samo I. ili I. + II.“ sa stanjem svog računa u II. stupu.", hitno: false },
        { emoji: "🌱", t: "Pojačaj III. stup ove dekade", opis: `Npr. 150 EUR mjesečno kroz 10 godina uz državni poticaj i 5 % prinosa daje oko ${fmt(calcStup2FV({bruto_placa: (150 + 99.54/12) * 20, god_stednje: 10, neto_prinos: 0.05}), 0)} EUR. Novac je dostupan od 55. godine.`, hitno: false },
        { emoji: "💼", t: "Planiranje dohotka od imovine u mirovini", opis: "Nekretnine, dividende, umirovljenički fond — planirajte sve izvore prihoda, ne samo državnu mirovinu.", hitno: false },
      ],
    },
    { range: [55,64], label: "55–64 god.", emoji: "🏁", color: C.purple,
      naslov: "Finalna ravnina — maksimizacija i odluke",
      subtitle: "Svaka godina rada sada vrijedi više nego ikad",
      koraci: [
        { emoji: "⏳", t: "Razmotri odgodu — +0,45 % po mjesecu", opis: "Svaki mjesec rada nakon zakonske dobi trajno povećava mirovinu iz I. stupa za 0,45 % (24 mj. = +10,8 %), uz dodatni staž. Odluka ovisi o zdravlju i financijama.", hitno: false },
        { emoji: "🏦", t: "Istraži MOD-ove i vrstu mirovine", opis: "Usporedite ponude oba MOD-a (Hrvatski MOD i Raiffeisen MOD). Vrsta mirovine se poslije ne može promijeniti. Ako imate bračnog druga, razmotrite obiteljsku mirovinu ili zajamčeno razdoblje.", hitno: true },
        { emoji: "⚖️", t: "Samo I. stup ili I. + II.?", opis: "Najvažnija odluka pri umirovljenju. Zatražite informativni izračun obje mogućnosti i usporedite u kalkulatoru „Samo I. ili I. + II.“.", hitno: true },
        { emoji: "💵", t: "Jednokratna isplata do 20 %?", opis: "Moguća ako je osnovna mirovina iz I. stupa veća od 115 % najniže. Više podignuto sada znači nižu doživotnu mirovinu iz II. stupa.", hitno: false },
        { emoji: "💰", t: "III. stup je dostupan od 55. godine", opis: "Pravo na isplatu imate s 55 godina, bez obzira radite li. Načine isplate provjerite kod svog dobrovoljnog fonda.", hitno: false },
        { emoji: "📋", t: "Prijavite se HZMO-u najmanje 3 mj. unaprijed", opis: "Postupak ostvarivanja prava traje — nemojte čekati zadnji tren. Zahtjev za starosnu mirovinu podnosi se u nadležnoj ispostavi HZMO-a ili e-Građani.", hitno: true },
        { emoji: "📄", t: "Provjerite sve radne knjižice i staž", opis: "Svaki neevidentirani staž (stari podaci, vojni rok) treba biti u evidenciji HZMO-a. Elektronički zapis (e-radna knjižica) vidite u e-Građanima.", hitno: true },
      ],
    },
    { range: [65,75], label: "65+ god.", emoji: "🌅", color: C.navy,
      naslov: "U mirovini ili pred samim odlaskom",
      subtitle: "Iskoristite prava koja vam pripadaju",
      koraci: [
        { emoji: "💼", t: "Rad uz mirovinu", opis: "Od 1.1.2026. nakon 65. godine možete raditi i više od pola radnog vremena; tada se isplaćuje 50 % mirovine. Uz rad do pola radnog vremena mirovina se isplaćuje u cijelosti.", hitno: false },
        { emoji: "🎁", t: "Godišnji dodatak na mirovinu", opis: "Svaki prosinac HZMO isplaćuje godišnji dodatak: za 2026. iznosi 6 EUR po godini staža (npr. 40 god. = 240 EUR). Zahtjev nije potreban.", hitno: false },
        { emoji: "👫", t: "Dio obiteljske mirovine", opis: "Ako ste udovica ili udovac s vlastitom mirovinom, uz nju možete primati i 27 % obiteljske mirovine (uz uvjete). Provjerite kod HZMO-a.", hitno: true },
        { emoji: "📉", t: "Prijevremena mirovina — penal prestaje sa 70", opis: "Ako ste otišli u prijevremenu mirovinu, HZMO s navršenih 70 godina sam ukida smanjenje i mirovina raste.", hitno: false },
        { emoji: "💶", t: "Porez na mirovinu", opis: "Vlada predlaže ukidanje poreza na mirovine od 1.1.2027. Do tada umirovljenici plaćaju pola obračunatog poreza, a mirovine do 600 EUR ne plaćaju porez.", hitno: false },
      ],
    },
  ];

  const profil = profili.find(p => dob >= p.range[0] && dob <= p.range[1]) || profili[1];

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
        <RangeTrack value={dob} min={18} max={75} step={1} color={profil.color}
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
          ["Otvori ili aktiviraj III. stup (DPS = besplatni novac od države)", "Odmah ako nisi"],
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
// ─── POČNI OVDJE ─────────────────────────────────────────────────────────────
const KORACI_PO_DOBI = [
  { do: 34, naslov: "Imate vremena — to je vaša prednost", koraci: [
    { t: "Provjerite uplate u II. stup", d: "Na regos.hr ili u e-Građanima provjerite uplaćuje li poslodavac svaki mjesec 5 % vaše bruto plaće.", tab: "stup2" },
    { t: "Otvorite III. stup", d: "Uz 55,30 EUR mjesečno država dodaje 99,54 EUR godišnje. Kroz 30+ godina složena kamata napravi veliku razliku.", tab: "stup3" },
    { t: "Razmislite o kategoriji fonda", d: "Ako ne izaberete, REGOS vas stavlja u kategoriju B. S 30+ godina do mirovine kategorija A dugoročno nosi više.", tab: "edu" },
  ]},
  { do: 49, naslov: "Sredina karijere — vrijeme za provjeru", koraci: [
    { t: "Pogledajte koliko vam nedostaje", d: "Planer pokazuje razliku između očekivane i željene mirovine i koliko štednje je potrebno.", tab: "planer" },
    { t: "Iskoristite III. stup i poslodavca", d: "Poslodavac vam može uplaćivati do 804 EUR godišnje u III. stup bez poreza i doprinosa.", tab: "stup3" },
    { t: "Provjerite staž", d: "U e-Građanima pogledajte elektronički zapis o stažu i javite HZMO-u ako nešto nedostaje.", tab: "vodic" },
  ]},
  { do: 57, naslov: "Bliže mirovini — pripremite odluke", koraci: [
    { t: "Zatražite informativni izračun", d: "Na mojamirovina.hr ili u e-Građanima vidjet ćete stvarne podatke o plaćama i stažu.", tab: "vodic" },
    { t: "Samo I. stup ili I. + II.?", d: "Isprobajte usporedbu sa stanjem svog računa u II. stupu. To je najvažnija odluka pri umirovljenju.", tab: "ukupno" },
    { t: "Smanjite rizik u II. stupu", d: "Kad do mirovine ostane manje od 10 godina, REGOS vas prebacuje iz kategorije A u B; u zadnjih 5 godina razmislite o C.", tab: "edu" },
  ]},
  { do: 200, naslov: "Pred mirovinom — vrijeme odluka", koraci: [
    { t: "Kada otići?", d: "Svaka godina ranije smanjuje mirovinu iz I. stupa za 2,4 % (do 70. godine), a svaka godina kasnije trajno je povećava za 5,4 %.", tab: "stup1" },
    { t: "Samo I. stup ili I. + II.?", d: "Usporedite obje mogućnosti; HZMO prije odluke šalje informativni izračun.", tab: "ukupno" },
    { t: "Vrsta mirovine u MOD-u i 20 % jednokratno", d: "Ako birate oba stupa: pojedinačna, obiteljska ili sa zajamčenim razdobljem. Odluka je trajna.", tab: "stup2" },
    { t: "Zahtjev HZMO-u na vrijeme", d: "Predajte zahtjev barem 3 mjeseca prije željenog datuma umirovljenja.", tab: "vodic" },
  ]},
];

function PocniOvdjeView({ onNavigate }) {
  const [godRodj, setGodRodj] = useState(1975);
  const [spol, setSpol] = useState("M");
  const [staz, setStaz] = useState(25);
  const [placa, setPlaca] = useState(PROSJ_BRUTO);
  const [uDrugom, setUDrugom] = useState(null);

  const dob = TEKUCA_GODINA - godRodj;
  const uII = uDrugom ?? godRodj >= 1962;
  const starDob = dobZaStarosnu(spol, godRodj);
  const prij = dobZaPrijevremenu(spol, godRodj);
  const godDo = Math.max(0, starDob - dob);
  const stazNaKraju = staz + godDo;
  const godDoPrij = Math.max(0, prij.dob - dob);
  const stazPrij = staz + godDoPrij;
  const prijMoguca = stazPrij >= prij.staz;
  const dugogodisnji = staz + Math.max(0, 60 - dob) >= 41;

  const pvb = placa / PROSJ_BRUTO;
  const godinaMir = godRodj + Math.ceil(starDob);
  const godII = uII ? Math.min(stazNaKraju, godinaMir - 2002) : 0;
  const rentaII = uII ? rentaDozivotna(calcStup2FV({ bruto_placa: placa, god_stednje: godII, neto_prinos: 0.03 }) * 0.8) : 0;
  const iz = calcIzbor({ pvb, staz: stazNaKraju, stazOd2002: godII, rentaII });
  const procjena = uII ? Math.max(iz.samoIUkupno, iz.kombinirana) : iz.samoIUkupno;
  const neto26 = calcNetoMirovina(procjena).neto;
  const neto27 = calcNetoMirovina(procjena, { godina: 2027 }).neto;
  const skupina = KORACI_PO_DOBI.find(k => dob <= k.do);

  const kartica = { background: C.white, borderRadius: 14, padding: 20, boxShadow: "0 2px 16px rgba(26,46,82,0.07)", border: `1px solid ${C.border}` };
  const link = (tab, tekst = "Detaljnije →") => (
    <button onClick={() => onNavigate(tab)} style={{ marginTop: 10, background: "none", border: "none", color: C.blue, fontWeight: 600, fontSize: 12, padding: 0, cursor: "pointer" }}>{tekst}</button>
  );

  return (
    <div className="fadeIn">
      <div style={{ marginBottom: 18 }}>
        <h1 style={{ fontSize: 26, marginBottom: 6 }}>Vaša mirovina u 3 koraka</h1>
        <p style={{ color: C.slate, fontSize: 14, lineHeight: 1.6, maxWidth: 720 }}>
          Unesite nekoliko podataka i odmah vidite <strong>kada</strong> možete u mirovinu, <strong>okvirno koliko</strong> ćete
          primati i <strong>što</strong> vam je sada pametno napraviti. Ništa se ne sprema i nije potrebna prijava.
        </p>
      </div>

      <div style={{ ...kartica, marginBottom: 18 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: "4px 24px" }}>
          <Slider label={`Godina rođenja (sada ${dob} god.)`} value={godRodj} min={1950} max={2008} step={1} format={v => `${v}.`} onChange={setGodRodj} />
          <Slider label="Godine staža do danas" value={staz} min={0} max={Math.max(0, Math.min(50, dob - 15))} step={1} format={v => `${v} god.`} onChange={setStaz} />
          <Slider label="Bruto plaća (mjesečno)" value={placa} min={MIN_PLACA} max={6000} step={50} format={v => `${fmt(v, 0)} EUR`} onChange={setPlaca} />
          <div>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.navy, marginBottom: 6 }}>Spol</div>
            <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
              {[["M", "Muškarac"], ["Ž", "Žena"]].map(([v, l]) => (
                <button key={v} onClick={() => setSpol(v)} style={{ flex: 1, padding: "7px", borderRadius: 8, border: `2px solid ${spol === v ? C.navy : C.border}`,
                  background: spol === v ? C.navy : C.white, color: spol === v ? "#fff" : C.navy, fontSize: 12, fontWeight: 600 }}>{l}</button>
              ))}
            </div>
            <label style={{ fontSize: 12, color: C.slate, display: "flex", gap: 6, alignItems: "center" }}>
              <input type="checkbox" checked={uII} onChange={e => setUDrugom(e.target.checked)} />
              Uplaćujem i u II. stup (mirovinski fond)
            </label>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 16, marginBottom: 18 }}>
        <div style={{ ...kartica, borderTop: `4px solid ${C.navy}` }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: C.navy, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>1 · Kada mogu u mirovinu</div>
          <div style={{ fontSize: 13, color: C.slate, lineHeight: 1.8 }}>
            <div><strong>Starosna:</strong> s {fmtDob(starDob)} {godDo > 0 ? `(oko ${godinaMir}.)` : "— već ispunjavate dob"}, uz najmanje 15 godina staža.</div>
            <div><strong>Prijevremena:</strong> s {fmtDob(prij.dob)} uz {fmtDob(prij.staz)} staža
              {" "}— {prijMoguca ? <span style={{ color: C.green, fontWeight: 600 }}>po vašim podacima moguće</span> : <span style={{ color: C.red, fontWeight: 600 }}>nemate dovoljno staža</span>}.</div>
            {dugogodisnji && <div><strong>Dugogodišnji osiguranik:</strong> s 60 godina i 41 godinom staža, bez umanjenja.</div>}
          </div>
          <div style={{ marginTop: 10, padding: "8px 12px", background: C.bg, borderRadius: 8, fontSize: 12, color: C.slate, lineHeight: 1.6 }}>
            Svaka godina ranije: <strong style={{ color: C.red }}>−2,4 %</strong> (do 70. godine) · svaka godina kasnije: <strong style={{ color: C.green }}>+5,4 %</strong> trajno.
          </div>
          {link("stup1", "Isprobajte odlazak ranije ili kasnije →")}
        </div>

        <div style={{ ...kartica, borderTop: `4px solid ${C.green}` }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: C.green, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>2 · Okvirno koliko</div>
          <div style={{ fontSize: 30, fontWeight: 800, color: C.navy, fontFamily: "'Playfair Display', serif" }}>{fmt(procjena, 0)} EUR</div>
          <div style={{ fontSize: 12, color: C.slateL, marginBottom: 8 }}>bruto mjesečno, u današnjem novcu, uz {stazNaKraju} god. staža</div>
          <div style={{ fontSize: 13, color: C.slate, lineHeight: 1.7 }}>
            Neto: <strong>{fmt(neto26, 0)} EUR</strong> po pravilima za 2026. · <strong>{fmt(neto27, 0)} EUR</strong> ako se od 2027. ukine porez na mirovine.
          </div>
          {uII && (
            <div style={{ marginTop: 8, fontSize: 12, color: C.slate, lineHeight: 1.6 }}>
              Samo I. stup: {fmt(iz.samoIUkupno, 0)} EUR · I. + II. stup: {fmt(iz.kombinirana, 0)} EUR.
            </div>
          )}
          <div style={{ marginTop: 8, fontSize: 10, color: C.slateL, lineHeight: 1.5 }}>
            Pretpostavka: plaća u istom omjeru prema prosjeku RH kroz cijeli radni vijek. Stvarni iznos ovisi o plaćama iz cijele karijere.
          </div>
          {link(uII ? "ukupno" : "stup1")}
        </div>

        <div style={{ ...kartica, borderTop: `4px solid ${C.gold}` }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: C.gold, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>3 · Što napraviti sada</div>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.navy, marginBottom: 8 }}>{skupina.naslov}</div>
          {skupina.koraci.map((k, i) => (
            <div key={i} style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: C.navy }}>{i + 1}. {k.t}</div>
              <div style={{ fontSize: 12, color: C.slate, lineHeight: 1.55 }}>{k.d}{" "}
                <button onClick={() => onNavigate(k.tab)} style={{ background: "none", border: "none", color: C.blue, fontSize: 12, padding: 0, cursor: "pointer" }}>→</button>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ ...kartica, background: C.navy, color: "#fff" }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.goldL, marginBottom: 8 }}>Kako funkcionira mirovina u Hrvatskoj — u tri rečenice</div>
        <div style={{ fontSize: 13, lineHeight: 1.75, color: "rgba(255,255,255,0.85)" }}>
          <strong>I. stup</strong> (HZMO): od plaće se uplaćuje 15 % (ili 20 % ako niste u II. stupu); mirovina ovisi o godinama staža i o tome koliko ste
          zarađivali u odnosu na prosjek, a država je usklađuje dvaput godišnje.{" "}
          <strong>II. stup</strong>: 5 % plaće ide na vaš osobni račun u mirovinskom fondu i pri umirovljenju postaje dodatna doživotna mirovina.{" "}
          <strong>III. stup</strong>: dobrovoljna štednja uz državni poticaj od 15 %.
        </div>
        <button onClick={() => onNavigate("edu")} style={{ marginTop: 12, background: C.gold, color: C.navy, border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>
          Nauči osnove sustava →
        </button>
      </div>
    </div>
  );
}

export default function App() {
  const [user, setUser] = useState(null);         // null = guest
  const [tab, setTabState] = useState("pocni");
  const setTab = (t) => { setTabState(t); window.scrollTo({ top: 0, behavior: "smooth" }); };
  const [authModal, setAuthModal] = useState(null); // null | reason string
  const [feedbackModal, setFeedbackModal] = useState(false);
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
    { id: "pocni",    icon: <ChevronRight size={15} />, label: "Počni ovdje" },
    { id: "stup1",    icon: <Calculator size={15} />, label: "I. stup" },
    { id: "stup2",    icon: <PiggyBank size={15} />,  label: "II. stup" },
    { id: "stup3",    icon: <TrendingUp size={15} />, label: "III. stup" },
    { id: "invalid",  icon: <Shield size={15} />,     label: "Invalidska i obiteljska" },
    { id: "planer",   icon: <DollarSign size={15} />, label: "Koliko mi nedostaje" },
    { id: "ukupno",   icon: <BarChart2 size={15} />,  label: "Samo I. ili I. + II." },
    { id: "vodic",    icon: <Clock size={15} />,      label: "Vodič po dobi" },
    { id: "edu",      icon: <BookOpen size={15} />,   label: "Osnove sustava" },
    { id: "blog",     icon: <Star size={15} />,       label: "Blog" },
    { id: "svijet",   icon: <Users size={15} />,      label: "Svijet" },
    { id: "scenariji",icon: <FileText size={15} />,   label: "Moji scenariji" },
  ];
  const grupe = [
    { id: "pocni", label: "Počni ovdje", tabs: ["pocni"] },
    { id: "izracuni", label: "Izračuni", tabs: ["stup1", "stup2", "stup3", "invalid", "planer"] },
    { id: "odluke", label: "Odluke pred mirovinu", tabs: ["ukupno", "vodic"] },
    { id: "nauci", label: "Nauči", tabs: ["edu", "blog", "svijet"] },
    { id: "scenariji", label: "Moji scenariji", tabs: ["scenariji"] },
  ];
  const aktivnaGrupa = grupe.find(g => g.tabs.includes(tab)) || grupe[0];

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

      {/* Feedback modal */}
      {feedbackModal && <FeedbackModal onClose={() => setFeedbackModal(false)} />}

      {/* Floating feedback button */}
      <button onClick={() => setFeedbackModal(true)} title="Prijavi grešku ili komentar" style={{
        position: "fixed", bottom: 24, right: 24, zIndex: 100,
        width: 52, height: 52, borderRadius: "50%", border: "none",
        background: C.navy, color: "#fff", fontSize: 22, cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center",
        boxShadow: "0 6px 20px rgba(26,46,82,0.35)",
      }}>
        💬
      </button>

      {/* NAV */}
      <nav className="main-nav" style={{
        background: C.navy, color: "#fff", padding: "0 24px",
        display: "flex", alignItems: "center", gap: 0, flexWrap: "wrap",
        boxShadow: "0 2px 16px rgba(0,0,0,0.18)", position: "sticky", top: 0, zIndex: 50,
        minHeight: 58,
      }}>
        <button onClick={() => setTab("pocni")} style={{ background: "none", border: "none", fontFamily: "'Playfair Display', serif", fontSize: 18, fontWeight: 700, color: "#fff", marginRight: 24, whiteSpace: "nowrap", padding: "14px 0" }}>
          🏛 Mirovinski Sustav HR
        </button>
        <div style={{ display: "flex", gap: 2, flex: 1, flexWrap: "wrap" }}>
          {grupe.map(g => {
            const aktivna = g.id === aktivnaGrupa.id;
            return (
              <button key={g.id} className="grp-btn" onClick={() => setTab(g.tabs[0])} style={{
                padding: "0 14px", height: 50, border: "none", background: "none",
                color: aktivna ? C.goldL : "rgba(255,255,255,0.75)",
                fontWeight: aktivna ? 700 : 500, fontSize: 13,
                borderBottom: aktivna ? `2.5px solid ${C.gold}` : "2.5px solid transparent",
                transition: "all 0.15s", cursor: "pointer", whiteSpace: "nowrap",
              }}>
                {g.label}
              </button>
            );
          })}
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

      {aktivnaGrupa.tabs.length > 1 && (
        <div className="sub-nav" style={{ background: C.white, borderBottom: `1px solid ${C.border}`, padding: "8px 24px", display: "flex", gap: 6, flexWrap: "wrap", position: "sticky", top: 58, zIndex: 40 }}>
          {aktivnaGrupa.tabs.map(id => {
            const t = tabs.find(x => x.id === id);
            return (
              <button key={id} onClick={() => setTab(id)} style={{
                padding: "6px 14px", borderRadius: 20, border: `1.5px solid ${tab === id ? C.navy : C.border}`,
                background: tab === id ? C.navy : C.white, color: tab === id ? "#fff" : C.navy,
                fontSize: 12, fontWeight: 600, display: "flex", alignItems: "center", gap: 6, cursor: "pointer",
              }}>
                {t.icon} {t.label}
              </button>
            );
          })}
        </div>
      )}

      {/* AVM Banner */}
      <div className="data-banner" style={{ background: `linear-gradient(135deg, ${C.gold}cc, ${C.goldL}aa)`, padding: "8px 24px", fontSize: 12, color: C.navy, display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
        <strong>📊 Aktualni podaci (HZMO):</strong>
        <span>AVM od {AVM_DATUM} = <strong>{fmt(AVM)} EUR</strong></span>
        <span>|</span>
        <span>Doprinos II. stupa = <strong>5% bruto</strong></span>
        <span>|</span>
        <span>Najniža mirovina = <strong>{fmt(AVM * MIN_MIROVINA_FAKTOR)} EUR × godine staža</strong></span>
        <span>|</span>
        <span>Prosječna bruto plaća (2025.) = <strong>{fmt(PROSJ_BRUTO, 0)} EUR</strong></span>
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
        {tab === "pocni" && <PocniOvdjeView onNavigate={setTab} />}
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
        <strong style={{ color: "rgba(255,255,255,0.8)" }}>Mirovinski Sustav HR</strong> — isključivo informativna platforma<br />
        Izvor: HZMO (mirovinsko.hr) · HANFA · mss.gov.hr · ZOMO · ZOMF · AVM {fmt(AVM)} EUR od {AVM_DATUM} · Podaci ažurirani: {PODACI_AZURIRANI}<br />
        ⚠ Kalkulatori su aproksimativni. Za točan izračun koristite mojamirovina.hr ili se obratite HZMO-u.
      </footer>
    </div>
  );
}
