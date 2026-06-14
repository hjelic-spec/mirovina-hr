# Mirovina HR — Projektni kontekst za Claude Code

## Što je ovaj projekt

Web aplikacija za izračun, edukaciju i planiranje mirovine prema hrvatskom mirovinskom sustavu.
Građena u jednoj claude.ai sesiji — ovaj fajl prenosi cijeli kontekst u Claude Code.

Cilj: **najkompletniji nezavisni kalkulator mirovine u Hrvatskoj**, edukacijski bolji od
`mojamirovina.hr` (koji ima stvarne podatke, ali loš UX i nema edukaciju).

---

## Tehnički stack

```
React 18 + Vite (SPA, bez backend-a)
├── src/App.jsx          — cijela aplikacija (~4700 linija, 1 fajl)
├── src/main.jsx         — React entry point
├── index.html           — HTML shell
├── vite.config.js       — Vite konfiguracija s manual chunks
├── vercel.json          — SPA routing rewrite za Vercel deploy
└── package.json
```

**Dependencije:**
- `react` + `react-dom` ^18.3
- `recharts` ^2.12 — svi grafovi (Bar, Line, Area, Pie)
- `lucide-react` ^0.383 — ikone
- `xlsx` (SheetJS) ^0.18 — Excel export

**Nema:** TypeScript, CSS fajlova, state management librarija, backend, baze podataka.
Sve u jednom JSX fajlu. Storage: `window.storage` (Anthropic persistent storage API u claude.ai
artefaktima) + in-memory fallback za lokalni development.

---

## Arhitektura App.jsx

### Konstante i matematika (linije 17–73)
```js
const AVM = 14.45               // EUR, od 1.7.2025. (HZMO)
const MIN_MIROVINA_FAKTOR = 1.06 // najniža mirovina = 106% AVM × staž × PF × MF
const MAX_VB = 3.8               // maksimalni vrijednosni bodovi

calcStup1({ pvb, staz, pf, mf, avm })  // → { ob, bruto, najniza, najvisa, efektivna }
calcStup2FV({ bruto_placa, god_stednje, neto_prinos })  // → kapital (FV anuitetna renta)
rentaDozivotna(kapital, e65=18)         // → EUR/mj. (kapital / (e65 × 12))
```

### Slider komponente (linije 626–852)
Tri razine:
- `RangeTrack` — shared range input s CSS custom thumb i coloured fill
- `Slider` — lean verzija za Ukupno/brze inpute
- `MiniSlider` — za Edukacija panele
- `SmartParam` — bogati blok: icon + zone badge + RangeTrack + editable input + dynamicExplain + expandable explain

Svi slajderi imaju **direktan unos vrijednosti** (klik na value input → tipka broj → Enter).
CSS je u `SLIDER_CSS` string injektiran u `<style>` taga.

### Auth (linija 870)
`AuthModal` — overlay koji se pojavljuje ON DEMAND samo kad korisnik:
1. Klikne "Spremi" bez prijave → `onAuthRequired(reason)` callback
2. Ode na tab "Scenariji" bez prijave → CTA kartica
3. Klikne "Prijava" u navbaru

Nema auth gate na početku — kalkulator je slobodan bez registracije.
Auth podatke čuva `window.storage` (key: `user:${email}`, `session:current`).

### Tabovi i routing (linija 4543)
```
stup1    → Stup1Calculator
stup2    → Stup2Calculator
stup3    → Stup3View
ukupno   → UkupnoView
invalid  → InvalidskaView
planer   → PlanerView
svijet   → UsporedbaSustaviView
vodic    → VodicView
edu      → EduView
blog     → BlogView (uključuje SP500MirovinaBlog + AI Q&A)
scenariji→ MojiScenarijiView
```

### Excel export (linija 75–388)
`exportToExcel(scenariji, username)` — generira `.xlsx` s 6 sheetova:
1. Pregled svih scenarija (29 kolona, pokriva I/II/III stup)
2. I. stup detalji (samo ako postoje)
3. II. stup detalji
4. III. stup + DPS
5. Projekcija AVM (20 polugodišta 2026–2035)
6. Formule i pravni izvori

Gumb je u `MojiScenarijiView`, dostupan samo prijavljenim korisnicima s ≥1 scenarijem.

---

## Pravni i podatkovni okvir

### I. stup — formula (ZOMO, NN 96/25, 1.7.2025.)
```
M = PVB × Staž × PF × MF × AVM

PVB = prosj. vrijednosni bodovi (vaša plaća / prosj. plaća RH, prosjek karijere)
Staž = mirovinski staž u godinama
PF = polazni faktor:
     starosna: 1.0
     prijevremena: 1.0 - 0.002 × broj_mj_ranije  (min 0.7)
     odgoda: 1.0 + 0.0045 × broj_mj_odgode
MF = mirovinski faktor:
     starosna/prijevremena: 1.0
     invalidska potpuna (od 1.1.2026.): 1.1
     invalidska djelomična (od 1.1.2026.): 0.9
AVM = 14.45 EUR (od 1.7.2025.)

Najniža mirovina = 1.06 × AVM × Staž × PF × MF
Najviša mirovina = Staž × 3.8 × PF × MF × AVM
Efektivna = MAX(najniža, MIN(najviša, bruto))
Dodatak = 0.27 × osnovna (od 1.1.2024., NN 156/2023)
```

### II. stup — formula (ZOMF)
```
Doprinos = 5% bruto plaće → osobni račun u OMF
FV = PMT × [((1 + r_mj)^n_mj - 1) / r_mj]
    PMT = bruto × 0.05 / 12
    r_mj = neto_prinos / 12
    n_mj = godine × 12

Naknada OMF (HANFA): max 0.25% (2024) → 0.24% (2025) → 0.20% (od 2029.)
Jednokratna isplata: max 20% kapitala, uvjet: I. stup > najniža mirovina
Doživotna renta: kapital / (e65 × 12) — aproksimacija (MOD koristi aktuarske tablice)
```

### III. stup — DPS (ZDOBF, Ministarstvo financija)
```
DPS = 15% × god. uplate, max 99.54 EUR/god.
Ostvaruje se uplatom ≥ 663.61 EUR/god. (= 55.30 EUR/mj.)
Bez poreza ni u fazi štednje ni pri isplati
Dostupno od 55. god. života
Uplata poslodavca: do 804 EUR/god. neoporezivo
```

### S&P 500 referentni prinosi (Blog sekcija)
```
11.5% god. — nominalno, zadnjih 40 god. (Fidelity, prosinac 2025.)
10.2% god. — dugoročni CAGR od 1928. (NYU Stern / Damodaran)
~7.2% god.  — realno, prilagođeno inflacijom (procjena)
```

---

## Komponente — kratki opis

| Komponenta | Opis | Linija |
|---|---|---|
| `Stup1Calculator` | Kalkulator I. stupa: PVB, staž, PF, MF, AVM slajderi; SmartParam blokovi; bar chart usporedba; spremi | 973 |
| `Stup2Calculator` | Kalkulator II. stupa: plaća, god. štednje, prinos, jednokratna isplata; inovativne varijante (4 tipa rente); area chart rast kapitala | 1316 |
| `UkupnoView` | Kombinirani prikaz I.+27%+II., pie chart strukture, live slajderi | 1657 |
| `Stup3View` | III. stup + DPS kalkulator; SmartParam; prednosti vs. bankarska štednja | 3587 |
| `InvalidskaView` | Invalidska + obiteljska mirovina; vrsta (potpuna MF=1.1 / djelomična MF=0.9); obiteljska % po tipu | 3778 |
| `PlanerView` | "Koliko mi nedostaje?" — željena mirovina vs. projekcija; akcijski plan | 3935 |
| `UsporedbaSustaviView` | Mercer 2025 rang lista; 6 sustava (NL, DK, IS, AU, SG, HR); live usporedba rente | 4136 |
| `VodicView` | Praktični vodič 18–64 god.; 4 profila; prioritizirani koraci | 4364 |
| `EduView` | Edukacijski centar: 8 sekcija (pregled, staž, PVB, PF, AVM, II. stup, prava osiguranika, reforma); sticky sidebar nav; live interaktivni grafovi | 2448 |
| `SP500MirovinaBlog` | Misaoni eksperiment: 20% plaće → S&P 500 vs HR sustav; 4 opcije prinosa; usporedba; + prednosti/rizici bez preskriptivnog zaključka | 1960 |
| `BlogView` | 5 blog postova; filter po kategoriji; AI stručnjak (Claude API) | 2281 |
| `MojiScenarijiView` | Pregled scenarija; tip badge; ukupno box; Excel export gumb | 1759 |
| `AuthModal` | On-demand overlay; login + register; demo account | 870 |

---

## Paleta boja

```js
const C = {
  navy:   "#1a2e52",  // primarna — navbar, naslovi
  gold:   "#c9922b",  // akcent — AVM banner, CTA gumbi
  green:  "#2d7a4f",  // pozitivno — rezultati, III. stup
  red:    "#c0392b",  // upozorenja, prijevremena kazna
  blue:   "#2563eb",  // I. stup, informacije
  purple: "#7c3aed",  // II. stup, inovativne varijante
  amber:  "#d97706",  // napomene, neutral warning
  slate:  "#4a5568",  // tekst labela
}
```

---

## Što je pokriveno

- [x] I. stup kalkulator (sve vrste mirovina)
- [x] II. stup kalkulator (akumulacija + 4 tipa rente + inovativne varijante)
- [x] III. stup + DPS simulator
- [x] Ukupna kombinirana mirovina
- [x] Invalidska i obiteljska mirovina
- [x] Planer "koliko mi nedostaje"
- [x] Usporedba 6 svjetskih sustava (Mercer 2025)
- [x] Edukacija — 8 interaktivnih sekcija
- [x] Prava osiguranika u II. stupu (MOD, REGOS, nasljeđivanje, kategorije A/B/C)
- [x] Reforma sustava — argumenti za i protiv, bez preskriptivnog zaključka
- [x] S&P 500 vs HR sustav — misaoni eksperiment
- [x] Blog + AI stručnjak (Claude API)
- [x] Praktični vodič po životnoj dobi (18–64)
- [x] Auth (guest-first, login on demand)
- [x] Scenariji (save/load, persistent storage)
- [x] Excel export (6 sheetova: scenariji, I/II/III stup detalji, AVM projekcija, formule)
- [x] Izvoz projektnog ZIP-a s deployem na Vercel
- [x] Hrvatska jezična provjera (18 ispravaka)

## Što nije pokriveno (potencijalni sljedeći koraci)

- [ ] Rad uz mirovinu (nova pravila od 1.1.2026. — ispod 3.5h/dan)
- [ ] Mirovinska dob per spol — kalkulator koji uzima spol pri izračunu PF
- [ ] Posebni mirovinski propisi (vojska, policija, vatrogasci)
- [ ] Međunarodni mirovinski ugovori (30+ zemalja, HZMO baza)
- [ ] Porezni aspekti mirovina (koji dohotci se oporezuju i kako)
- [ ] Godišnji dodatak na mirovinu (novi institut, ZOMO 2025.)
- [ ] Mobilna optimizacija (trenutno desktop-first layout)
- [ ] PWA manifest za instalaciju na mobitel
- [ ] Real data import s mojamirovina.hr (zahtijeva REGOS API — nije javno dostupan)
- [ ] Vlastita domena (kandidati: `mirovinski-kalkulator.hr`, `mirovina-kalkulator.com`)

---

## Deploy

```bash
# Lokalni razvoj
npm install
npm run dev          # http://localhost:5173

# Build za produkciju
npm run build        # → dist/

# Vercel (preporučeno)
# 1. Push na GitHub
# 2. vercel.com → New Project → Import → Deploy
# vercel.json je već konfiguriran za SPA routing
```

---

## Napomene o window.storage

U claude.ai artefaktima `window.storage` je Anthropic persistent storage API.
**Lokalno ne postoji** — App.jsx ima try/catch fallback koji tiho pada na null:

```js
async function stGet(key) {
  try { const r = await window.storage.get(key); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
}
```

Za lokalni razvoj, zamijeni s `localStorage`:
```js
async function stGet(key) {
  try { return JSON.parse(localStorage.getItem(key)); } catch { return null; }
}
async function stSet(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); return true; } catch { return false; }
}
async function stDel(key) { try { localStorage.removeItem(key); } catch {} }
```

---

## Ključni pravni izvori

| Zakon / Izvor | Sadržaj | URL |
|---|---|---|
| ZOMO (NN 96/25, 1.7.2025.) | Osnovna mirovina, PF, MF, AVM, usklađivanje | mirovinsko.hr |
| ZOMF (NN 19/14 + izmjene) | II. stup, kategorije OMF, nasljeđivanje | mss.gov.hr |
| ZDOBF | III. stup, DPS | mfin.gov.hr |
| HANFA propis | Naknade OMF (2024–2029.) | hanfa.hr |
| NN 156/2023 | Dodatak 27% na mirovinu (od 1.1.2024.) | nn.hr |
| Mercer CFA Global Pension Index 2025 | Rang lista svjetskih sustava | mercer.com |
| OECD Pensions at a Glance 2025 | Stope zamjene po zemljama | oecd.org |
| Fidelity.com | S&P 500 prinos zadnjih 40 god. = 11.5% | fidelity.com |
| NYU Stern / Damodaran | S&P 500 dugoročni CAGR = 10.2% (1928–2024) | pages.stern.nyu.edu |
