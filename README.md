# 🏛 Mirovina HR — Kalkulator i edukacijska platforma

Interaktivna web aplikacija za izračun i planiranje mirovine prema hrvatskom mirovinskom sustavu.

## Što sadrži

- **I. stup** — kalkulator s PVB, staž, polazni faktor, AVM
- **II. stup** — akumulacija OMF, vrste isplate, inovativne varijante
- **III. stup** — DPS kalkulator, prednosti dobrovoljne štednje
- **Ukupno** — kombinirana mirovina sva tri stupa
- **Invalidska i obiteljska mirovina**
- **Planer** — "Koliko mi nedostaje?"
- **Usporedba svjetskih sustava** — NL, DK, IS, AU, SG
- **Edukacija** — interaktivni vodiči za svaki parametar
- **Praktični vodič** — što učiniti po životnoj dobi (18–64)
- **Blog i AI** — 5 članaka + AI stručnjak (Claude API)
- **Scenariji** — spremi i izvozi u Excel

## Pokretanje lokalno

```bash
npm install
npm run dev
```

Otvori http://localhost:5173

## Build za produkciju

```bash
npm run build
```

Generirani `dist/` folder uploadaj na bilo koji statički hosting.

## Deploy na Vercel (preporučeno)

1. Puši projekt na GitHub
2. Idi na https://vercel.com/new
3. Importiraj repozitorij
4. Klikni **Deploy** — gotovo!

Vercel automatski detektira Vite i zna kako buildati.

## Tehnologije

- React 18 + Vite
- Recharts (grafovi)
- Lucide React (ikone)
- SheetJS / xlsx (Excel izvoz)
- Claude API (AI stručnjak u Blog tabu)

## Izvori podataka

- HZMO (mirovinsko.hr) — AVM, ZOMO, formule
- HANFA (hanfa.hr) — OMF, MOD, naknade
- mss.gov.hr — prava osiguranika, nasljeđivanje
- Mercer CFA Global Pension Index 2025
- OECD Pensions at a Glance 2025

## Napomena

Kalkulator je isključivo informativne prirode. Za točan izračun koristite mojamirovina.hr putem e-Građani.
