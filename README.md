# ABC Analysis Tool

Automatizovaný Python skript na spracovanie predajných dát a tvorbu ABC analýzy podľa obratu a zisku. Výstupom je prehľadný Excel report rozdelený podľa mesiacov a kvartálov, vrátane zoskupenia podľa dodávateľa.

---

## Funkcie

- Spracovanie Excel súboru s predajnými dátami
- Čistenie dát (odstránenie testov, darčekov, záporných ziskov atď.)
- Normalizácia produktových kódov podľa dodávateľa
- Výpočet:
  - Obratu
  - Zisku
  - Podielov
  - Kumulatívnych podielov
  - ABC klasifikácie (A/B/C)
- Export do Excelu:
  - mesačné tabuľky (`Jan`, `Jan_summary`, ...)
  - kvartálne tabuľky (`1Q`, `1Q_summary`, ...)
- Automatický názov výstupu podľa dátumu

---

## Ukážka výstupu

| Master kód | Dodávateľ   | Množstvo | Obrat  | Zisk | Obrat_cum  | ABC_obrat  |
|------------|-------------|----------|--------|------|------------|------------|
| 123-XYZ    | Dodavatel 3 | 180      | 1200 € | 450 €| 14.8 %     | A          |

_Súčasťou výstupu je aj súhrnná tabuľka po skupinách A/B/C._

---

## Poznámka k anonymizácii

Pre účely zverejnenia boli názvy skutočných dodávateľov nahradené neutrálnymi označeniami `Dodavatel 1`, `Dodavatel 2` atď.  
Logika úpravy produktových kódov (napr. rôzne štruktúry podľa dodávateľa) však **zostáva zachovaná**, aby bol skript plne funkčný a realistický.

---

## ▶️ Spustenie

### 1. Nainštaluj závislosti:

```bash
pip install pandas numpy openpyxl
```

### 2. Spusti skript:

```bash
python abc_analysis.py
```

(Skript používa natvrdo definovaný súbor. Prispôsob si názov súboru priamo v skripte.)

---

## Výstup

Vytvorí sa Excel súbor:
```
ABC_25_13.07.2025.xlsx
```

Obsahuje viacero sheetov – pre každý mesiac a kvartál samostatne:
- `Jan`, `Jan_summary`
- `1Q`, `1Q_summary`
- ...

---

## Tipy

- Skript si vieš upraviť aj na export do PDF alebo .csv
- Možné rozšíriť o CLI (napr. výber súboru cez `argparse`)
- Vhodné na použitie v controllingu, logistike, e-commerce, predajnej analýze

---

## Autor

Projekt vytvoril [Tadeáš Galbavý](https://github.com/galbavy) – dátový analytik so zameraním na e-commerce a automatizáciu reportingu.
