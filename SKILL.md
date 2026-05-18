---
name: ita-generator
description: >
  Generuje typizované ITA zadání pro IT vývojáře systému Agenda (ERP/CRM Stavebniny DEK).
  Použij tento skill vždy, když uživatel napíše "Vytvoř zadání pro IT", "Vytvoř ITA",
  "Napiš IT zadání", "Vytvoř zadání pro developery" nebo cokoliv podobného.
  Uživatel obvykle přiloží textový přepis schůzky nebo popis workflow a screenshoty.
  Skill řídí celý proces: analýzu vstupů, kladení doplňujících otázek a finální generování
  .docx souboru ve správném formátu pomocí skriptu scripts/build_ita.py.
---

# ITA Generator – Generátor IT zadání pro DEK

## Co tento skill dělá

Přijme nestrukturovaný popis (přepis schůzky, workflow, screenshoty) a vygeneruje
typizovaný ITA dokument ve formátu .docx, který vývojáři Agendy akceptují jako zadání.

Vzorový dokument: `assets/ITA-vzor.docx`
Generovací skript: `scripts/build_ita.py`

---

## Krok 1 – Analýza vstupů

Po obdržení vstupu od uživatele proveď tiché zpracování:

- Přečti veškeré přiložené dokumenty (přepisy, .docx, .txt)
- Prohlédni přiložené screenshoty a popiš co na nich vidíš
- Identifikuj: systém/modul, popis změny, kdo je žadatel, co se má stát

Nikdy nezačínej generovat dokument bez položení doplňujících otázek (Krok 2).

---

## Krok 2 – Doplňující otázky

Před generováním se VŽDY zeptej na položky, které nejsou jasné ze vstupů.
Pokládej pouze otázky na chybějící nebo nejasné informace — neptej se na věci,
které ze vstupu jasně vyplývají.

### Povinné položky hlavičky (pokud chybí):
- **Oblast vývoje** – např. "CRM / Modul prodej", "Cenotvorba", "Sklad"
- **Název projektu** – stručný název změny
- **Žadatel** – jméno osoby, která zadání podává
- **Novinky (zveřejnit)** – ano / ne (zda se změna komunikuje zákazníkům)
- **Odkaz do Freela** – URL úkolu ve Freelo (může být "doplní žadatel")
- **Nutný termín zveřejnění** – pokud je, jinak prázdné

### Věcné otázky (pokud nejsou jasné ze vstupu):
- Co přesně je stávající stav a co se mění?
- Jaké jsou hranice změny – co do zadání NEPATŘÍ?
- Jsou negativní scénáře (co se stane, když podmínka není splněna)?
- Přenáší se nějaká data nebo vazby do navazujících dokladů/záznamů?

### Pravidlo:
Pokud je odpověď zřejmá z přiloženého přepisu nebo screenshotů, nepokládej tu otázku.
Raději uveď co jsi odvodil a zeptej se jen na potvrzení: "Ze záznamu jsem pochopil X — je to správně?"

---

## Krok 3 – Generování dokumentu

Po odsouhlasení všech informací spusť skript:

```bash
cd <cesta ke skill složce>
python3 scripts/build_ita.py --output /mnt/user-data/outputs/ITA-<nazev>.docx
```

Skript přijme data jako JSON přes stdin nebo jako argumenty — viz komentáře v build_ita.py.

Po vygenerování zavolej `present_files` s cestou k výstupu.

---

## Struktura ITA dokumentu

### Hlavička (prostý text s tabulátory, přesně jako ve vzoru)
```
Verze z DD.M.RRRR
Oblast vývoje:      <hodnota>
Novinky (zveřejnit): <ano/ne>
Podniky/holding:    Stavebniny DEK CZ
Externí test:       ano
Žadatel:            <jméno>
Nutný termín zveřejnění: <hodnota nebo prázdné>
Název projektu:     <název>
Odkaz do Freela:    <URL>
```

### Sekce dokumentu (v tomto pořadí):
1. **Byznys cíl a přínos** – Proč to děláme? Stávající stav + co chybí.
2. **User Story** – Pohledem koncového uživatele, první osoba.
3. **Funkční požadavky** – F1, F2, F3... Každý má bold nadpis + odstavce + odrážky.
4. **Popis workflow** – Scénáře jako číslovaný seznam: `1. [Kde] Co se děje`
5. **Akceptační kritéria** – Skupiny (Skupina 1, 2...), každé AC jako odrážkový seznam:
   - **AC X.Y: Název**
   - **Předpoklad:** text
   - **Akce:** text  *(pokud relevantní)*
   - **Výsledek:** text
   - **A zároveň:** text  *(pokud relevantní, může být víc)*
6. **Reference – screenshoty** – Bold název souboru + pomlčka + popis

### Pravidla formátování:
- **Žádné tabulky** – systém přebírá data z dokumentu do poznámkového bloku
- **Žádné podbarvení** ani barevné oddělovače
- Skupinové nadpisy v AC jsou tučné, ale nejsou odrážka (prostý odstavec)
- Funkční požadavky (F1, F2...) jsou tučné, ale nejsou nadpis stylem Nzev
- Workflow scénáře jsou číslované, každý scénář začíná od 1

---

## Tipy pro analýzu přepisů schůzek

Přepisy jsou často nestrukturované a opakující se. Hledej:
- Co uživatel **dnes dělá** (stávající stav)
- Co uživatel **chce dělat** (nová funkcionalita)
- **Technické detaily** – API, příznak položky, vazba, trigger
- **Hraniční případy** – "co když...", "pokud ne...", "ale pozor na..."
- **Co se nemá změnit** – ručně vložené položky, jiné doklady

Screenshoty analyzuj a uveď co vidíš — pomáhají ověřit workflow.

---

## Příklad doplňujících otázek (vzorový výstup Kroku 2)

```
Ze záznamu jsem pochopil toto:
- Systém: Agenda, modul prodej
- Změna: při vytvoření Nabídky/Objednávky z Poptávky se automaticky otevře popup NabídkoBota
- Vložení položek probíhá přes API

Před generováním potřebuji doplnit:
1. Oblast vývoje – je to "CRM / Modul prodej" nebo jinak?
2. Odkaz do Freela – máš už úkol založený?
3. Novinky (zveřejnit) – bude se tato změna komunikovat zákazníkům (ano/ne)?
```
