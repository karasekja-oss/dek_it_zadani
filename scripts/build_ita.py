#!/usr/bin/env python3
"""
build_ita.py – Generátor ITA dokumentu pro DEK IT zadání

Použití:
    python3 build_ita.py --data data.json --output vystup.docx
    python3 build_ita.py --data data.json  (výstup: ITA-<nazev_projektu>.docx)

Formát data.json: viz konec tohoto souboru (sekce DATA FORMAT)

Závislosti:
    pip install docx2python python-docx
    Nebo pro generování z XML: pouze stdlib + kopie vzoru

Poznámka: Skript edituje XML vzorového dokumentu přímo,
čímž zachovává přesné styly, fonty a numbering definice originálu.
Vzorový soubor: assets/ITA-vzor.docx
"""

import json
import os
import re
import shutil
import sys
import argparse
import subprocess
import tempfile
from pathlib import Path

# ── Cesty ────────────────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).parent
ASSETS_DIR = SCRIPT_DIR.parent / "assets"
VZOR_PATH = ASSETS_DIR / "ITA-vzor.docx"

# Cesta ke scriptům pro práci s docx (ze skills public)
SKILLS_SCRIPTS = Path("/mnt/skills/public/docx/scripts/office")

# ── Pomocné XML funkce ────────────────────────────────────────────────────────

_counter = [1]

def _pid():
    c = _counter[0]; _counter[0] += 1
    return f"{c * 17 + 100:08X}", f"{c * 13 + 200:08X}"

def _r(text, bold=False, italic=False):
    """Jeden XML run s Calibri fontem."""
    b = "<w:b/><w:bCs/>" if bold else ""
    i = "<w:i/><w:iCs/>" if italic else ""
    space = ' xml:space="preserve"' if text.startswith(" ") or text.endswith(" ") else ""
    # Escapování XML znaků
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    return f'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>{b}{i}</w:rPr><w:t{space}>{text}</w:t></w:r>'

def _p_base(ppr_inner, *runs):
    pa, ta = _pid()
    runs_xml = "".join(runs)
    return (
        f'<w:p w14:paraId="{pa}" w14:textId="{ta}" w:rsidR="00000001" w:rsidRDefault="00000001">'
        f"<w:pPr>{ppr_inner}</w:pPr>{runs_xml}</w:p>"
    )

def _font_rpr():
    return '<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>'

def p_verze(text):
    return _p_base('<w:pStyle w:val="Vrazncitt"/>', _r(text))

def p_tab_row(label, value):
    pa, ta = _pid()
    val_xml = f'<w:tab/><w:t xml:space="preserve">{value}</w:t>' if value else "<w:tab/>"
    label_esc = label.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return (
        f'<w:p w14:paraId="{pa}" w14:textId="{ta}" w:rsidR="00000001" w:rsidRDefault="00000001">'
        f'<w:pPr><w:tabs><w:tab w:val="left" w:pos="5730"/><w:tab w:val="left" w:pos="5760"/></w:tabs></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{label_esc} </w:t>{val_xml}</w:r></w:p>'
    )

def p_empty():
    pa, ta = _pid()
    return f'<w:p w14:paraId="{pa}" w14:textId="{ta}" w:rsidR="00000001" w:rsidRDefault="00000001"/>'

def p_nzev(text):
    """Nadpis stylem Nzev (Byznys cíl, User Story, Funkční požadavky...)"""
    return _p_base('<w:pStyle w:val="Nzev"/>', _r(text))

def p_body(*runs):
    """Tělo textu."""
    rpr = f'<w:rPr>{_font_rpr()}</w:rPr>'
    return _p_base(f'<w:spacing w:before="240" w:after="240"/>{rpr}', *runs)

def p_bold(text):
    """Tučný odstavec – skupinový nadpis nebo F1/F2 nadpis."""
    return p_body(_r(text, bold=True))

def p_bullet(numid, ilvl, *runs):
    """Odrážkový odstavec (AC položky, odrážky v F sekci)."""
    rpr = f'<w:rPr>{_font_rpr()}</w:rPr>'
    return _p_base(
        f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{numid}"/></w:numPr>'
        f'<w:spacing w:before="240" w:after="240"/>{rpr}',
        *runs
    )

def p_numbered(numid, kde, co):
    """Číslovaný krok workflow: 1. [Kde] Co se děje"""
    return p_bullet(numid, "0", _r(f"[{kde}] ", bold=True), _r(co))

def p_ac_title(numid, code, title):
    """AC nadpis: - AC X.Y: Název (tučně, odrážka level 0)"""
    return p_bullet(numid, "0", _r(f"{code}: {title}", bold=True))

def p_ac_line(numid, label, text):
    """AC řádek: - Předpoklad: text (label tučně, text normálně, level 1)"""
    return p_bullet(numid, "1", _r(label, bold=True), _r(f" {text}"))

# ── Sestavení těla dokumentu ──────────────────────────────────────────────────

def build_body(d):
    """
    d = slovník s daty dokumentu (viz DATA FORMAT níže)
    Vrací seznam XML odstavců.
    """
    ps = []

    # ── Hlavička ──
    ps.append(p_verze(f"Verze z {d['verze']}"))
    ps.append(p_tab_row("Oblast vývoje:", d.get("oblast_vyvoje", "")))
    ps.append(p_tab_row("Novinky (zveřejnit):", d.get("novinky", "")))
    ps.append(p_tab_row("Podniky/holding:", d.get("podnik", "Stavebniny DEK CZ")))
    ps.append(p_tab_row("Externí test:", d.get("externi_test", "ano")))
    ps.append(p_tab_row("Žadatel:", d.get("zadatel", "")))
    ps.append(p_tab_row("Nutný termín zveřejnění (například platnost nějakého zákona):", d.get("termin", "")))
    ps.append(p_tab_row("Název projektu:", d.get("nazev_projektu", "")))
    ps.append(p_tab_row("Odkaz do Freela:", d.get("freelo", "")))
    ps.append(p_empty()); ps.append(p_empty())

    # ── Byznys cíl ──
    ps.append(p_nzev("Byznys cíl a přínos (Proč to děláme?)"))
    for para in d.get("byznys_cil", []):
        ps.append(p_body(_r(para)))
    ps.append(p_empty())

    # ── User Story ──
    ps.append(p_nzev("User Story"))
    for para in d.get("user_story", []):
        ps.append(p_body(_r(para)))
    ps.append(p_empty())

    # ── Funkční požadavky ──
    ps.append(p_nzev("Funkční požadavky (Jak to má fungovat?)"))
    for f in d.get("funkcni_pozadavky", []):
        ps.append(p_bold(f["nadpis"]))  # např. "F1 – Popis"
        for para in f.get("odstavce", []):
            if isinstance(para, str):
                ps.append(p_body(_r(para)))
            elif isinstance(para, list):
                # seznam odrážek
                for item in para:
                    ps.append(p_bullet("38", "1", _r(item)))
        ps.append(p_empty())

    ps.append(p_empty())

    # ── Workflow ──
    ps.append(p_nzev("Popis workflow"))
    # Scénáře – každý má vlastní numId (42, 44, ...) pro restart číslování od 1
    scenar_numids = ["42", "44", "46", "48"]  # pro až 4 scénáře
    for i, scenar in enumerate(d.get("workflow", [])):
        ps.append(p_bold(scenar["nazev"]))
        ps.append(p_empty())
        numid = scenar_numids[i] if i < len(scenar_numids) else "42"
        for krok in scenar["kroky"]:
            ps.append(p_numbered(numid, krok["kde"], krok["co"]))
        ps.append(p_empty())

    ps.append(p_empty())

    # ── Akceptační kritéria ──
    ps.append(p_nzev("Akceptační kritéria"))
    # Skupiny AC – každá skupina dostane vlastní numId pro AC tituly (38, 39, 40...)
    ac_numids = ["38", "39", "40", "38", "39"]
    for i, skupina in enumerate(d.get("ac_skupiny", [])):
        ps.append(p_bold(skupina["nazev"]))  # "Skupina 1 – ..."
        numid = ac_numids[i % len(ac_numids)]
        for ac in skupina["ac"]:
            ps.append(p_ac_title(numid, ac["kod"], ac["nazev"]))
            for polozka in ac.get("polozky", []):
                # polozka = {"label": "Předpoklad", "text": "..."}
                ps.append(p_ac_line(numid, polozka["label"] + ":", polozka["text"]))
        ps.append(p_empty())

    ps.append(p_empty())

    # ── Reference ──
    ps.append(p_nzev("Reference – screenshoty"))
    if d.get("figma"):
        ps.append(p_body(_r("Prototyp Figma – ", bold=True), _r(d["figma"])))
    for ref in d.get("reference", []):
        ps.append(p_body(_r(ref["soubor"], bold=True), _r(f" – {ref['popis']}")))

    return ps


def build_xml(d, vzor_xml):
    """Sestaví finální document.xml z dat a vzorového XML."""
    paragraphs = build_body(d)

    # Přidáme numbering pro workflow scénáře (numId 44, 46, 48) pokud chybí
    # (numId 42 je už ve vzoru jako decimal)

    header = vzor_xml[: vzor_xml.index("<w:body>") + len("<w:body>")]
    footer = "</w:body>\n</w:document>"
    body_content = "\n    ".join(paragraphs)

    sectpr = """<w:sectPr w:rsidR="45F47F45">
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>"""

    return header + "\n    " + body_content + "\n    " + sectpr + "\n" + footer


def add_extra_numids(numbering_xml, count=3):
    """Přidá extra decimal numId (44, 46, 48) pro vícescénářový workflow."""
    # Zjistíme abstractNumId pro decimal (numId 42 -> abstractNum 0)
    m = re.search(r'<w:num w:numId="42".*?<w:abstractNumId w:val="(\d+)"', numbering_xml, re.DOTALL)
    if not m:
        return numbering_xml
    aid = m.group(1)

    new_nums = ""
    for i in range(count):
        nid = 44 + i * 2
        new_nums += f'  <w:num w:numId="{nid}">\n    <w:abstractNumId w:val="{aid}"/>\n  </w:num>\n'

    return numbering_xml.replace("</w:numbering>", new_nums + "</w:numbering>")


# ── Hlavní funkce ─────────────────────────────────────────────────────────────

def generate(data_dict, output_path):
    """Vygeneruje ITA .docx soubor."""

    if not VZOR_PATH.exists():
        raise FileNotFoundError(f"Vzorový soubor nenalezen: {VZOR_PATH}")

    unpack = SKILLS_SCRIPTS / "unpack.py"
    pack = SKILLS_SCRIPTS / "pack.py"

    if not unpack.exists():
        raise FileNotFoundError(f"Skripty pro docx nenalezeny: {SKILLS_SCRIPTS}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        unpacked = tmpdir / "unpacked"

        # Rozbalení vzoru
        result = subprocess.run(
            ["python3", str(unpack), str(VZOR_PATH), str(unpacked)],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise RuntimeError(f"Unpack selhal: {result.stderr}")

        # Úprava document.xml
        doc_xml_path = unpacked / "word" / "document.xml"
        vzor_xml = doc_xml_path.read_text(encoding="utf-8")
        new_xml = build_xml(data_dict, vzor_xml)
        doc_xml_path.write_text(new_xml, encoding="utf-8")

        # Úprava numbering.xml – přidání extra numId pro workflow
        num_xml_path = unpacked / "word" / "numbering.xml"
        num_xml = num_xml_path.read_text(encoding="utf-8")
        num_xml = add_extra_numids(num_xml)
        num_xml_path.write_text(num_xml, encoding="utf-8")

        # Zabalení
        result = subprocess.run(
            ["python3", str(pack), str(unpacked), str(output_path),
             "--original", str(VZOR_PATH)],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise RuntimeError(f"Pack selhal: {result.stderr}")

    print(f"✓ Dokument vygenerován: {output_path}")
    return output_path


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generátor ITA dokumentu")
    parser.add_argument("--data", required=True, help="Cesta k JSON souboru s daty")
    parser.add_argument("--output", help="Výstupní .docx soubor")
    args = parser.parse_args()

    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)

    if not args.output:
        nazev = data.get("nazev_projektu", "dokument").replace(" ", "_").replace("/", "-")[:40]
        args.output = f"ITA-{nazev}.docx"

    generate(data, args.output)


if __name__ == "__main__":
    main()


# ── DATA FORMAT ───────────────────────────────────────────────────────────────
"""
Příklad data.json:

{
  "verze": "18.5.2026",
  "oblast_vyvoje": "CRM / Modul prodej",
  "novinky": "ne",
  "podnik": "Stavebniny DEK CZ",
  "externi_test": "ano",
  "zadatel": "Jan Karásek",
  "termin": "",
  "nazev_projektu": "NabídkoBot – plná integrace do modulu prodej",
  "freelo": "https://app.freelo.io/task/12345",

  "byznys_cil": [
    "NabídkoBot je AI systém...",
    "Stávající stav: ..."
  ],

  "user_story": [
    "Jako obchodník...",
    "Dále chci mít možnost..."
  ],

  "funkcni_pozadavky": [
    {
      "nadpis": "F1 – Automatické otevření popup okna při vytvoření dokladu",
      "odstavce": [
        "Při kliknutí na tlačítko...",
        ["odrážka 1", "odrážka 2", "odrážka 3"],
        "Další odstavec..."
      ]
    },
    {
      "nadpis": "F2 – Vložení označených položek do dokladu přes API",
      "odstavce": [
        "Po kliknutí...",
        ["odrážka 1", "odrážka 2"]
      ]
    }
  ],

  "workflow": [
    {
      "nazev": "Scénář 1 – Vytvoření nabídky / objednávky přímo z přehledu poptávek",
      "kroky": [
        {"kde": "Přehled poptávek", "co": "Uživatel vidí seznam poptávek..."},
        {"kde": "Systém Agenda", "co": "Agenda vytvoří nový doklad..."}
      ]
    },
    {
      "nazev": "Scénář 2 – Opakované vyvolání NabídkoBota",
      "kroky": [
        {"kde": "Detail dokladu", "co": "Uživatel otevře existující Nabídku..."}
      ]
    }
  ],

  "ac_skupiny": [
    {
      "nazev": "Skupina 1 – Automatické otevření popup okna",
      "ac": [
        {
          "kod": "AC 1.1",
          "nazev": "Automatické otevření popup okna při vytvoření Nabídky",
          "polozky": [
            {"label": "Předpoklad", "text": "Uživatel je v přehledu poptávek..."},
            {"label": "Akce", "text": "Uživatel klikne na tlačítko Vytvořit Nabídku."},
            {"label": "Výsledek", "text": "Vytvoří se nový doklad Nabídka."},
            {"label": "A zároveň", "text": "Ihned se otevře popup okno..."}
          ]
        }
      ]
    }
  ],

  "figma": "https://www.figma.com/proto/...",

  "reference": [
    {
      "soubor": "pohled_na_seznam_poptavek.jpg",
      "popis": "Přehled poptávek v modulu prodej..."
    }
  ]
}
"""
