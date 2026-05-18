# DEK IT Zadání – ITA Generator

Nástroj pro automatické generování typizovaných IT zadání (ITA dokumentů) pro vývojáře systému Agenda ve Stavebniny DEK.

## Struktura repozitáře

```
dek_it_zadani/
├── SKILL.md              ← Instrukce pro Claude (skill)
├── README.md             ← Tento soubor
├── assets/
│   └── ITA-vzor.docx     ← Vzorový dokument (základ pro generování)
└── scripts/
    └── build_ita.py      ← Python skript pro generování .docx
```

## Jak to použít

### V Claude.ai (Projekt)

1. Otevřete sdílený Claude projekt „DEK IT Zadání"
2. Přiložte přepis schůzky nebo popište workflow (text, .docx, .txt)
3. Přiložte screenshoty (volitelné)
4. Napište: **„Vytvoř zadání pro IT"**

Claude se zeptá na doplňující informace a vygeneruje `.docx` dokument.

### Přímé použití skriptu

```bash
python3 scripts/build_ita.py --data moje_data.json --output ITA-muj-projekt.docx
```

Formát `data.json` je popsán na konci souboru `scripts/build_ita.py`.

## Aktualizace vzoru

Pokud se změní požadovaný formát ITA dokumentu:
1. Aktualizujte `assets/ITA-vzor.docx`
2. Commitněte změnu do gitu
3. Změna se automaticky projeví při příštím generování

## Závislosti skriptu

Skript využívá nástroje z `claude.ai` prostředí (`/mnt/skills/public/docx/scripts/`).
Při přímém spuštění mimo Claude je potřeba mít tyto skripty dostupné.
