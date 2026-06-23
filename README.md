# LernKompass

**Automatisierte Unterrichtsplanung für Bildungsmaßnahmen**

LernKompass erstellt aus zwei Eingabedateien einen vollständigen, formatierten Unterrichtsplan als Excel-Datei — inklusive farbcodiertem Jahreskalender.

---

## Funktionsweise

```
Zeitplan (Excel)  +  Modulplan (Excel/Word)
          ↓
    LernKompass
          ↓
  Unterrichtsplan (Excel)
  ├── Sheet 1: Unterrichtsplan (tabellarisch)
  ├── Sheet 2: Kalenderübersicht (monatlich)
  └── Sheet 3: Jahreskalender (farbcodiert, bedingte Formatierung)
```

### Eingaben

| Datei | Format | Inhalt |
|-------|--------|--------|
| Zeitplan | `.xlsx` | Jahreskalender mit Unterrichts-, Ferien-, Praktikums- und Prüfungstagen |
| Modulplan | `.xlsx` oder `.docx` | Modulliste mit Modul-Nr., Modulbezeichnung, UE |

### Ausgabe

Excel-Datei `Unterrichtsplan_[Kursname]_[Jahr].xlsx` mit:

- **Sheet 1 – Unterrichtsplan**: Modul | Bezeichnung | Beginn | Ende | Tage | UE | Bemerkungen
- **Sheet 2 – Kalenderübersicht**: Unterrichtstage und UE pro Monat, kumuliert
- **Sheet 3 – Jahreskalender**: Alle Jahre nebeneinander, tagesgenau befüllt und farbcodiert

#### Farbcodierung (bedingte Formatierung)

| Kürzel | Bedeutung | Farbe |
|--------|-----------|-------|
| `Fe` | Ferien | Grün |
| `Pr` | Praktikum | Gelb |
| `Prüf` | Prüfung | Rot |
| `Ft` | Feiertag | Rosa |
| `M1`–`M16` | Unterrichtsmodule | je eigene Farbe |
| `IM1`–`IM3`, `IM0` | Fachmodule | je eigene Farbe |

Da die Farbzuordnung über bedingte Formatierung läuft, kann der Nutzer Zellwerte manuell ändern — die Farbe passt sich automatisch an.

---

## Verwendung als Claude Code Skill

Der Skill liegt unter `skills/unterrichtsplan/SKILL.md`.  
Installation: Datei nach `~/.claude/skills/unterrichtsplan/SKILL.md` kopieren.

Aufruf in Claude Code:

```
/unterrichtsplan
```

Oder natürlichsprachlich: *„Erstelle einen Unterrichtsplan aus diesen zwei Dateien."*

---

## Direkte Nutzung (Python-Skript)

```bash
pip install openpyxl pandas
python scripts/generate_unterrichtsplan.py \
  --zeitplan "Zeitplan_Ort.xlsx" \
  --modulplan "Modulplan_Kurs.xlsx" \
  --kursname "KBM" \
  --ue-pro-tag 9
```

---

## Projektstruktur

```
LernKompass/
├── README.md
├── skills/
│   └── unterrichtsplan/
│       └── SKILL.md          # Claude Code Skill
├── scripts/
│   └── generate_unterrichtsplan.py   # Standalone Python-Skript
└── beispiele/
    └── README.md             # Hinweise zu Beispieldateien
```

---

## Zukunft / Roadmap

- [ ] Modulplan als eigenständige Excel-Vorlage (`Modulplan_Template.xlsx`)
- [ ] Copilot-Prompt für Microsoft 365 / Teams (2-Dokumente-Limit)
- [ ] Unterstützung mehrerer Bundesländer (automatische Feiertagserkennung)
- [ ] Web-Interface (Power Apps oder einfaches HTML-Tool)

---

## Hintergrund

Entwickelt für **BBQ – Baumann Bildung und Qualifizierung GmbH** im Rahmen der Neuzertifizierung von Bildungsmaßnahmen.  
Der idealtypische Modulplan muss je nach Startdatum, Bundesland und Ferienlage flexibel angepasst werden — LernKompass automatisiert genau diesen Schritt.

---

*Projekt: [constellationbuilds/LernKompass](https://github.com/constellationbuilds/LernKompass)*
