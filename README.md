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
| `Pr+` | Praktikum + Fachpraktische Begleitung (FPB) | Gelb (wie `Pr`) |
| `Prüf` | Prüfung | Rot |
| `Ft` | Feiertag | Hellgrün (heller als `Fe`) |
| beliebige Modul-IDs | Unterrichtsmodule | je eigene Farbe (dynamisch) |

Die **Modulfarben werden dynamisch vergeben** — egal ob `M1`–`M18`, `M0`–`M16`
oder `IM0`–`IM3`: jede tatsächlich vorkommende Modul-ID bekommt automatisch eine
eigene Farbe aus der Palette. Die Anzahl der Module ist frei variabel.

Da die Farbzuordnung über bedingte Formatierung läuft, kann der Nutzer Zellwerte manuell ändern — die Farbe passt sich automatisch an. Legende und Kalender verwenden dieselbe Farbquelle und können nicht auseinanderlaufen.

#### Automatischer UE-Ausgleich

Der Modulplan ist **idealtypisch** (maximaler UE-Umfang). Bietet der konkrete
Zeitplan weniger Unterrichtstage als der Modulplan benötigt, kürzt LernKompass
**automatisch und ausschließlich die Prüfungsvorbereitung** – die
**Fachpraktische Begleitung (FPB) wird nie gekürzt**. Die Kürzung wird in der
Spalte *Bemerkungen* dokumentiert. Es erfolgt **keine interaktive Rückfrage**
(wichtig für den unbeaufsichtigten Betrieb unter Copilot).

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

## Verwendung mit GitHub Copilot (VS Code, Agent-Modus)

LernKompass läuft auch unter **GitHub Copilot Chat** – **ohne ständige
Erlaubnisabfragen**. Das Prinzip: die gesamte Logik steckt in *einem* Skript-Aufruf,
und ein einziger Befehl löst höchstens eine Abfrage aus (mit Freigabeliste: keine).

Mitgeliefert:

- [`.github/prompts/unterrichtsplan.prompt.md`](.github/prompts/unterrichtsplan.prompt.md) — Slash-Befehl `/unterrichtsplan`
- [`.github/copilot-instructions.md`](.github/copilot-instructions.md) — steuert auch freie Anfragen auf den einen Befehl
- [`.vscode/settings.json`](.vscode/settings.json) — gibt `python` ohne Rückfrage frei

Vollständige Einrichtung inkl. der wichtigen Unterscheidung **GitHub Copilot vs.
Microsoft 365 Copilot**: siehe **[docs/COPILOT-Setup.md](docs/COPILOT-Setup.md)**.

> Hinweis: **Microsoft 365 Copilot** (Word/Excel/Teams) kann *keine* lokalen
> Skripte ausführen – dafür wäre ein Neubau auf Power Platform nötig.

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
├── .github/
│   ├── copilot-instructions.md       # Repo-weite Copilot-Steuerung
│   └── prompts/
│       └── unterrichtsplan.prompt.md # Copilot-Skill: /unterrichtsplan
├── .vscode/
│   └── settings.json                 # Auto-Freigabe für python (keine Abfragen)
├── docs/
│   └── COPILOT-Setup.md              # Einrichtung Copilot ohne Erlaubnisabfragen
├── skills/
│   └── unterrichtsplan/
│       └── SKILL.md          # Claude Code Skill
├── scripts/
│   └── generate_unterrichtsplan.py   # Standalone Python-Skript (Engine)
└── beispiele/
    └── README.md             # Hinweise zu Beispieldateien
```

---

## Zukunft / Roadmap

- [x] GitHub-Copilot-Integration (Prompt + Auto-Freigabe, ohne Erlaubnisabfragen)
- [x] Automatische Kürzung der Prüfungsvorbereitung bei UE-Defizit
- [x] Variable Modulanzahl mit dynamischer Farbvergabe
- [ ] Modulplan als eigenständige Excel-Vorlage (`Modulplan_Template.xlsx`)
- [ ] Microsoft 365 Copilot / Power Platform (lokale Skripte dort nicht ausführbar → Neubau)
- [ ] Web-Interface (Power Apps oder einfaches HTML-Tool)

---

## Hintergrund

Entwickelt für **BBQ – Baumann Bildung und Qualifizierung GmbH** im Rahmen der Neuzertifizierung von Bildungsmaßnahmen.  
Der idealtypische Modulplan muss je nach Startdatum, Bundesland und Ferienlage flexibel angepasst werden — LernKompass automatisiert genau diesen Schritt.

---

*Projekt: [constellationbuilds/LernKompass](https://github.com/constellationbuilds/LernKompass)*
