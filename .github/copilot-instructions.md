# LernKompass – Projektanweisungen für GitHub Copilot

Dieses Repository erzeugt aus einem **Zeitplan** (Excel) und einem **Modulplan**
(Excel/Word) einen fertigen **Unterrichtsplan** (Excel).

## Wichtigste Regel: genau EIN Befehl

Wenn der Nutzer einen Unterrichtsplan erstellen, generieren oder berechnen
möchte – oder einen Zeitplan + Modulplan bereitstellt – führe **einen einzigen**
Terminal-Befehl aus und sonst nichts:

```
python scripts/generate_unterrichtsplan.py --zeitplan "<Zeitplan.xlsx>" --modulplan "<Modulplan.xlsx>" --kursname "<Kursname>"
```

- **Implementiere die Logik NICHT in Python/Chat neu.** Das Skript liest beide
  Excel-Dateien und schreibt die Ausgabe selbst. Rufe ausschließlich das
  vorhandene Skript auf.
- **Zerlege die Aufgabe NICHT in mehrere Tool-Schritte** (kein manuelles
  Einlesen, kein schrittweises Bauen). Jeder zusätzliche Tool-Aufruf löst eine
  weitere Erlaubnisabfrage aus – das soll vermieden werden.
- Verwende **keine** Befehlsverkettung (`cd … ; python …`). Setze das
  Arbeitsverzeichnis voraus und rufe `python …` direkt auf.
- Melde nach Erfolg nur den **Ausgabepfad** (steht in der Zeile `[INFO] Saved:`).

## Verhalten des Skripts (nicht nachbauen, nur kennen)

- Der **Zeitplan ist die einzige Wahrheitsquelle** für alle Tagestypen
  (Zahl = Unterrichtstag, `Fe`, `Pr`, `Pr/4`→`Pr+`, `Prüf`, `Ft`, leer = WE/Ft).
- Bei einem **Defizit** an Unterrichtstagen kürzt das Skript **automatisch nur
  die Prüfungsvorbereitung** – niemals die Fachpraktische Begleitung (FPB).
  Es fragt dabei **nicht** nach.
- Die Anzahl der Module ist **variabel**; Modul-IDs werden automatisch erkannt
  und eingefärbt.

## Standardwerte

- `--ue-pro-tag` ist optional (Standard: `9`).
- `--output` ist optional (Standard: neben der Zeitplan-Datei,
  `Unterrichtsplan_<Kursname>_<Jahr>.xlsx`).
- Voraussetzung: `pip install openpyxl pandas` (einmalig).
