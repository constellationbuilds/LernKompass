---
name: unterrichtsplan
description: >
  Erstellt automatisch einen Unterrichtsplan als Excel-Datei aus zwei Eingaben:
  einem Zeitplan (Excel-Kalender mit Unterrichts-, Ferien- und Praktikumstagen)
  und einem Modulplan (Excel oder Word mit Modul-Nr, Modulbezeichnung, UE).
  Verwende diesen Skill immer wenn der Nutzer einen Unterrichtsplan erstellen,
  generieren oder berechnen möchte, "/unterrichtsplan" eingibt, Begriffe wie
  "Lehrplan", "Stundenplan", "Modulplan umsetzen" oder "Unterrichtsplanung"
  verwendet, oder zwei Dateien (Zeitplan + Modulplan) für eine Bildungsmaßnahme
  vorliegen.
---

# Unterrichtsplan-Generator

Dieser Skill erstellt einen vollständigen Unterrichtsplan als Excel-Datei.

## Schritt 1: Eingaben erfragen

Frage den Nutzer nach:
1. **Zeitplan-Datei** (Excel): Kalender mit Unterrichts-/Ferien-/Praktikumstagen
2. **Modulplan-Quelle**: Entweder eine eigenständige Excel/Word-Datei ODER eine Konzeptionsdatei (Word), aus der der Modulplan extrahiert wird
3. **Kursname** (optional, z.B. "KBM", "IT-Fachkraft", "u-IMMO") — wird im Dateinamen verwendet
4. **UE pro Tag** (optional, Standard: 9)

## Schritt 2: Zeitplan parsen

Der Zeitplan hat Jahresblätter (z.B. "1. Jahr", "2. Jahr", "3. Jahr").
Jedes Blatt ist ein tagesgenaues Kalender-Raster:

| Zellwert | Bedeutung |
|----------|-----------|
| Leer / Wochentag-Nummer | Unterrichtstag |
| `Fe` | Ferien |
| `Pr` | Praktikum |
| `Prüf` | Prüfung |
| Wochenende (Sa/So) | kein Unterricht |

Extrahiere eine **geordnete Liste aller Unterrichtstage** (tatsächliche Datumsangaben) über alle Jahresblätter.

**Hinweis zum Datenformat:** Die Jahresblätter enthalten in den Zusammenfassungszeilen (ca. Zeilen 37–44) Statistiken: Theorie-Tage, Stunden/Monat, Praktikumstage, Ferientage. Nutze diese zur Plausibilitätsprüfung.

## Schritt 3: Modulplan parsen

Lese die Modulliste mit den Feldern:
- **Modul**: Modul-ID (z.B. "1", "IMMO 1", "0")
- **Modulbezeichnung**: Titel des Moduls
- **UE**: Unterrichtseinheiten (Ganzzahl)

Falls der Modulplan aus einer Konzeptionsdatei (Word) extrahiert werden muss, suche nach Tabellen mit Spalten wie "Modul", "UE", "Stunden", "Unterrichtseinheiten".

Berechne: **Benötigte Tage pro Modul** = `ceil(UE / UE_pro_Tag)`

## Schritt 4: UE-Bilanz prüfen

Vergleiche verfügbare Unterrichtstage mit benötigten Tagen.

**Bei Defizit:** Zeige dem Nutzer eine Tabelle:

| Modul | Bezeichnung | UE geplant | Tage benötigt | Tage verfügbar | Differenz |
|-------|-------------|-----------|---------------|----------------|-----------|
| ...   | ...         | ...       | ...           | ...            | ...       |

Frage interaktiv, welche Module gekürzt werden sollen und um wie viel UE.
**Warte auf Bestätigung** bevor du fortfährst.

## Schritt 5: Excel-Output erstellen

Verwende `openpyxl` für die Ausgabe.

### Blatt 1: "Unterrichtsplan"

**Aufbau:**
- Zeile 1: Titel "Unterrichtsplan – [Kursname]" (merged, fett, Arial 16pt, dunkelblau #1F4E79)
- Zeile 2: Metadaten (Erstellungsdatum | Gesamtdauer | Gesamt-UE | Gesamt-Tage)
- Zeile 3: leer
- Zeile 4: Spaltenüberschriften
- Ab Zeile 5: Datensätze
- Letzte Zeile: Summenzeile

**Spalten:**

| Spalte | Inhalt | Breite |
|--------|--------|--------|
| A | Nr | 5 |
| B | Modul | 10 |
| C | Modulbezeichnung | 45 |
| D | Beginn | 14 |
| E | Ende | 14 |
| F | Tage | 8 |
| G | UE | 8 |
| H | Bemerkungen | 30 |

**Design:**
- Header-Zeile: Hintergrund `#1F4E79`, weißer fetter Text, Arial 11pt
- Ungerade Datenzeilen: weiß (`#FFFFFF`)
- Gerade Datenzeilen: hellblau (`#DDEEFF`)
- Summenzeile: hellgrau (`#D9D9D9`), fett
- Datumsformat: `DD.MM.YYYY`
- Alle Zellen: `Alignment(vertical='center')`, Zeilenhöhe 18pt
- Äußerer Rahmen um die gesamte Tabelle, innere Gitternetzlinien dünn grau

### Blatt 2: "Kalenderübersicht"

Monatliche Zusammenfassung:

| Jahr | Monat | Unterrichtstage | UE | Kumuliert UE |
|------|-------|----------------|----|--------------|

**Design:** gleiche Farbgebung wie Blatt 1.

### Ausgabedateiname

```
Unterrichtsplan_[Kursname]_[StartJahr].xlsx
```

Ausgabe im gleichen Verzeichnis wie die Eingabedateien, oder wenn nicht möglich im Arbeitsverzeichnis.

## Fehlerbehandlung

- **Unerwartete Zeitplan-Struktur**: Erkläre was gefunden wurde, frage den Nutzer
- **UE-Rundungsdifferenz**: Notiere Differenz in der Bemerkungsspalte der letzten Zeile
- **Modul mit 0 UE** (z.B. Prüfungstage): 1 Tag einplanen, Bemerkung "Prüfungstag"
- **Modulplan nicht gefunden**: Gib klare Fehlermeldung mit Hinweis auf erwartetes Format

## Beispiel-Daten (Referenz)

Entwickelt und getestet mit:
- Zeitplan: `Zeitplan_u-KK_HH.xlsx` (Hamburg, 2026–2028, 337 Unterrichtstage, 3.081 UE)
- Modulplan: aus `Konzeption u-IMMO.docx` extrahiert (20 Module, 9 UE/Tag)
- Output-Format orientiert sich an: `Unterrichtsplan u-IMMO.xlsx`
