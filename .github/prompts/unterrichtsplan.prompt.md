---
mode: agent
description: Erstellt einen Unterrichtsplan (Excel) aus Zeitplan + Modulplan.
tools: ['runInTerminal']
---

Erzeuge den Unterrichtsplan. Führe **genau einen** Terminal-Befehl aus und sonst
nichts – kein manuelles Einlesen, keine Zwischenschritte, keine Rückfragen:

```
python scripts/generate_unterrichtsplan.py --zeitplan "${input:zeitplan:Pfad zur Zeitplan-Excel}" --modulplan "${input:modulplan:Pfad zur Modulplan-Datei}" --kursname "${input:kursname:Kursname, z. B. u-STFA}"
```

Hinweise:
- Das Skript liest beide Dateien selbst ein und schreibt die Ausgabe selbst.
  Baue die Logik **nicht** nach.
- Verkette den Befehl nicht mit `cd …` o. Ä.; rufe `python …` direkt auf.
- Ein eventuelles Defizit an Unterrichtstagen wird vom Skript automatisch über
  die Kürzung der Prüfungsvorbereitung gelöst – greife nicht ein.
- Wenn der Lauf erfolgreich war, gib nur den Ausgabepfad aus der Zeile
  `[INFO] Saved: …` zurück.
