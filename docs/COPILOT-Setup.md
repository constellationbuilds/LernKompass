# LernKompass mit Copilot Chat – ohne ständige Erlaubnisabfragen

Kurzantwort: **Ja, das bekommen wir hin** – mit *GitHub Copilot in VS Code
(Agent-Modus)*. Der Trick ist, dass die gesamte Logik in **einem einzigen
Skript-Aufruf** steckt. Ein einziger Befehl = höchstens **eine** Erlaubnisabfrage –
und wenn `python` einmal freigegeben ist, **null** Abfragen.

---

## Welches „Copilot" funktioniert?

| Variante | Lokales Python-Skript ausführbar? | Geeignet? |
|----------|-----------------------------------|-----------|
| **GitHub Copilot in VS Code (Agent-Modus)** | ✅ Ja – führt `python …` lokal aus, liest/schreibt lokale Excel-Dateien | **Ja – diese Variante nutzen** |
| **Microsoft 365 Copilot** (Word/Excel/Teams) | ❌ Nein – Cloud-LLM ohne lokales Terminal/Dateizugriff | Nein (Neubau auf Power Platform nötig) |

Der bestehende Workflow (lokales Python liest/schreibt lokale Excel-Dateien)
läuft **nur** unter GitHub Copilot in VS Code. Für M365 Copilot müsste die Logik
auf Power Automate / Office Scripts / Copilot Studio neu gebaut werden – das ist
ein eigenes Projekt (siehe Roadmap im README).

---

## Warum keine Abfragen? Das Prinzip

- **Ein selbst-enthaltener Befehl = höchstens eine Abfrage.** Das Skript
  `generate_unterrichtsplan.py` erledigt alle drei Schritte intern (Zeitplan
  lesen → Modulplan lesen → Unterrichtsplan schreiben). Copilot muss nur **einen**
  Terminal-Befehl absetzen.
- **Mehrere Tool-Schritte = viele Abfragen.** Deshalb sagen
  `.github/copilot-instructions.md` und die Prompt-Datei dem Modell ausdrücklich:
  Aufgabe **nicht** zerlegen, nur das Skript aufrufen.
- Auch das **Defizit-Handling läuft ohne Rückfrage**: Reichen die
  Unterrichtstage nicht, kürzt das Skript automatisch die Prüfungsvorbereitung
  (nie die FPB). Es gibt also keine interaktive Eingabeaufforderung mehr.

---

## Einrichtung (einmalig)

### 1. Voraussetzungen
```powershell
pip install openpyxl pandas
```
`python` (oder `py`) muss im PATH sein.

### 2. Auto-Freigabe für `python`

Diese Repo-Einstellungen liegen schon in [`.vscode/settings.json`](../.vscode/settings.json)
und greifen, sobald der LernKompass-Ordner in VS Code geöffnet ist.

Damit es in **jedem** Ordner (auch im Kursordner auf OneDrive) ohne Abfrage
läuft, dieselben Schlüssel zusätzlich in die **Benutzer-Einstellungen** eintragen:
`Strg+Shift+P` → *Preferences: Open User Settings (JSON)* →

```jsonc
"chat.tools.terminal.enableAutoApprove": true,
"chat.tools.terminal.autoApprove": {
  "python": true, "py": true, "python3": true,
  "rm": false, "del": false, "Remove-Item": false, "kill": false
}
```

> Hinweis: Die genauen Schlüsselnamen wurden 2025 mehrfach umbenannt. Falls VS Code
> sie nicht erkennt, in den Einstellungen nach **„auto approve"** suchen und den
> dort angezeigten Schlüssel verwenden (ggf. `chat.agent.terminal.allowList`).

### 3. Workspace vertrauen
Beim Öffnen des Ordners **„Vertrauen / Trust"** anklicken. Im *Restricted Mode*
werden Tools trotz Freigabeliste blockiert.

### 4. (Alternative ohne JSON) „Immer zulassen"
Bei der **ersten** Abfrage im Dropdown **„Immer zulassen / Always allow"** →
Workspace wählen. VS Code schreibt die Regel dann selbst in die Einstellungen.
Ab dann: keine Abfragen mehr.

---

## Nutzung

Im Copilot-Chat (Agent-Modus):

```
/unterrichtsplan
```
→ fragt im Chat nach Zeitplan-, Modulplan-Pfad und Kursname.

Oder natürlichsprachlich, z. B.:
> „Bau mir einen Unterrichtsplan aus diesen zwei Dateien: … und …, Kurs u-STFA."

Dank `.github/copilot-instructions.md` führt beides auf denselben **einen**
Befehl – und damit auf null Erlaubnisabfragen im eingerichteten Zustand.

---

## Windows/PowerShell-Stolperfallen

- **Keine Befehlsverkettung** (`cd … ; python …`): Verkettete Befehle werden
  einzeln geprüft – ein nicht freigegebenes `cd` löst trotzdem eine Abfrage aus.
  Die Anweisungsdateien verbieten Verkettung deshalb.
- **`py` vs. `python` vs. `python3`**: alle drei freigeben (siehe oben).
- **Backslashes in Regex-Regeln** in `settings.json` doppelt schreiben (`\\.`).
- **Unternehmens-Policy**: In verwalteten Umgebungen kann eine Gruppenrichtlinie
  Auto-Approve global deaktivieren. Bleiben Abfragen trotz korrekter Einstellung,
  ist das die wahrscheinliche Ursache (IT ansprechen).
