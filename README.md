# MultiData Werkbank

Tampermonkey-Userscript, das in der MultiData-Weboberfläche (`csp.synaforce.com/mdweb`) eine schnellere Arbeitsoberfläche einblendet: Ticketsuche mit Filtern und Vorlagen, Tickets und Kommentare bearbeiten, Zeiten buchen, Mails mit Anhängen senden, Tickets anlegen, QuickClose, Zeiten-Übersicht mit Personio-Abgleich. Die Werkbank kann in mehreren Browser-Tabs parallel laufen; die Tabs halten sich gegenseitig aktuell.

Das Skript nutzt die gleiche interne Schnittstelle wie die MultiData-Weboberfläche, und zwar mit der eigenen, angemeldeten Sitzung. Es gibt keine eigenen Zugangsdaten und keinen Server. Alle Einstellungen bleiben lokal im Browser.

## Installation (Kolleg:innen)

1. Die Browser-Erweiterung **Tampermonkey** installieren (Chrome/Edge Web Store).
2. Nur Chrome/Edge: Unter `chrome://extensions` bei Tampermonkey **„Details“ → „Nutzerskripts zulassen“** (bzw. den Entwicklermodus) aktivieren, sonst laufen keine Skripte.
3. Diesen Link öffnen, dann in Tampermonkey auf **Installieren** klicken:
   **https://raw.githubusercontent.com/lukasebner1998-ship-it/multidata-werkbank/main/stable/multidata-werkbank.user.js**
4. MultiData neu laden. Unten rechts erscheint **⚡ Werkbank** (oder Alt+Q).
5. Beim ersten Start unter **⚙ → Team-Konfiguration** die Datei importieren, die im Team geteilt wird. Alternativ selbst festlegen, welche Kunden nicht abrechenbar bzw. intern sind.

Updates: Tampermonkey prüft regelmäßig (standardmäßig täglich) und installiert neue Versionen automatisch. Sofort prüfen geht über das Tampermonkey-Menü → „Nach Userscript-Updates suchen“.

## Kanäle

| Kanal | Datei | Für wen |
|---|---|---|
| Beta | `beta/multidata-werkbank.user.js` | Entwicklung/Test (Version `x.y.z.n`, Name „… (Beta)“) |
| Stable | `stable/multidata-werkbank.user.js` | alle Kolleg:innen (Version `x.y.z`) |

Bitte **nicht beide** gleichzeitig aktiv haben; es läuft ohnehin nur die zuerst geladene.

## Release-Ablauf

1. Änderung in `beta/multidata-werkbank.user.js` einchecken und `@version` erhöhen: `1.5.0.1` → `1.5.0.2` … bzw. für ein neues Release `1.5.1.1`. Die Beta-Installation aktualisiert sich automatisch.
2. Selbst testen.
3. Freigeben: **Actions → „Beta freigeben“ → Run workflow** (optional mit kurzer Notiz). Der Workflow
   - kopiert die Beta nach `stable/`,
   - setzt `@name` ohne „(Beta)“, `@version` ohne die 4. Stelle (`1.5.1.3` → `1.5.1`) und die Update-URLs auf `stable/`,
   - legt einen Git-Tag und ein GitHub-Release an.
   Er bricht ab, wenn die Stable-Version nicht höher würde.
4. Kolleg:innen erhalten das Update automatisch über Tampermonkey.

## Team-Konfiguration

Kundenspezifisches (welche Kunden nicht abrechenbar bzw. intern sind) steht **nicht im Code**. Diese Angaben werden unter ⚙ gepflegt und als JSON-Datei weitergegeben (Format siehe `team-config.example.json`). Die echte Team-Datei gehört **nicht** in dieses Repository; sie wird intern geteilt, z. B. über Teams oder SharePoint.

## Anhänge im Browser öffnen

- **.eml** wird direkt in der Werkbank als Mail angezeigt: Kopfzeilen, Text und Anhänge, auch weitergeleitete Mails. Das passiert vollständig im Browser, ohne Upload. Externe Bilder lädt die Vorschau erst auf Klick, Skripte in der Mail werden nie ausgeführt.
- **Word, Excel, PowerPoint und Visio** öffnen sich in Office Online, wenn Microsoft 365 eingerichtet ist (siehe unten). Die Werkbank lädt die Datei dazu als Kopie in den eigenen OneDrive-Ordner `Werkbank-Anhänge/Ticket <Nr>/`. Ein erneutes Öffnen ersetzt die Kopie. Änderungen werden **nicht** nach MultiData zurückgeschrieben.

### Microsoft 365 einrichten (einmalig, durch eine Admin-Person)

1. In Entra ID unter **App-Registrierungen** → „Neue Registrierung“ anlegen.
   - Name: `MultiData Werkbank`
   - Kontotypen: nur dieses Organisationsverzeichnis
   - Umleitungs-URI: Plattform **Single-Page-Anwendung (SPA)**, `https://csp.synaforce.com/mdweb/werkbank-auth`
2. Unter **API-Berechtigungen** die delegierten Microsoft-Graph-Berechtigungen `User.Read`, `Files.ReadWrite` und `offline_access` hinzufügen. Danach optional „Administratorzustimmung erteilen“, damit niemand einzeln zustimmen muss.
3. Kein Client-Secret anlegen; eine SPA braucht keins.
4. In der Werkbank unter ⚙ → Microsoft 365 die **Verzeichnis-ID (Mandant)** und die **Anwendungs-ID** eintragen, dann „Übernehmen“. Danach die Team-Konfiguration exportieren; Mandant und Client-ID sind darin enthalten, Anmeldedaten nicht.

Beim ersten Öffnen einer Office-Datei meldet man sich einmal mit dem eigenen Microsoft-Konto an. Die Anmeldung bleibt im Browser gespeichert, bis man sich unter ⚙ abmeldet.

## Datenschutz

- Lokal im Browser (localStorage) liegen: Signatur, Textbausteine, Entwürfe, Filter und Vorlagen, Personio-Zeiten (falls importiert), Posteingangs-Vorschau und Protokoll.
- Diese Daten werden nirgendwohin übertragen. Unter ⚙ → „Alle Werkbank-Daten in diesem Browser löschen“ lassen sie sich entfernen.
- Das Skript spricht mit `csp.synaforce.com`. Nur wenn Microsoft 365 eingerichtet ist, kommen beim Öffnen von Office-Dateien `login.microsoftonline.com` und `graph.microsoft.com` hinzu (eigener Mandant, eigenes OneDrive).
- Lokal gespeichert wird dann zusätzlich die Microsoft-Anmeldung (Token).

## Hinweise

- Inoffizielles Werkzeug; es nutzt die interne Web-API von MultiData. Nach MultiData-Updates können einzelne Funktionen ausfallen, bis das Skript angepasst ist.
- Fehler und Wünsche bitte als GitHub-Issue oder direkt an Lukas Ebner.
