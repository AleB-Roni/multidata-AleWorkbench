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

## Arbeiten im Ticket

- **Kopf:** Betreff, Status, Priorität, Queue, Bearbeiter, Firma, Parent-Ticket und „Keine Benachrichtigungen“ werden nur vorgemerkt (gelb markiert) und gemeinsam mit **Speichern** übernommen; **Verwerfen** stellt den alten Stand wieder her. Bearbeiter und Firma sucht man direkt im Feld. Beim Firmenwechsel kann MultiData abrechenbare Stunden zurücksetzen – die Werkbank weist darauf hin.
- **🚗 Vor-Ort-Termin:** Haken beim Buchen (oder beim Bearbeiten eines Kommentars) setzt im Kommentar *Wegzeit* und *Wegstrecke* auf 1 – das Kennzeichen für die Buchhaltung. Solche Kommentare tragen das Etikett „🚗 vor Ort“.
- **Reiter unten:** *💬 Kommentare* und *🕘 Verlauf*. Der Verlauf zeigt Anlage, Kommentare, Mails, Anhänge und alle Feldänderungen aus dem MultiData-Protokoll (Tabelle `MDERRDA_MDERRDAP`: wer, wann, Feld, alt → neu – dasselbe wie der Reiter „Protokoll“ in MultiData). Ist das Protokoll nicht abrufbar, zeigt die Werkbank ersatzweise die Änderungen, die sie selbst erkannt hat.
- **Texteditor** (Kommentar, Kommentar bearbeiten, Signatur, Textbausteine):
  - **Tab** / **Umschalt+Tab** rücken Listenpunkte und Absätze ein bzw. aus, ebenso die Knöpfe ⇥ / ⇤. Eine neue Unterebene ist eine Aufzählung (•), außer die Liste hat schon nummerierte Unterebenen. In Tabellen springt Tab zur nächsten Zelle.
  - **Enter** in einem leeren Punkt und **Rücktaste** am Anfang eines Punkts rücken ihn eine Ebene aus. Auf oberster Ebene wird er ein normaler Absatz direkt unter der Liste; die Einfügemarke bleibt dort. Folgen noch Punkte, läuft die Nummerierung weiter.
  - „• Liste“ / „1. Liste“ schalten die Liste aus oder wechseln die Art.
  - **Einfügen:** Fremde Schriftarten, -größen, Abstände und Zeilenhöhen werden entfernt, damit der Text aussieht wie der eigene und man normal weiterschreiben kann. Fett, kursiv, unterstrichen, farbige Schrift, Hervorhebungen, Ausrichtung, Listen, Tabellen, Links und Bilder bleiben. Listen aus Word, Google Docs oder Web-Editoren werden echte Listen.
  - **Strg+Z** / **Strg+Y** machen auch Listen- und Einzugsänderungen rückgängig bzw. wieder.
  - **Größe:** Das Textfeld lässt sich unten rechts größer und kleiner ziehen. Die Werkbank merkt sich die Höhe je Editor (Kommentar, Kommentar bearbeiten, Signatur).
- **🔗 Teilen** kopiert die originale MultiData-Adresse des Tickets (`…/mdweb/rwi/MDERRDA/%7B%22SID%22:808626%7D`) in die Zwischenablage, ohne MultiData zu öffnen. Mit Umschalt+Klick zusätzlich „#Nr. Betreff“ davor.
- Große Tickets bauen sich stückweise auf; Kommentarinhalte werden erst gerendert, wenn sie in die Nähe des sichtbaren Bereichs kommen. Antwortet MultiData nicht, gibt es nach 30 s (plus einem automatischen zweiten Versuch) eine Fehlermeldung statt eines endlosen „Lade …“.

## Internes Zeiten-Ticket (Restzeiten)

Unter ⚙ → *Internes Zeiten-Ticket* wird das monatliche interne Ticket festgelegt (z. B. „Zeiten LE NOV 26“; Kürzel, Betreff-Muster und Firma sind einstellbar, Standard-Firma ist die erste interne Firma der Team-Konfiguration).

- **＋ Neues Ticket für <Monat>** legt es an: Priorität *Niedrig*, Status *Warten intern*, keine Benachrichtigungen. Das Ticket des Vormonats wird dabei per QuickClose geschlossen. Gibt es für den Monat schon eins, wird keins doppelt angelegt.
- Im **Zeiten-Tab** hat jeder Tag (bei den eigenen Zeiten) den Knopf **＋ intern**: Stunden eintragen, *Buchen* – der Rest wird als Kommentar ohne Text auf das interne Ticket **des jeweiligen Monats** gebucht. Mit Personio-Abgleich ist die Differenz schon vorausgefüllt („+1,25 h intern“), „Alle Fehlzeiten intern buchen“ erledigt alle Tage auf einmal.

## Protokoll

Alle Aktionen der Werkbank (nur in diesem Browser). Ticketnummern sind Links (Klick: öffnen, Mittelklick: neuer Tab); oben kann man nach Text oder Ticket-Nr. filtern.

## Anhänge im Browser öffnen

- **.eml** wird direkt in der Werkbank als Mail angezeigt: Kopfzeilen, Text und Anhänge, auch weitergeleitete Mails. Das passiert vollständig im Browser, ohne Upload. Externe Bilder lädt die Vorschau erst auf Klick, Skripte in der Mail werden nie ausgeführt.
- **Vorschau ohne Download** („👁 Vorschau“ bzw. „👁 Anzeigen“): Die Datei wird nur in den Arbeitsspeicher geladen und in einem neuen Fenster angezeigt, ohne Upload und ohne Anmeldung. Auch Anhänge innerhalb von .eml-Mails lassen sich so ansehen.
  - Bilder: png, jpg, gif, webp, bmp, svg, avif (Klick wechselt zwischen eingepasst und Originalgröße)
  - Video und Audio: mp4, webm, mov, mp3, wav, m4a, ogg … (was der Browser abspielen kann)
  - Word: docx (Text, Tabellen, Bilder, Kopf-/Fußzeilen; komplexe Layouts können abweichen)
  - Excel: xlsx, xlsm, xlsb, xls, ods, csv (alle sichtbaren Blätter als Reiter, je Blatt die ersten 2.000 Zeilen, Werte statt Formeln)
  - Text: txt, log, json, xml, ps1 …
  - Keine Vorschau: altes .doc, PowerPoint, Visio – dafür „In Office öffnen“.
  - Dokumentinhalte laufen in einer abgeschotteten Ansicht ohne Skripte. Die Bibliotheken dafür ([docx-preview](https://github.com/VolodymyrBaydalka/docxjs), [SheetJS](https://sheetjs.com), JSZip) lädt Tampermonkey einmalig mit Prüfsumme.
- **Word, Excel, PowerPoint und Visio**: „🖥 In Office öffnen“ lädt die Originaldatei herunter, der Browser öffnet sie im lokal installierten Office. Änderungen werden **nicht** nach MultiData zurückgeschrieben. Die Datei öffnet sich in der geschützten Ansicht; zum Bearbeiten „Bearbeitung aktivieren“.

### Office-Dateien ohne Klick öffnen

Damit sich die Datei nach dem Download sofort in Office öffnet, braucht der Browser für den Dateityp die Einstellung „Dateien dieses Typs immer öffnen“. Beim ersten Öffnen eines Dateityps zeigt die Werkbank dazu eine kurze Anleitung für Chrome bzw. Edge. Das ist einmal je Dateityp (docx, xlsx …) nötig. Unter ⚙ → Office-Dateien lassen sich die Hinweise wieder einblenden.

## Datenschutz

- Lokal im Browser (localStorage) liegen: Signatur, Textbausteine, Entwürfe, Filter und Vorlagen, Personio-Zeiten (falls importiert), Posteingangs-Vorschau, Protokoll, Ticket-Verlauf (zuletzt gesehener Stand der letzten 400 Tickets) und die Nummern der internen Zeiten-Tickets.
- Diese Daten werden nirgendwohin übertragen. Unter ⚙ → „Alle Werkbank-Daten in diesem Browser löschen“ lassen sie sich entfernen.
- Das Skript spricht im Betrieb ausschließlich mit `csp.synaforce.com`. Vorschauen entstehen im Browser; „In Office öffnen“ lädt in den Downloads-Ordner. Nichts wird in eine Cloud hochgeladen.
- Tampermonkey lädt beim Installieren/Aktualisieren zusätzlich die Vorschau-Bibliotheken von `cdn.jsdelivr.net` und `cdn.sheetjs.com` (fest versioniert, mit Prüfsumme).

## Hinweise

- Inoffizielles Werkzeug; es nutzt die interne Web-API von MultiData. Nach MultiData-Updates können einzelne Funktionen ausfallen, bis das Skript angepasst ist.
- Fehler und Wünsche bitte als GitHub-Issue oder direkt an Lukas Ebner.
