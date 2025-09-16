# PDF Invoice Extraction (Python)

Ein Python-Skript zur automatischen Extraktion von Rechnungsbeträgen aus mehreren PDF-Dateien und zur Aggregation der Ergebnisse in einer zentralen Excel-Übersicht.

## Das Problem

Die manuelle Übertragung von Daten aus PDF-Rechnungen in Excel-Tabellen ist ein häufiger, aber extrem zeitaufwendiger und fehleranfälliger Prozess. Mitarbeiter müssen jede PDF einzeln öffnen, den korrekten Gesamtbetrag finden und ihn manuell in eine Liste kopieren.

**Ausgangssituation:** Ein Ordner voller PDF-Rechnungen, aus denen jeweils der Gesamtbetrag extrahiert werden muss.

<img width="990" height="755" alt="ScreenshotRechnung" src="https://github.com/user-attachments/assets/af6331aa-28fd-4d5e-871b-412b8c81dea9" />


## Die Lösung

Dieses Skript automatisiert den gesamten Extraktions- und Sammelprozess. Es liest alle PDF-Dateien in einem bestimmten Ordner, identifiziert und extrahiert den höchsten numerischen Wert (in der Regel der Gesamtbetrag) aus jeder Datei und speichert die Ergebnisse übersichtlich in einer neu erstellten Excel-Datei.

**Das Ergebnis:** Eine perfekt formatierte Excel-Tabelle mit einer Auflistung aller Rechnungsdateien und der zugehörigen Gesamtbeträge, erstellt in Sekunden.

<img width="567" height="592" alt="ScreenshotExcel" src="https://github.com/user-attachments/assets/ece40b7e-7763-4b33-bbe3-05b177abbaab" />


## Funktionsweise im Detail

1.  **PDF-Verarbeitung:** Das Skript iteriert durch alle Dateien in einem vordefinierten Quellordner und verarbeitet ausschließlich Dateien mit der Endung `.pdf`.
2.  **Textextraktion:** Für jede PDF-Datei wird der gesamte Textinhalt der ersten Seite mithilfe der `pypdf`-Bibliothek ausgelesen.
3.  **Mustererkennung mit Regex:** Ein regulärer Ausdruck (`re.compile(r"[0-9]+,[0-9]+")`) wird verwendet, um alle Zahlen im Format "1234,56" zu finden, die typisch für Geldbeträge sind.
4.  **Identifikation des Gesamtbetrags:** Da auf einer Rechnung mehrere Beträge stehen können (Einzelpreise, Nettobetrag, etc.), identifiziert das Skript den **höchsten** gefundenen Betrag als den wahrscheinlichsten Gesamtbetrag.
5.  **Datenaggregation:** Die Dateinamen und die zugehörigen extrahierten Beträge werden in einer Liste gesammelt. Die Konsole gibt während des Prozesses Live-Feedback.
<img width="1578" height="183" alt="ScreenshotKonsole" src="https://github.com/user-attachments/assets/3fe26740-2664-44ad-a227-3021bc818e4a" />

6.  **Excel-Reporting:** Am Ende des Prozesses werden die gesammelten Daten mithilfe der `openpyxl`-Bibliothek in eine neue, sauber formatierte Excel-Datei namens `Rechnungsbetraege.xlsx` geschrieben.

## Verwendete Technologien
- **Python**
- **pypdf** (zum Lesen und Extrahieren von Text aus PDF-Dateien)
- **openpyxl** (zum Erstellen und Bearbeiten von `.xlsx`-Dateien)
- **re (Regular Expressions)** (zur Mustererkennung und Extraktion der Beträge)
- **os** (zur Interaktion mit dem Dateisystem)
