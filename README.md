markdown# 🏭 SAP Lieferanten-Export & Reporting Tool v2.2

Professionelles Python-Tool für automatisierten SAP-Datenexport mit Sanktionslistencheck, intelligenter Farbcodierung und Management-Dashboard.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![Performance](https://img.shields.io/badge/Performance-1200%20records%20%3C%201s-green)
![DSGVO](https://img.shields.io/badge/DSGVO-Konform-green)
![Status](https://img.shields.io/badge/Status-Production%20Ready-success)

## 🎯 Problem & Lösung

**Problem:** Zollabteilungen verbringen täglich 2-3 Stunden mit manuellem SAP-Export, Datenbereinigung und Excel-Aufbereitung.

**Lösung:** Vollautomatische Verarbeitung mit Sanktionscheck, intelligenter Farbcodierung und Management-Dashboard in **unter 1 Sekunde** für 1.200 Datensätze.

## ⚡ Performance

- **1.200 Datensätze:** 0,916 Sekunden
- **10.000 Datensätze:** ~7,6 Sekunden (hochgerechnet)
- **Getestet auf:** MacBook Air M1, Python 3.13

## ✨ Features

### 📊 Dashboard (Sheet 1)
- **Kennzahlen-Übersicht** auf einen Blick
- Gesamtanzahl Lieferanten, Vollständig/Unvollständig
- Farbverteilung-Statistik (Grün/Gelb/Rot/Weiß)
- ⚠️ **Sanktionstreffer** hervorgehoben
- **Top 5 Länder** mit Anzahl
- Timestamp der Erstellung

### 🎨 Intelligente Farbcodierung (Sheet 2)
- 🟢 **Grün:** Vollständig + kein Sanktionstreffer
- 🟡 **Gelb:** Telefonnummer fehlt (aber sonst OK)
- 🔴 **Rot:** Auf EU-Sanktionsliste
- ⚪ **Weiß:** Andere unvollständige Felder

### 🔍 Sanktionslistencheck
- Automatischer Abgleich gegen EU-Sanktionsliste
- Eigenes Sheet "SANKTIONSLISTE" (dunkelrot markiert)
- Einfach erweiterbar (CSV-basiert)

### 🧹 Datenbereinigung
- **Duplikaterkennung** nach Lieferantennummer
- **Telefonnummern-Formatierung** (+49-Format)
- **Vollständigkeitsprüfung** (Pflichtfelder)

### 📈 5 Excel-Sheets
1. **DASHBOARD** - Management-Übersicht
2. **Alle Lieferanten** - Farbcodiert nach Status
3. **Nach Land** - Pivot-Tabelle mit Statistik
4. **Prüfliste** - Unvollständige Datensätze (hellrot)
5. **SANKTIONSLISTE** - Kritische Treffer (dunkelrot)

### 📝 Logging
- Detaillierte Log-Datei mit Timestamp
- Statistiken pro Durchlauf
- Fehlerprotokollierung

### 🔐 DSGVO-Konform
- 100% lokale Verarbeitung
- Keine Cloud-Uploads
- Keine personenbezogenen Daten

## 🚀 Schnellstart

### Installation
```bash
# Repository klonen
git clone https://github.com/luisschenk/sap-lieferanten-export-tool.git
cd sap-lieferanten-export-tool

# Abhängigkeiten installieren
pip3 install -r requirements.txt
Ausführung
bash# Testdaten erstellen (1.200 Beispiel-Lieferanten)
python3 create_test_data.py

# Tool ausführen
python3 sap_export_tool.py

# Excel-Report öffnen
open Lieferanten_Report_*.xlsx
📊 Output-Beispiel
Nach Ausführung erhalten Sie:
Excel-Report (Lieferanten_Report_YYYYMMDD_HHMMSS.xlsx):

📊 Sheet 1 "DASHBOARD": Kennzahlen, Top 5 Länder, Sanktionstreffer
🎨 Sheet 2 "Alle Lieferanten": Farbcodiert (1.200 Zeilen)
📈 Sheet 3 "Nach Land": Pivot-Tabelle
⚠️ Sheet 4 "Prüfliste": Unvollständige Datensätze
🚨 Sheet 5 "SANKTIONSLISTE": Kritische Treffer

Log-Datei (sap_export_log.txt):
[2025-10-09 12:56:34] === EXPORT GESTARTET ===
[2025-10-09 12:56:34] Eingelesen: 1215 Datensaetze
[2025-10-09 12:56:34] Sanktionsliste geladen: 3 Eintraege
[2025-10-09 12:56:34] Duplikate entfernt: 15
[2025-10-09 12:56:34] Unvollstaendig: 149
[2025-10-09 12:56:34] SANKTIONSTREFFER: 30
[2025-10-09 12:56:34] Farbverteilung - Gruen: 1021, Gelb: 12, Rot: 30
[2025-10-09 12:56:35] === EXPORT BEENDET ===
🏢 Produktiv-Einsatz
SAP-Integration (Echtdaten)
Option 1: CSV-Export aus SAP
python# Transaktion SE16 → Tabelle LFA1 → Export als CSV
# Umbenennen zu: sap_lieferanten.csv
Option 2: SAP RFC (mit pyrfc)
pythonfrom pyrfc import Connection

conn = Connection(
    user='SAPUSER', passwd='PASSWORD',
    ashost='sap.server.de', sysnr='00', client='100'
)
result = conn.call('RFC_READ_TABLE', QUERY_TABLE='LFA1')
Sanktionsliste anpassen
eu_sanctions.csv:
csvFirmenname,Land,Grund,Datum
Ihre Firma,Russland,Sanktionsgrund,2024-01-01
Automatisierung (täglich um 2 Uhr)
Windows Task Scheduler:
batchschtasks /create /tn "SAP_Export" /tr "python3 C:\path\sap_export_tool.py" /sc daily /st 02:00
Linux/Mac Cron:
bashcrontab -e
# Zeile hinzufügen:
0 2 * * * cd /path/to/tool && python3 sap_export_tool.py
🔧 Anpassungen
Pflichtfelder ändern
In sap_export_tool.py, Zeile 15:
pythonPFLICHTFELDER = ['Lieferantennr', 'Name', 'Land', 'Strasse']  # Nach Bedarf erweitern
Farblogik anpassen
In sap_export_tool.py, Funktion bestimme_farbkategorie():
pythondef bestimme_farbkategorie(row):
    if row['Sanktion'] == True:
        return 'ROT'
    elif row['Vollstaendig'] == True and row['Umsatz'] > 100000:  # NEU
        return 'GRUEN'
    # ...
📈 Use Cases

Zollabteilungen: Täglicher Lieferanten-Export mit Sanktionscheck
Einkauf: Lieferanten-Stammdaten für Verhandlungen
Controlling: Statistiken nach Land/Region
Compliance: Automatisches Sanktionslistenscreening
Management: Dashboard-Reporting

🛠️ Technologie-Stack

Python 3.8+
pandas: Datenverarbeitung, Excel-Export
openpyxl: Excel-Formatierung (Farben, Styles, Fonts)
re: Telefonnummern-Formatierung (Regex)

🎯 ROI für Kunden
Zeitersparnis:

Vorher: 2-3h/Tag manuell
Nachher: 0h (automatisch)
= 40-60h/Monat gespart

Kosteneinsparung:

Bei 50€/h Mitarbeiterkosten
= 2.000-3.000€/Monat

Risikominimierung:

Sanktionsverstöße vermeiden
Compliance sichergestellt
= unbezahlbar

🤝 Beitragen
Feedback und Verbesserungsvorschläge willkommen!
Kontakt: Luis@schenk.com

👤 Autor
Luis Schenk
Wirtschaftsjurist | Customs & Trade Compliance Specialist
📧 Luis@schenk.com
📍 Weinheim, Rhein-Neckar-Region, Deutschland
💼 Verfügbar: 15-20h/Woche (Werkstudent), ab 2025 Vollzeit
🎯 Weitere Projekte

Zoll-Präferenzkalkulator - Automatisierte EU-Präferenzursprungsberechnung
Custom GPT: Zoll-Assistent - KI-Assistent für Zollfragen
Custom GPT: Legal Translator - Mehrsprachige Legal-Tech (DE/EN/ES/FR/IT)

💼 Services

KI-gestützte Zoll-Automatisierung: Präferenzkalkulationen, Lieferantenerklärungen, Sanktionslistenscreening
SAP/Excel-Integration: Python-basierte Workflow-Automatisierung
Business Intelligence: DSGVO-konforme Datenverarbeitung und Dashboards
Custom-Tool-Entwicklung: Maßgeschneiderte Automatisierungen für Compliance & Zoll

📄 Lizenz
MIT License - Für kommerzielle und private Nutzung frei verwendbar.
Version: 2.2
Letzte Aktualisierung: Oktober 2025
GitHub: https://github.com/luisschenk/sap-lieferanten-export-tool
