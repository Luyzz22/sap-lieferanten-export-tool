import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
import os
import sys
from datetime import datetime
import re

INPUT_CSV = "sap_lieferanten.csv"
OUTPUT_EXCEL = "Lieferanten_Report_{}.xlsx".format(datetime.now().strftime("%Y%m%d_%H%M%S"))
LOG_FILE = "sap_export_log.txt"
EMAIL_EMPFAENGER = "zoll-team@kunde.de"
PFLICHTFELDER = ['Lieferantennr', 'Name', 'Land']
SANKTIONSLISTE_CSV = "eu_sanctions.csv"

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"
    print(log_entry)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_entry + "\n")

def bereinige_telefonnummer(telefon):
    if pd.isna(telefon) or telefon == "":
        return ""
    telefon_clean = re.sub(r'[^\d+]', '', str(telefon))
    if not telefon_clean.startswith('+'):
        if telefon_clean.startswith('0'):
            telefon_clean = '+49' + telefon_clean[1:]
        else:
            telefon_clean = '+49' + telefon_clean
    return telefon_clean

def pruefe_vollstaendigkeit(row):
    for feld in PFLICHTFELDER:
        if pd.isna(row[feld]) or row[feld] == "":
            return False
    return True

def lade_sanktionsliste():
    if not os.path.exists(SANKTIONSLISTE_CSV):
        log_message(f"Warnung: {SANKTIONSLISTE_CSV} nicht gefunden.")
        return []
    try:
        df_sanctions = pd.read_csv(SANKTIONSLISTE_CSV, encoding='utf-8')
        sanktionen = df_sanctions['Firmenname'].str.lower().tolist()
        log_message(f"Sanktionsliste geladen: {len(sanktionen)} Eintraege")
        return sanktionen
    except Exception as e:
        log_message(f"Fehler beim Laden: {e}")
        return []

def bereinige_firmenname_fuer_matching(name):
    """
    Bereinigt Firmenname für intelligentes Matching:
    - Lowercase
    - Entfernt Sonderzeichen
    - Entfernt Stopwörter (GmbH, AG, etc.)
    - Gibt Set von Wörtern zurück
    """
    if pd.isna(name) or name == "":
        return set()
    
    # Stopwörter die ignoriert werden
    stopwords = {
        'gmbh', 'ag', 'kg', 'co', 'ltd', 'inc', 'se', 'sa',
        '&', 'und', 'and', 'et', 'the', 'de', 'la', 'le',
        'mbh', 'ohg', 'gbr', 'ug', 'eg', 'ev'
    }
    
    # Lowercase und Sonderzeichen entfernen
    name_clean = str(name).lower().strip()
    # Ersetze Sonderzeichen durch Leerzeichen
    name_clean = re.sub(r'[^a-zäöüß0-9\s]', ' ', name_clean)
    # Mehrfache Leerzeichen entfernen
    name_clean = re.sub(r'\s+', ' ', name_clean)
    
    # In Wörter splitten und Stopwörter entfernen
    words = set([w for w in name_clean.split() if w not in stopwords and len(w) > 1])
    
    return words

def pruefe_sanktionsliste_optimiert(firmenname, sanktionsliste, verbose=False):
    """
    OPTIMIERTE VERSION: Word-Based Matching
    
    Prüft ob ALLE wichtigen Wörter aus der Sanktionsliste
    im Firmennamen vorkommen.
    
    WICHTIG: Mindestens 2 Wörter müssen in der Sanktionsliste sein,
    sonst zu viele False Positives!
    
    Beispiele:
    - "Schwarz Group AG" matched "schwarz group" ✓ (2 Wörter)
    - "Schwarz Industries" matched "schwarz group" ✗ (group fehlt!)
    - "Mueller Schwarz KG" matched "schwarz group" ✗ (group fehlt!)
    - "Krüger & Sons GmbH" matched "krüger sons" ✓ (2 Wörter)
    - "Hartmann Trading" matched "hartmann ag" ✗ (nur 1 Wort nach Stopwörtern!)
    
    Args:
        firmenname: Name der zu prüfenden Firma
        sanktionsliste: Liste von sanktionierten Firmennamen
        verbose: Wenn True, logge jeden Treffer
    
    Returns:
        True wenn Treffer, False wenn kein Treffer
    """
    if not sanktionsliste or pd.isna(firmenname) or firmenname == "":
        return False
    
    # Firmenname bereinigen
    firmenname_words = bereinige_firmenname_fuer_matching(firmenname)
    
    if not firmenname_words:
        return False
    
    # Prüfe gegen jede Sanktion
    for sanktion in sanktionsliste:
        sanktion_words = bereinige_firmenname_fuer_matching(sanktion)
        
        if not sanktion_words:
            continue
        
        # SCHUTZ: Mindestens 2 Wörter müssen in Sanktionsliste sein
        # Sonst zu viele False Positives (z.B. nur "Hartmann")
        if len(sanktion_words) < 2:
            if verbose:
                log_message(f"ÜBERSPRUNGEN: '{sanktion}' hat nur {len(sanktion_words)} Wort(e) - zu unspezifisch")
            continue
        
        # WICHTIG: ALLE Wörter aus Sanktionsliste müssen im Firmennamen sein
        if sanktion_words.issubset(firmenname_words):
            if verbose:
                log_message(f"TREFFER: '{firmenname}' matched '{sanktion}' (Wörter: {sanktion_words})")
            return True
    
    return False

def bestimme_farbkategorie(row):
    if row['Sanktion'] == True:
        return 'ROT'
    elif row['Vollstaendig'] == True:
        return 'GRUEN'
    elif pd.isna(row['Telefon']) or row['Telefon'] == '':
        return 'GELB'
    else:
        return 'WEISS'

def hauptprogramm():
    print("=" * 70)
    print("SAP LIEFERANTEN-EXPORT v2.3 (mit optimiertem Sanktionscheck)")
    print("=" * 70)
    log_message("=== EXPORT GESTARTET ===")
    
    if not os.path.exists(INPUT_CSV):
        print(f"FEHLER: {INPUT_CSV} nicht gefunden!")
        sys.exit(1)
    
    df = pd.read_csv(INPUT_CSV, encoding='utf-8')
    log_message(f"Eingelesen: {len(df)} Datensaetze")
    
    sanktionsliste = lade_sanktionsliste()
    
    duplikate_vorher = len(df)
    df = df.drop_duplicates(subset=['Lieferantennr'], keep='first')
    duplikate_entfernt = duplikate_vorher - len(df)
    
    df['Telefon'] = df['Telefon'].apply(bereinige_telefonnummer)
    df['Vollstaendig'] = df.apply(pruefe_vollstaendigkeit, axis=1)
    
    # OPTIMIERTER SANKTIONSCHECK (mit verbose=True für Logging)
    df['Sanktion'] = df['Name'].apply(
        lambda x: pruefe_sanktionsliste_optimiert(x, sanktionsliste, verbose=True)
    )
    
    df['Farbkategorie'] = df.apply(bestimme_farbkategorie, axis=1)
    
    anzahl_unvollstaendig = len(df[df['Vollstaendig'] == False])
    anzahl_sanktionen = len(df[df['Sanktion'] == True])
    anzahl_gruen = len(df[df['Farbkategorie'] == 'GRUEN'])
    anzahl_gelb = len(df[df['Farbkategorie'] == 'GELB'])
    anzahl_rot = len(df[df['Farbkategorie'] == 'ROT'])
    anzahl_weiss = len(df) - anzahl_gruen - anzahl_gelb - anzahl_rot
    
    log_message(f"Duplikate entfernt: {duplikate_entfernt}")
    log_message(f"Unvollstaendig: {anzahl_unvollstaendig}")
    log_message(f"SANKTIONSTREFFER: {anzahl_sanktionen}")
    log_message(f"Farbverteilung - Gruen: {anzahl_gruen}, Gelb: {anzahl_gelb}, Rot: {anzahl_rot}")
    
    # Top 5 Länder
    top_laender = df.groupby('Land').size().reset_index(name='Anzahl').sort_values('Anzahl', ascending=False).head(5)
    
    # Kritische ohne Name/Land
    kritisch = len(df[(df['Name'] == '') | (df['Land'] == '')])
    
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        # SHEET 0: DASHBOARD
        dashboard_data = {
            'Kennzahl': [
                'GESAMTANZAHL LIEFERANTEN',
                'Vollstaendig',
                'Unvollstaendig',
                '',
                'FARBVERTEILUNG',
                'Gruen (OK)',
                'Gelb (Telefon fehlt)',
                'Rot (SANKTION)',
                'Weiss (Sonstige)',
                '',
                'KRITISCH',
                'Ohne Name/Land',
                'Duplikate entfernt',
                '',
                'Erstellt am'
            ],
            'Wert': [
                len(df),
                len(df[df['Vollstaendig'] == True]),
                anzahl_unvollstaendig,
                '',
                '',
                anzahl_gruen,
                anzahl_gelb,
                anzahl_rot,
                anzahl_weiss,
                '',
                '',
                kritisch,
                duplikate_entfernt,
                '',
                datetime.now().strftime('%d.%m.%Y %H:%M')
            ]
        }
        df_dashboard = pd.DataFrame(dashboard_data)
        df_dashboard.to_excel(writer, sheet_name='DASHBOARD', index=False, startrow=2)
        
        # Top 5 Länder daneben
        top_laender.to_excel(writer, sheet_name='DASHBOARD', index=False, startrow=2, startcol=4)
        
        # Sheet 1: Alle Lieferanten
        df_output = df.drop(['Vollstaendig', 'Sanktion', 'Farbkategorie'], axis=1)
        df_output.to_excel(writer, sheet_name='Alle Lieferanten', index=False)
        
        # Sheet 2: Nach Land
        pivot = df.groupby('Land').size().reset_index(name='Anzahl')
        pivot = pivot.sort_values('Anzahl', ascending=False)
        pivot.to_excel(writer, sheet_name='Nach Land', index=False)
        
        # Sheet 3: Prüfliste
        df_unvollstaendig = df[df['Vollstaendig'] == False].drop(['Vollstaendig', 'Sanktion', 'Farbkategorie'], axis=1)
        df_unvollstaendig.to_excel(writer, sheet_name='Pruefliste', index=False)
        
        # Sheet 4: Sanktionsliste
        df_sanktionen = df[df['Sanktion'] == True].drop(['Vollstaendig', 'Sanktion', 'Farbkategorie'], axis=1)
        if len(df_sanktionen) > 0:
            df_sanktionen.to_excel(writer, sheet_name='SANKTIONSLISTE', index=False)
            log_message(f"Sheet SANKTIONSLISTE: {len(df_sanktionen)} Treffer")
    
    wb = load_workbook(OUTPUT_EXCEL)
    
    # DASHBOARD FORMATIEREN
    if 'DASHBOARD' in wb.sheetnames:
        ws = wb['DASHBOARD']
        
        # Überschrift
        ws['A1'] = 'LIEFERANTEN-ANALYSE DASHBOARD'
        ws['A1'].font = Font(size=16, bold=True)
        ws['A1'].alignment = Alignment(horizontal='left')
        
        # Spaltenbreiten
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 15
        
        # Überschrift Top 5 Länder
        ws['E1'] = 'TOP 5 LAENDER'
        ws['E1'].font = Font(size=14, bold=True)
        
        # Kritische Werte rot markieren
        rot_bg = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
        if anzahl_rot > 0:
            ws['B10'].fill = rot_bg
        if kritisch > 0:
            ws['B14'].fill = rot_bg
        
        # Fette Überschriften
        ws['A3'].font = Font(bold=True, size=12)
        ws['A7'].font = Font(bold=True, size=12)
        ws['A13'].font = Font(bold=True, size=12)
        
        log_message("Dashboard formatiert")
    
    # Alle Lieferanten farbcodieren
    if 'Alle Lieferanten' in wb.sheetnames:
        ws = wb['Alle Lieferanten']
        gruen = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        gelb = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        rot = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
        
        for idx, row in df.iterrows():
            excel_zeile = idx + 2
            farbe = None
            
            if row['Farbkategorie'] == 'GRUEN':
                farbe = gruen
            elif row['Farbkategorie'] == 'GELB':
                farbe = gelb
            elif row['Farbkategorie'] == 'ROT':
                farbe = rot
            
            if farbe:
                for spalte in range(1, ws.max_column + 1):
                    ws.cell(row=excel_zeile, column=spalte).fill = farbe
        
        log_message("Farbcodierung angewandt")
    
    # Prüfliste formatieren
    if 'Pruefliste' in wb.sheetnames:
        ws = wb['Pruefliste']
        hellrot = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
        for zeile in range(2, ws.max_row + 1):
            for spalte in range(1, ws.max_column + 1):
                ws.cell(row=zeile, column=spalte).fill = hellrot
    
    # Sanktionsliste formatieren
    if 'SANKTIONSLISTE' in wb.sheetnames:
        ws = wb['SANKTIONSLISTE']
        dunkelrot = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        for zeile in range(2, ws.max_row + 1):
            for spalte in range(1, ws.max_column + 1):
                ws.cell(row=zeile, column=spalte).fill = dunkelrot
    
    wb.save(OUTPUT_EXCEL)
    
    print("\n" + "=" * 70)
    print("ERFOLGREICH!")
    print(f"Datensaetze: {len(df)}")
    print(f"Gruen: {anzahl_gruen} | Gelb: {anzahl_gelb} | Rot: {anzahl_rot} | Weiss: {anzahl_weiss}")
    print(f"SANKTIONSTREFFER: {anzahl_rot}")
    print(f"\nExcel: {OUTPUT_EXCEL}")
    print("=" * 70)
    log_message("=== EXPORT BEENDET ===")

if __name__ == "__main__":
    hauptprogramm()
