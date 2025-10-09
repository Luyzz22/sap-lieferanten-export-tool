import pandas as pd
import random

# 10 Länder für realistische Verteilung
laender = [
    'Deutschland', 'Frankreich', 'Italien', 'Spanien', 'Polen',
    'Niederlande', 'Belgien', 'Oesterreich', 'Schweiz', 'Tschechien'
]

# Erweiterte Firmenliste (inkl. Sanktionsfirmen)
firmen_basis = [
    'Mueller', 'Schmidt', 'Weber', 'Fischer', 'Meyer', 'Wagner', 'Becker',
    'Schulz', 'Hoffmann', 'Schaefer', 'Koch', 'Bauer', 'Richter', 'Klein',
    'Wolf', 'Schroeder', 'Neumann', 'Schwarz', 'Zimmermann', 'Braun',
    'Krueger', 'Hartmann', 'Lange', 'Schmitt', 'Werner', 'Krause', 'Meier'
]

# SANKTIONSFIRMEN (müssen häufiger vorkommen)
sanktions_firmen = [
    'Krüger & Sons',
    'Schwarz Group', 
    'Hartmann AG'
]

firmen_suffixe = ['GmbH', 'AG', '& Co', 'KG', 'Industries', 'Automotive', 
                  'Parts', 'Logistics', 'Trading', 'Components', 'Materials',
                  'Supply', 'Tech', 'Exports', 'Group', 'Ltd', 'Corporation']

strassen = [
    'Hauptstrasse', 'Industriestrasse', 'Bahnhofstrasse', 'Gartenweg',
    'Feldstrasse', 'Ringstrasse', 'Am Markt', 'Berliner Allee',
    'Muenchner Strasse', 'Koelner Ring', 'Frankfurter Weg'
]

staedte = [
    'Muenchen', 'Berlin', 'Hamburg', 'Koeln', 'Frankfurt', 'Stuttgart',
    'Duesseldorf', 'Dortmund', 'Essen', 'Leipzig', 'Bremen', 'Dresden',
    'Hannover', 'Nuernberg', 'Duisburg', 'Bochum', 'Wuppertal', 'Bielefeld',
    'Bonn', 'Muenster', 'Karlsruhe', 'Mannheim', 'Augsburg', 'Wiesbaden'
]

def generiere_firmenname():
    """Generiert realistischen Firmennamen"""
    basis = random.choice(firmen_basis)
    suffix = random.choice(firmen_suffixe)
    return f"{basis} {suffix}"

def generiere_telefon(fehlt=False):
    """Generiert Telefonnummer (oder leer wenn fehlt=True)"""
    if fehlt:
        return ""
    
    formate = [
        f"+49 {random.randint(100,999)} {random.randint(1000000,9999999)}",
        f"0{random.randint(100,999)}-{random.randint(1000000,9999999)}",
        f"+49{random.randint(100,999)}{random.randint(1000000,9999999)}",
        f"0{random.randint(100,999)} {random.randint(1000000,9999999)}"
    ]
    return random.choice(formate)

def generiere_testdaten(anzahl=1200):
    """Generiert realistische Testdaten für 1200 Lieferanten"""
    
    print(f"Generiere {anzahl} Testdatensaetze...")
    
    daten = []
    sanktions_counter = 0
    
    for i in range(anzahl):
        # 15% unvollständige Datensätze
        ist_unvollstaendig = random.random() < 0.15
        
        # 8% ohne Telefonnummer (zusätzlich)
        telefon_fehlt = random.random() < 0.08
        
        # Alle 40-50 Einträge eine Sanktionsfirma einfügen
        if i % random.randint(40, 50) == 0 and sanktions_counter < 30:
            firma = random.choice(sanktions_firmen)
            sanktions_counter += 1
        else:
            if ist_unvollstaendig:
                firma = ''
            else:
                firma = generiere_firmenname()
                
                # Manchmal Variationen der Sanktionsfirmen
                if random.random() < 0.02:
                    if 'Schwarz' in firma:
                        firma = random.choice(sanktions_firmen)
                        sanktions_counter += 1
        
        datensatz = {
            'Lieferantennr': f"LF{1000 + i:06d}",
            'Name': firma,
            'Strasse': f"{random.choice(strassen)} {random.randint(1, 200)}",
            'PLZ': f"{random.randint(10000, 99999)}",
            'Ort': random.choice(staedte),
            'Land': random.choice(laender) if not ist_unvollstaendig else '',
            'Telefon': generiere_telefon(fehlt=telefon_fehlt)
        }
        
        daten.append(datensatz)
    
    # 10-15 Duplikate hinzufügen (für Duplikatserkennung)
    anzahl_duplikate = random.randint(10, 15)
    for _ in range(anzahl_duplikate):
        original_idx = random.randint(0, len(daten) - 1)
        daten.append(daten[original_idx].copy())
    
    print(f"Generiert: {len(daten)} Datensaetze (inkl. {anzahl_duplikate} Duplikate)")
    print(f"Sanktionsfirmen eingebaut: ca. {sanktions_counter}")
    
    return daten

if __name__ == "__main__":
    print("=" * 70)
    print("SAP LIEFERANTEN - GROSSER TESTDATEN-GENERATOR")
    print("=" * 70)
    
    # 1200 Datensätze generieren
    testdaten = generiere_testdaten(1200)
    df = pd.DataFrame(testdaten)
    
    # Als CSV speichern
    dateiname = "sap_lieferanten.csv"
    df.to_csv(dateiname, index=False, encoding='utf-8')
    
    # Statistik
    anzahl_ohne_name = len(df[df['Name'] == ''])
    anzahl_ohne_telefon = len(df[df['Telefon'] == ''])
    
    print("\n" + "=" * 70)
    print(f"✅ {dateiname} erfolgreich erstellt!")
    print("=" * 70)
    print(f"📊 Statistik:")
    print(f"   • Datensaetze gesamt:        {len(df)}")
    print(f"   • Ohne Firmenname:           {anzahl_ohne_name} (~15%)")
    print(f"   • Ohne Telefonnummer:        {anzahl_ohne_telefon} (~8%)")
    print(f"   • Verschiedene Laender:      {df['Land'].nunique()}")
    print(f"   • Enthaltene Sanktionsfirmen: ~30")
    print("\n🚀 Starte nun das Hauptprogramm mit:")
    print("   python3 sap_export_tool.py")
    print("\n⏱️  Performance-Test mit 1.200 Datensaetzen!")
    print("=" * 70)
