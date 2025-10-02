import pandas as pd
import simplekml
import re
import hashlib

# --- Konfiguracja ---
PLIK_WEJSCIOWY = 'dane_radiowe.xlsx'
PLIK_WYJSCIOWY = 'linie_radiowe.kmz'
SZEROKOSC_LINII = 3 #szerokość linii na mapie
KOLUMNA_ID = 'Nr_pozw/dec'

# --- Funkcje ---

PREDEFINED_COLORS = {
    'P4 Sp. z o.o.': simplekml.Color.purple,
    'T-Mobile Polska S.A.': simplekml.Color.magenta,
    'ORANGE POLSKA S.A.': simplekml.Color.orange,
    'Towerlink Poland Sp. z o.o.': simplekml.Color.green
}

def przypisz_kolor(operator_name):
    if operator_name in PREDEFINED_COLORS:
        return PREDEFINED_COLORS[operator_name]
    else:
        hash_object = hashlib.sha256(str(operator_name).encode())
        hex_dig = hash_object.hexdigest()
        bgr = hex_dig[0:6]
        return f"ff{bgr[4:6]}{bgr[2:4]}{bgr[0:2]}"

def dms_to_dd(dms):
    if not isinstance(dms, str):
        return None
    
    parts = re.findall(r"(\d+)[^\d\w\.]*([NSEW])[^\d\w\.]*(\d+)'([\d\.]+)", dms.replace("''", "'"))
    if not parts:
        return None
    parts = parts[0]
    degrees = float(parts[0])
    direction = parts[1]
    minutes = float(parts[2])
    seconds = float(parts[3])
    dd = degrees + minutes / 60 + seconds / 3600
    if direction in ['S', 'W']:
        dd *= -1
    return dd

# --- Główna część skryptu ---

try:
    df = pd.read_excel(PLIK_WEJSCIOWY)
except FileNotFoundError:
    print(f"Błąd: Nie znaleziono pliku '{PLIK_WEJSCIOWY}'.")
    exit()

print(f"Wczytano {len(df)} wierszy.")
df.drop_duplicates(subset=[KOLUMNA_ID], keep='first', inplace=True)
print(f"Po usunięciu duplikatów zostało {len(df)} unikalnych linków.")

df['lat_tx'] = df['Sz_geo_Tx'].apply(dms_to_dd)
df['lon_tx'] = df['Dl_geo_Tx'].apply(dms_to_dd)
df['lat_rx'] = df['Sz_geo_Rx'].apply(dms_to_dd)
df['lon_rx'] = df['Dl_geo_Rx'].apply(dms_to_dd)
df.dropna(subset=['lat_tx', 'lon_tx', 'lat_rx', 'lon_rx'], inplace=True)
print(f"Po przetworzeniu współrzędnych do dalszej pracy jest {len(df)} linków.")

kml = simplekml.Kml(name="Linki Radioliniowe (pełne informacje)")

for operator in sorted(df['Operator'].unique()):
    op_folder = kml.newfolder(name=str(operator))
    df_operator = df[df['Operator'] == operator]
    kolor_operatora = przypisz_kolor(operator)
    
    for index, row in df_operator.iterrows():
        opis_linku = f"""
            <![CDATA[
            <div style="font-family: Arial, sans-serif; font-size: 14px; max-width: 500px;">
                <h3 style="background-color: #e8e8e8; padding: 5px;">Informacje Główne</h3>
                <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
                    <tr><td style="background-color: #f2f2f2; width: 40%;"><b>Operator</b></td><td>{row.get('Operator', 'Brak danych')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Nr pozwolenia</b></td><td>{row.get(KOLUMNA_ID, 'Brak danych')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Data ważności</b></td><td>{row.get('Data_ważn_pozw/dec', 'Brak danych')}</td></tr>
                </table>

                <h3 style="background-color: #e8e8e8; padding: 5px; margin-top: 15px;">Parametry Radiowe</h3>
                <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
                    <tr><td style="background-color: #f2f2f2; width: 40%;"><b>Częstotliwość</b></td><td>{row.get('f [GHz]', '?')} GHz</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Przepływność</b></td><td>{row.get('Przepływność [Mb/s]', '?')} Mb/s</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Szerokość kanału</b></td><td>{row.get('Szer_kan [MHz]', '?')} MHz</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Modulacja</b></td><td>{row.get('Rodz_modu-lacji', 'Brak danych')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Moc EIRP</b></td><td>{row.get('EIRP [dBm]', '?')} dBm</td></tr>
                </table>

                <h3 style="background-color: #e8e8e8; padding: 5px; margin-top: 15px;">Lokalizacja i Sprzęt</h3>
                <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
                    <tr style="background-color: #d0d0d0;"><th colspan="2">Nadajnik (Tx)</th></tr>
                    <tr><td style="background-color: #f2f2f2; width: 40%;"><b>Lokalizacja</b></td><td>{row.get('Miejscowość Tx', '')}, {row.get('Ulica Tx', '')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Wys. n.p.m.</b></td><td>{row.get('H_t_Tx [m npm]', '?')} m</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Producent anteny</b></td><td>{row.get('Prod_ant_Tx', 'Brak danych')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Typ anteny</b></td><td>{row.get('Typ_ant_Tx', 'Brak danych')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Zysk anteny</b></td><td>{row.get('Zysk_ant_Tx [dBi]', '?')} dBi</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Wys. zawieszenia</b></td><td>{row.get('H_ant_Tx [m npt]', '?')} m</td></tr>
                    
                    <tr style="background-color: #d0d0d0;"><th colspan="2">Odbiornik (Rx)</th></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Lokalizacja</b></td><td>{row.get('Miejscowość Rx', '')}, {row.get('Ulica Rx', '')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Wys. n.p.m.</b></td><td>{row.get('H_t_Rx [m npm]', '?')} m</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Producent anteny</b></td><td>{row.get('Prod_ant_Rx', 'Brak danych')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Typ anteny</b></td><td>{row.get('Typ_ant_Rx', 'Brak danych')}</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Zysk anteny</b></td><td>{row.get('Zysk_ant_Rx [dBi]', '?')} dBi</td></tr>
                    <tr><td style="background-color: #f2f2f2;"><b>Wys. zawieszenia</b></td><td>{row.get('H_ant_Rx [m npt]', '?')} m</td></tr>
                </table>
            </div>
            ]]>
        """
        linia = op_folder.newlinestring(
            name=f"{row.get('Miejscowość Tx', '')} - {row.get('Miejscowość Rx', '')}",
            description=opis_linku
        )
        linia.coords = [(row['lon_tx'], row['lat_tx']), (row['lon_rx'], row['lat_rx'])]
        linia.style.linestyle.color = kolor_operatora
        linia.style.linestyle.width = SZEROKOSC_LINII
        linia.altitudemode = simplekml.AltitudeMode.clamptoground

kml.savekmz(PLIK_WYJSCIOWY)

print(f"\nGotowe! Plik '{PLIK_WYJSCIOWY}' został pomyślnie wygenerowany z pełnymi informacjami.")
