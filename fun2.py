import pandas as pd
from geopy.geocoders import Nominatim
from pyproj import Transformer
import time

def add_coordinates_to_zoo(file_path, sheet_name="Sheet1", api_delay=1):
    """
    Geokoduje adresy zoo i przelicza współrzędne na układ EPSG:2180, zapisując wyniki do pliku Excel.
    
    :param file_path: Ścieżka do pliku Excel zawierającego kolumnę 'Address'.
    :param sheet_name: Nazwa arkusza w pliku Excel (domyślnie "Sheet1").
    :param api_delay: Opóźnienie między zapytaniami do API geokodującego (w sekundach).
    :return: None
    """
    # Wczytanie pliku Excel
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Nie udało się wczytać pliku Excel: {e}")
        return
    
    # Sprawdzenie, czy kolumna Address istnieje
    if "Address" not in df.columns:
        print("Brakuje kolumny 'Address' w pliku.")
        return
    
    # Inicjalizacja geolokatora i transformatora
    geolocator = Nominatim(user_agent="Zoo_Geocoder")
    transformer = Transformer.from_crs("EPSG:4326", "EPSG:2180", always_xy=True)
    
    # Przygotowanie nowych kolumn
    df["X"] = None
    df["Y"] = None

    # Geokodowanie adresów
    for index, row in df.iterrows():
        address = row["Address"]
        if pd.isna(address):
            print(f"Adres pusty dla rekordu {index}. Pomijam.")
            continue

        try:
            # Geokodowanie adresu
            location = geolocator.geocode(address)
            if location:
                lon, lat = location.longitude, location.latitude
                # Transformacja współrzędnych na EPSG:2180
                x, y = transformer.transform(lon, lat)
                df.at[index, "X"] = x
                df.at[index, "Y"] = y
                print(f"Adres '{address}' geokodowany: X={x}, Y={y}")
            else:
                print(f"Nie znaleziono współrzędnych dla adresu: {address}")
        except Exception as e:
            print(f"Błąd przy geokodowaniu adresu '{address}': {e}")
        
        # Opóźnienie dla uniknięcia ograniczeń API
        time.sleep(api_delay)
    
    # Zapis wyników do nowego pliku
    output_path = file_path.replace(".xlsx", "_with_coordinates.xlsx")
    try:
        df.to_excel(output_path, index=False)
        print(f"Współrzędne dodane i zapisane w pliku: {output_path}")
    except Exception as e:
        print(f"Nie udało się zapisać pliku Excel: {e}")

# Przykład użycia
file_path = ""
add_coordinates_to_zoo(file_path)
