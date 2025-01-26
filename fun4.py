import requests
import pandas as pd

def fetch_wikipedia_and_gbif_data_with_fallback(file_path, sheet_name="Sheet1"):
    """
    Pobiera dane o zwierzętach z Wikipedia API (opis, zdjęcie) i GBIF API (gromada).
    Jeśli dla nazwy łacińskiej nie znaleziono danych w Wikipedii, wyszukuje ponownie po nazwie angielskiej.
    
    :param file_path: Ścieżka do pliku Excel z tabelą zawierającą kolumny 'English Name' i 'Latin Name'.
    :param sheet_name: Nazwa arkusza w Excelu (domyślnie 'Sheet1').
    :return: None (zapisuje zmodyfikowany plik Excel na dysku)
    """
    # Wczytanie tabeli z pliku Excel
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Nie udało się wczytać pliku Excel: {e}")
        return
    
    # Sprawdzenie, czy wymagane kolumny istnieją
    if "Latin Name" not in df.columns or "English Name" not in df.columns:
        print("Kolumny 'Latin Name' i 'English Name' są wymagane w tabeli.")
        return
    
    # Przygotowanie kolumn wynikowych
    df["Description"] = None
    df["Class"] = None
    df["Image URL"] = None
    
    # Podstawowe URL dla API
    wikipedia_base_url = "https://en.wikipedia.org/w/api.php"
    gbif_base_url = "https://api.gbif.org/v1/species/match"
    
    for idx, row in df.iterrows():
        latin_name = row["Latin Name"]
        english_name = row["English Name"]
        
        if pd.isna(latin_name) and pd.isna(english_name):
            continue
        
        # Funkcja pomocnicza do wyszukiwania w Wikipedia API
        def fetch_from_wikipedia(name):
            params = {
                "action": "query",
                "titles": name,
                "prop": "extracts|pageimages",
                "exintro": True,
                "explaintext": True,
                "piprop": "original",
                "format": "json"
            }
            response = requests.get(wikipedia_base_url, params=params)
            result = response.json()
            
            # Przetwarzanie wyników z Wikipedia API
            pages = result.get("query", {}).get("pages", {})
            for page_id, page_data in pages.items():
                if page_id == "-1":  # Nie znaleziono artykułu
                    return None, None
                else:
                    description = page_data.get("extract", None)
                    image_url = page_data.get("original", {}).get("source", None)
                    return description, image_url
            return None, None
        
        try:
            # 1. Wikipedia API: Najpierw wyszukujemy po nazwie łacińskiej
            description, image_url = fetch_from_wikipedia(latin_name)
            
            # 2. Jeśli brak danych, próbujemy po nazwie angielskiej
            if not description and not image_url and not pd.isna(english_name):
                description, image_url = fetch_from_wikipedia(english_name)
            
            # Zapis wyników do tabeli
            df.at[idx, "Description"] = description
            df.at[idx, "Image URL"] = image_url
            
            # 3. GBIF API: Pobieranie gromady (class)
            gbif_params = {"name": latin_name}
            gbif_response = requests.get(gbif_base_url, params=gbif_params)
            gbif_result = gbif_response.json()
            
            # Pobieranie nazwy gromady z wyniku GBIF
            df.at[idx, "Class"] = gbif_result.get("class", None)
        
        except Exception as e:
            print(f"Error processing {latin_name} or {english_name}: {e}")
    
    # Zapisanie wyników z powrotem do pliku Excel
    output_path = file_path.replace(".xlsx", "_from_api.xlsx")
    try:
        df.to_excel(output_path, index=False)
        print(f"Zaktualizowano dane i zapisano do pliku: {output_path}")
    except Exception as e:
        print(f"Nie udało się zapisać pliku Excel: {e}")


# Ścieżka do pliku
file_path = ""

# Uruchomienie funkcji
fetch_wikipedia_and_gbif_data_with_fallback(file_path)