import pandas as pd

def fill_missing_data(file_path, sheet_name="Sheet1"):
    """
    Uzupełnia brakujące dane w tabeli:
    - Jeśli brakuje opisu, wpisuje "Description not available".
    - Jeśli brakuje linku do fotografii, wstawia domyślny link do ikony zdjęcia niedostępnego.
    - Jeśli opis jest niepełny "{Latin Name} may refer to:", zamienia go na "Description not available".
    
    :param file_path: Ścieżka do pliku Excel.
    :param sheet_name: Nazwa arkusza w Excelu (domyślnie 'Sheet1').
    :return: None (zapisuje zmodyfikowany plik Excel na dysku)
    """
    try:
        # Wczytanie tabeli z pliku Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Nie udało się wczytać pliku Excel: {e}")
        return

    # Sprawdzenie, czy kolumny 'Description' i 'Image URL' istnieją
    if "Description" not in df.columns or "Image URL" not in df.columns:
        print("Brakuje kolumny 'Description' lub 'Image URL' w pliku.")
        return

    # Link domyślnej grafiki
    default_image_url = "https://cdn4.vectorstock.com/i/1000x1000/76/48/photo-not-found-icon-symbol-sign-vector-22437648.jpg"

    # Licznik zmian
    missing_description_count = df["Description"].isna().sum()
    missing_image_url_count = df["Image URL"].isna().sum()
    invalid_description_count = 0

    # Uzupełnianie braków w opisach
    df["Description"].fillna("Description not available", inplace=True)

    # Sprawdzanie i zamiana niepełnych opisów
    for index, row in df.iterrows():
        description = row["Description"]
        if isinstance(description, str) and description.endswith("may refer to:"):
            df.at[index, "Description"] = "Description not available"
            invalid_description_count += 1
            print(f"Zastąpiono niepełny opis dla rekordu w wierszu {index}.")

    # Uzupełnianie braków w linkach do zdjęć
    df["Image URL"].fillna(default_image_url, inplace=True)

    # Zapisanie zmodyfikowanego pliku
    output_path = file_path.replace(".xlsx", "_filled.xlsx")
    try:
        df.to_excel(output_path, index=False)
        print(f"Braki uzupełnione:")
        print(f" - Opisy: {missing_description_count} brakujących wpisów uzupełniono.")
        print(f" - Niepełne opisy: {invalid_description_count} wpisów zastąpiono.")
        print(f" - Linki do zdjęć: {missing_image_url_count} brakujących wpisów uzupełniono.")
        print(f"Zapisano plik: {output_path}")
    except Exception as e:
        print(f"Nie udało się zapisać pliku Excel: {e}")

# Ścieżka do pliku Excel
file_path = ""
fill_missing_data(file_path)
