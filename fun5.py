import pandas as pd

def create_english_group_column(file_path, sheet_name="Sheet1"):
    """
    Tworzy nowe pole "English Group" w pliku Excel, tłumacząc i uogólniając nazwy gromad.
    Jeśli brak danych dla gromady, przypisuje ostatnią dostępną wartość "English Group".
    
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
    
    # Sprawdzenie, czy kolumna 'Class' istnieje
    if "Class" not in df.columns:
        print("Brakuje kolumny 'Class' w pliku.")
        return
    
    # Mapowanie gromad na grupy w języku angielskim
    class_to_group = {
        "Mammalia": "Mammals",
        "Aves": "Birds",
        "Crocodylia": "Reptiles",
        "Squamata": "Reptiles",
        "Testudines": "Reptiles",
        "Amphibia": "Amphibians",
        "Dipneusti": "Fish",
        "Elasmobranchii": "Fish"
    }
    
    # Dodanie kolumny 'English Group' i wstępne przypisanie wartości
    df["English Group"] = None
    last_known_group = None  # Przechowuje ostatnią znaną grupę

    for idx, row in df.iterrows():
        class_name = row.get("Class")
        
        if not pd.isna(class_name):  # Gdy gromada jest dostępna
            # Przypisz grupę na podstawie mapowania lub "Invertebrates", jeśli brak w mapie
            group = class_to_group.get(class_name, "Invertebrates")
            df.at[idx, "English Group"] = group
            last_known_group = group  # Aktualizuj ostatnią znaną grupę
            print(f"Przypisano '{group}' dla klasy '{class_name}' (wiersz {idx + 1}).")
        else:
            # Jeśli brak gromady, przypisz ostatnią znaną grupę
            if last_known_group:
                df.at[idx, "English Group"] = last_known_group
                print(f"Przypisano ostatnią znaną grupę '{last_known_group}' dla wiersza {idx + 1}.")
            else:
                print(f"Brak danych do przypisania dla wiersza {idx + 1}.")

    # Zapisanie wyników do nowego pliku
    output_path = file_path.replace(".xlsx", "_eng.xlsx")
    try:
        df.to_excel(output_path, index=False)
        print(f"Dodano kolumnę 'English Group' i zapisano plik: {output_path}")
    except Exception as e:
        print(f"Nie udało się zapisać pliku Excel: {e}")


file_path = ""
create_english_group_column(file_path)
