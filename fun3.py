import pandas as pd
from PyPDF2 import PdfReader

def extract_animal_zoo_data(pdf_path, animal_excel_path, zoo_excel_path, output_excel_path):
    """
    Tworzy nową tabelę zawierającą relację między zwierzętami a kodami zoo oraz ich lokalizacjami.
    
    :param pdf_path: Ścieżka do pliku PDF z danymi.
    :param animal_excel_path: Ścieżka do pliku Excel z danymi zwierząt.
    :param zoo_excel_path: Ścieżka do pliku Excel z kodami zoo i ich lokalizacjami.
    :param output_excel_path: Ścieżka do pliku wynikowego Excel.
    :return: None
    """
    # Wczytanie danych z plików Excel
    animal_df = pd.read_excel(animal_excel_path)
    zoo_df = pd.read_excel(zoo_excel_path)

    # Upewnienie się, że wymagane kolumny istnieją
    if not {'Latin Name', 'UniqueID'}.issubset(animal_df.columns):
        print("Brak kolumn 'Latin Name' lub 'UniqueID' w pliku ze zwierzętami.")
        return
    if not {'Code', 'X', 'Y'}.issubset(zoo_df.columns):
        print("Brak kolumn 'Code', 'X', 'Y' w pliku z kodami zoo.")
        return

    # Wczytanie pliku PDF
    reader = PdfReader(pdf_path)
    pdf_lines = [line.strip() for page in reader.pages for line in page.extract_text().split("\n")]

    # Tworzenie listy do wyników
    results = []

    # Mapowanie danych dla szybkiego dostępu
    latin_name_to_id = dict(zip(animal_df['Latin Name'], animal_df['UniqueID']))
    zoo_code_to_location = {
        row['Code']: (row['X'], row['Y']) for _, row in zoo_df.iterrows()
    }

    current_animal_id = None

    # Analiza pliku PDF
    for line in pdf_lines:
        # Sprawdź, czy linia zawiera nazwę łacińską zwierzęcia
        if line in latin_name_to_id:
            current_animal_id = latin_name_to_id[line]  # Ustaw obecne zwierzę
        elif current_animal_id and line in zoo_code_to_location:
            # Jeśli znaleziono kod zoo, dodaj rekord do wyników
            x, y = zoo_code_to_location[line]
            results.append({
                'UniqueID': current_animal_id,
                'Code': line,
                'X': x,
                'Y': y
            })

    # Tworzenie DataFrame z wyników
    result_df = pd.DataFrame(results)

    # Zapis wyników do pliku Excel
    try:
        result_df.to_excel(output_excel_path, index=False)
        print(f"Wyniki zapisano do pliku: {output_excel_path}")
    except Exception as e:
        print(f"Błąd podczas zapisywania pliku: {e}")

# Przykład użycia
pdf_path = ""
animal_excel_path = ""
zoo_excel_path = ""
output_excel_path = ""

extract_animal_zoo_data(pdf_path, animal_excel_path, zoo_excel_path, output_excel_path)
