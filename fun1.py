import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_animal_data_with_id_and_iucn(pdf_path, output_excel_path, starting_id=1001):
    """
    Ekstraktuje dane zwierząt z pliku PDF, przypisuje unikalne ID i zapisuje dane w Excelu.
    Sprawdza unikalność rekordów na podstawie English Name, Latin Name i IUCN Status.
    
    :param pdf_path: Ścieżka do pliku PDF.
    :param output_excel_path: Ścieżka do pliku wyjściowego Excel.
    :param starting_id: Początkowa wartość unikalnego ID dla rekordów (domyślnie 1001).
    """
    reader = PdfReader(pdf_path)
    animal_records = []  # Lista do przechowywania danych (UniqueID, English Name, Latin Name, IUCN Status)
    unique_id = starting_id  # Początkowy ID

    for page in reader.pages:
        lines = page.extract_text().split("\n")
        
        for i, line in enumerate(lines):
            if "CITES" in line:  # Znaleziono linię z informacjami o zwierzęciu
                # Znajdź angielską nazwę
                if "OBSOLETE" in line:
                    english_name_match = re.search(r'OBSOLETE.*?/ (.+?),', line)
                else:
                    english_name_match = re.search(r'/ (.+?),', line)
                
                if english_name_match:
                    english_name = english_name_match.group(1)
                    
                    # Znajdź łacińską nazwę (linia poprzedzająca obecną, o ile istnieje)
                    latin_name = lines[i - 1].strip() if i > 0 else None
                    
                    # Znajdź status IUCN
                    iucn_match = re.search(r'IUCN: (.+?)(?:,|$)', line)
                    iucn_status = iucn_match.group(1) if iucn_match else "Not available"

                    # Sprawdź, czy rekord już istnieje (ignorując ID)
                    is_duplicate = any(
                        record[1] == english_name and 
                        record[2] == latin_name and 
                        record[3] == iucn_status
                        for record in animal_records
                    )

                    if not is_duplicate:
                        # Dodaj dane do listy, przypisując unikalne ID
                        animal_records.append((unique_id, english_name, latin_name, iucn_status))
                        unique_id += 1

    # Tworzenie DataFrame i zapisywanie do pliku Excel
    df = pd.DataFrame(animal_records, columns=["UniqueID", "English Name", "Latin Name", "IUCN Status"])
    df.to_excel(output_excel_path, index=False)
    print(f"Plik Excel zapisano jako: {output_excel_path}")
    print(f"Liczba zapisanych rekordów: {len(animal_records)}")

# Ścieżki do pliku PDF i wyjściowego pliku Excel
pdf_path = ""
output_excel_path = ""

# Uruchomienie funkcji
extract_animal_data_with_id_and_iucn(pdf_path, output_excel_path)

