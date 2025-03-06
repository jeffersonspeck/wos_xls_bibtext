import os
import pandas as pd
from tqdm import tqdm

def find_xls_files():
    """Procura arquivos XLS no diretório atual."""
    return [file for file in os.listdir() if file.endswith('.xls')]

def clean_value(value):
    """Converte valores para string e remove espaços extras, evitando erros."""
    if pd.isna(value):
        return None 
    return str(value).strip()

def extract_wos_data(df):
    """Extrai e processa dados do Web of Science a partir de um DataFrame pandas."""
    bib_entries = []
    keys = {}

    for _, row in tqdm(df.iterrows(), total=len(df), desc="Processing entries"):
        first_author = clean_value(row.get("Authors"))
        publication_year = clean_value(row.get("Publication Year"))

        if first_author and publication_year:
            first_author_surname = first_author.split(",")[0].split()[-1]
            base_key = f"{first_author_surname}_{publication_year}"
        else:
            continue

        key = base_key
        counter = 97  # 'a' em ASCII
        while key in keys:
            key = f"{base_key}{chr(counter)}"
            counter += 1
        keys[key] = True

        bib_entry = f"@article{{{key},\n"

        bib_fields = {
            "title": "Article Title",
            "author": "Authors",
            "editor": "Book Editors",
            "journal": "Source Title",
            "year": "Publication Year",
            "volume": "Volume",
            "number": "Issue",
            "doi": "DOI",
            "issn": "ISSN",
            "eissn": "eISSN",
            "isbn": "ISBN",
            "abstract": "Abstract",
            "language": "Language",
            "keywords": "Author Keywords",
            "publisher": "Institution",
            "conference": "Conference Title",
            "address": "Institution Address",
            "pubmed_id": "Pubmed Id",
            "document_type": "Document Type",
            "degree": "Degree Name",
        }

        start_page = clean_value(row.get("Start Page"))
        end_page = clean_value(row.get("End Page"))
        if start_page and end_page:
            pages = f"{start_page}-{end_page}"
        elif start_page:
            pages = start_page
        else:
            pages = None

        if pages:
            bib_entry += f"  pages={{ {pages} }},\n"

        for field, column in bib_fields.items():
            value = clean_value(row.get(column))
            if value:
                bib_entry += f"  {field}={{ {value} }},\n"

        bib_entry += "}\n\n"
        bib_entries.append(bib_entry)

    return bib_entries

def convert_xls_to_bib():
    """Lê arquivos XLS, processa os dados e gera um arquivo BibTeX."""
    xls_files = find_xls_files()
    
    if not xls_files:
        print("No .xls files found in the current directory.")
        return

    print(f"\nProcessing the following XLS files:")
    for file in xls_files:
        print(f" - {file}")

    all_bib_entries = []
    for xls_file in xls_files:
        try:
            df = pd.read_excel(xls_file, engine="xlrd")
            bib_entries = extract_wos_data(df)
            all_bib_entries.extend(bib_entries)
        except Exception as e:
            print(f"Error processing {xls_file}: {e}")

    if all_bib_entries:
        output_name = input("\nEnter the output file name (without extension, press Enter for default): ").strip()
        if not output_name:
            output_name = "WebOfScience_Results"
        bibtex_file = f"{output_name}.bib"

        with open(bibtex_file, "w", encoding="utf-8") as bibfile:
            bibfile.writelines(all_bib_entries)

        print(f"\nBibTeX file generated: {bibtex_file}")
    else:
        print("\nNo valid entries found for BibTeX generation.")

convert_xls_to_bib()
