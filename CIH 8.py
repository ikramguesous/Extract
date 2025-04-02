#!/usr/bin/env python
# coding: utf-8

# In[8]:


import re
import pandas as pd
import PyPDF2
import os

def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def normalize_text(text):
    text = text.replace("’", "'").replace("‘", "'")
    text = text.replace("\u00A0", " ").replace("\u202F", " ")
    return text

def find_table_in_text(text, start_patterns, end_patterns, max_lines=100):
    text = normalize_text(text)
    for start_pattern in start_patterns:
        start_match = re.search(start_pattern, text, re.IGNORECASE)
        if start_match:
            start_idx = start_match.end()
            remaining_text = text[start_idx:]

            lines = []
            for line in remaining_text.split("\n"):
                line_stripped = line.strip()
                lines.append(line_stripped)

                # Vérifie si la ligne correspond à l'un des motifs de fin
                for end_pattern in end_patterns:
                    line_normalized = normalize_text(line_stripped)
                    if re.search(end_pattern, line_normalized, re.IGNORECASE):
                        return "\n".join(lines).strip()

                if len(lines) >= max_lines:
                    return "\n".join(lines).strip()

    return None


def clean_numeric_value(value):
    if not value or value.strip() in ["-", "+"]:
        return "-"
    value = value.replace(" ", "")
    cleaned = re.sub(r"[^\d\.,+\-]", "", value)
    cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return cleaned

def extract_table_data(text):
    text = normalize_text(text)
    lines = text.split("\n")
    pattern = re.compile(
        r'^\s*(?:[+\-]\s*)?(.*?)\s{2,}((?:[+\-]?\d[\d\s\.,]*)|-)\s{2,}((?:[+\-]?\d[\d\s\.,]*)|-)\s*$'
    )
    extracted_data = []
    for line in lines:
        match = pattern.match(line.strip())
        if match:
            libelle = match.group(1).strip()
            val1 = clean_numeric_value(match.group(2))
            val2 = clean_numeric_value(match.group(3))
            extracted_data.append([libelle, val1, val2])
    return extracted_data

def extract_bank_tables(text, table_configs):
    tables = {}
    for table_name, config in table_configs.items():
        raw_table_text = find_table_in_text(
            text,
            config["start_patterns"],
            config["end_patterns"],
            config.get("max_lines", 100)
        )
        if raw_table_text:
            data = extract_table_data(raw_table_text)
            if data:
                columns = ["Libellé", "val1", "val2"]
                df = pd.DataFrame(data, columns=columns)
                tables[table_name] = df
            else:
                tables[table_name] = pd.DataFrame()
        else:
            tables[table_name] = pd.DataFrame()
    return tables

def save_to_excel(financial_data, output_file):
    if not financial_data:
        print("Aucune donnée à enregistrer.")
        return
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for table_name, df in financial_data.items():
            df.to_excel(writer, sheet_name=table_name[:31], index=False)
            workbook = writer.book
            worksheet = writer.sheets[table_name[:31]]
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D9D9D9',
                'border': 1
            })
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(0, 0, 40)
            worksheet.set_column(1, 2, 15)
    print(f"Données financières enregistrées dans {output_file}")

bank_configs = {
    "boa": {
        "Bilan Actif": {
            "start_patterns": [r"ACTIF\s+\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}/\d{1,2}/\d{4}"
                                                , r"ACTIF"],
            "end_patterns": [r"TOTAL DE L'ACTIF"],
            "max_lines": 40
        },
        "Bilan Passif": {
            "start_patterns": [r"PASSIF\s+\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}/\d{1,2}/\d{4}"
                                , r"PASSIF"],
            "end_patterns": [r"TOTAL DU PASSIF"],
            "max_lines": 50
        },
        "Compte de Produits et Charges": {
            "start_patterns": [
                        r"(?m)^COMPTE DE PRODUITS ET CHARGES\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})"

    ],
            "end_patterns": [
                        r"(?m)^R[ÉE]SULTAT NET DE L[’']EXERCICE"
    ],
            "max_lines": 50
        },
        "Etat des Soldes de Gestion": {
            "start_patterns": [r"(?m)^[ÉE]TAT DES SOLDES DE GESTION.*?\d{1,2}/\d{1,2}/\d{4}.*?\d{1,2}/\d{1,2}/\d{4}"

],
            "end_patterns": [r"AUTOFINANCEMENT"],
            "max_lines": 50
        }
    },
    "bcp": {
        "Bilan Actif": {
            "start_patterns": [
                r"BILAN\s*ACTIF[^0-9]*\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}/\d{1,2}/\d{4}",
                r"ACTIF\s+\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}/\d{1,2}/\d{4}"

            ],
            "end_patterns": [
                r"TOTAL DE L'ACTIF",
                r"TOTAL\s+DE\s+L['’]ACTIF"
            ],
            "max_lines": 40
        },
        "Bilan Passif": {
            "start_patterns": [
                r"BILAN\s*PASSIF[^0-9]*\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}/\d{1,2}/\d{4}",
                r"PASSIF\s+\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}/\d{1,2}/\d{4}"

            ],
            "end_patterns": [
                r"TOTAL DU PASSIF",
                r"TOTAL\s+PASSIF",
                r"COMPTE DE PRODUITS"
            ],
            "max_lines": 50
        },
        "Compte de Produits et Charges": {
            "start_patterns": [
                r"COMPTE DE PRODUITS ET CHARGES\s*\n(?:.*\n)?\s*30/06/2024\s+30/06/2023",
                r"PRODUITS D'EXPLOITATION BANCAIRE\s*\n(?:.*\n)?\s*30/06/2024\s+30/06/2023"

            ],
            "end_patterns": [
                r"RÉSULTAT NET DE L'EXERCICE",
                r"ETAT DES SOLDES"
            ],
            "max_lines": 50
        },
        "Etat des Soldes de Gestion": {
            "start_patterns": [
                r"ÉTAT\s+DES\s+SOLDES\s+DE\s+GESTION",
                r"TABLEAU\s+DE\s+FORMATION\s+DE\s+RÉSULTAT"
            ],
            "end_patterns": [
                r"RÉSULTAT\s+NET\s+DE\s+L'EXERCICE",
                r"AUTOFINANCEMENT"
            ],
            "max_lines": 50
        }
    },
    "cdm": {
        "Bilan Actif": {
            "start_patterns": [
                r"BILAN\s+AU\s+30\s+JUIN\s+2024",
                r"ACTIF"
            ],
            "end_patterns": [
                r"TOTAL DE L'ACTIF"
            ],
            "max_lines": 40
        },
        "Bilan Passif": {
            "start_patterns": [
                r"PASSIF\s+30/06/2024\s+31/12/2023",
                r"PASSIF"
            ],
            "end_patterns": [
                r"TOTAL DU PASSIF"
            ],
            "max_lines": 50
        },
        "Compte de Produits et Charges": {
            "start_patterns": [
                r"COMPTE DE PRODUITS ET CHARGES\s+AU\s+30\s+JUIN\s+2024"
            ],
            "end_patterns": [
                r"RÉSULTAT NET DE L'EXERCICE"
            ],
            "max_lines": 50
        },
        "Etat des Soldes de Gestion": {
            "start_patterns": [
                r"[ÉE]TAT DES SOLDES DE GESTION\s+AU\s+30\s+JUIN\s+2024"
            ],
            "end_patterns": [
                r"R[ÉE]SULTAT NET DE L[’']EXERCICE"
            ],
            "max_lines": 50
        }
    },
    "cih": {
        "Bilan Actif": {
            "start_patterns": [
                r"BILAN ACTIF\s+.*?ACTIF\s+juin-24\s+déc-23"
            ],
            "end_patterns": [
                r"TOTAL\s+ACTIF\S*(?:.*?\n)?"
            ],
            "max_lines": 40
        },
        "Bilan Passif": {
            "start_patterns": [
                r"BILAN PASSIF\s+.*?PASSIF\s+juin-24\s+déc-23"
            ],
            "end_patterns": [
                r"TOTAL PASSIF(?:.*?\n)?"
            ],
            "max_lines": 50
        },
        "Compte de Produits et Charges": {
            "start_patterns": [
                r"COMPTE DE PRODUITS ET CHARGES\s+Libellé\s+juin-24\s+juin-23"
            ],
            "end_patterns": [
                r"RESULTAT NET DE L'EXERCICE(?:.*?\n)?"
            ],
            "max_lines": 50
        },
        "Etat des Soldes de Gestion": {
            "start_patterns": [
                r"ÉTAT DES SOLDES DE GESTION\s+Libellé\s+juin-24\s+juin-23"
            ],
            "end_patterns": [
                r"RESULTAT NET DE L'EXERCICE(?:.*?\n)?"
            ],
            "max_lines": 50
        }
    },
    "cfg": {
        "Bilan Actif": {
            "start_patterns": [
                r"BILAN\s*\n?.*ACTIF[^0-9]*30/06/2024\s+31/12/2023",
                r"ACTIF\s+30/06/2024\s+31/12/2023"
            ],
            "end_patterns": [
                r"TOTAL DE L'ACTIF",
                r"TOTAL\s+ACTIF",
                r"PASSIF"
            ],
            "max_lines": 50
        },
        "Bilan Passif": {
            "start_patterns": [
                r"BILAN\s*\n?.*PASSIF[^0-9]*30/06/2024\s+31/12/2023",
                r"PASSIF\s+30/06/2024\s+31/12/2023"
            ],
            "end_patterns": [
                r"TOTAL DU PASSIF",
                r"TOTAL\s+PASSIF",
                r"COMPTE DE PRODUITS"
            ],
            "max_lines": 50
        },
        "Compte de Produits et Charges": {
            "start_patterns": [
                r"COMPTE DE PRODUITS ET CHARGES.*?30/06/2024\s+30/06/2023",
                r"PRODUITS D'EXPLOITATION BANCAIRE.*?30/06/2024\s+30/06/2023"
            ],
            "end_patterns": [
                r"RESULTAT NET DE L'EXERCICE",
                r"ETAT DES SOLDES"
            ],
            "max_lines": 50
        },
        "Etat des Soldes de Gestion": {
            "start_patterns": [
                r"ETAT DES SOLDES DE GESTION.*?30/06/2024\s+30/06/2023",
                r"TABLEAU DE FORMATION DES RESULTATS.*?30/06/2024\s+30/06/2023"
            ],
            "end_patterns": [
                r"RESULTAT NET DE L'EXERCICE",
                r"CAPACITÉ D'AUTOFINANCEMENT"
            ],
            "max_lines": 50
        }
    }
}

def get_bank_type(pdf_filename, bank_configs):
    pdf_filename_lower = pdf_filename.lower()
    for bank in bank_configs:
        if bank.lower() in pdf_filename_lower:
            return bank
    return None

def main():
    pdf_path = "BCP_RFS_juin_2024.pdf"
    if not os.path.exists(pdf_path):
        print(f"Erreur: Fichier '{pdf_path}' inexistant.")
        return

    print(f"Extraction du texte du PDF: {pdf_path}")
    full_text = extract_text_from_pdf(pdf_path)
    if not full_text:
        print("Échec de l'extraction du texte du PDF.")
        return

    text_output = os.path.splitext(pdf_path)[0] + "_texte_extrait.txt"
    with open(text_output, "w", encoding="utf-8") as f:
        f.write(full_text)
    print(f"Texte extrait enregistré dans {text_output}")

    bank_type = get_bank_type(pdf_path, bank_configs)
    if bank_type is None:
        print("Aucun type de banque trouvé dans le nom du fichier.")
        return

    print(f"Configuration utilisée pour la banque: {bank_type}")
    table_configs = bank_configs[bank_type]

    print("Extraction des tableaux financiers...")
    financial_data = extract_bank_tables(full_text, table_configs)
    if not financial_data:
        print("Aucune donnée extraite.")
        return

    output_file = os.path.splitext(pdf_path)[0] + "_Tableaux_Extraits.xlsx"
    save_to_excel(financial_data, output_file)
    print(f"\nExtraction terminée et enregistrée dans {output_file}")

if __name__ == "__main__":
    main()


# In[ ]:




