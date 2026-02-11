import camelot
import pandas as pd
import re


def clean_number(x):
    if not isinstance(x, str):
        return x
    s = x.replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except:
        return x


def is_actif_line(row):
    text = " ".join(str(x) for x in row)
    return bool(re.search(r"\bAC\s*\d+|\bTOTAL\b", text.upper()))


def normalize_columns(row):
    """
    Force max 6 colonnes CMF
    """
    row = [clean_number(x) for x in row if str(x).strip()]
    return row[:6]


def extract_actif(pdf_path, page_num, is_scanned=False):
    """
    Fonction appelée par main.py
    Doit retourner une liste de dictionnaires
    """

    print(f">>> Extraction ACTIF | page {page_num}")

    try:
        tables = camelot.read_pdf(
            pdf_path,
            flavor="stream",
            pages=str(page_num),
            row_tol=10,
            strip_text="\n",
        )

        print("Tables détectées :", tables.n)

        if tables.n == 0:
            print(":x: Aucun tableau détecté")
            return None

        rows = []

        for table in tables:
            for _, row in table.df.iterrows():
                if is_actif_line(row):
                    rows.append(normalize_columns(row.tolist()))

        if not rows:
            print(":x: Aucune ligne ACTIF détectée")
            return None

        df = pd.DataFrame(rows)

        # Force colonnes CMF
        df.columns = [
            "DESIGNATION",
            "BRUT",
            "AMORT_PROV",
            "NET_N",
            "NET_N1",
            
        ][:len(df.columns)]

        print(f":white_check_mark: {len(df)} lignes ACTIF extraites")

        # :warning: IMPORTANT :
        # main.py attend une liste de dictionnaires
        return df.to_dict("records")

    except Exception as e:
        print(f":x: Erreur extraction ACTIF : {e}")
        return None
'''if __name__ == "__main__":
    pdf_path = "COMPAGNIE_MEDITERRANEENNE_D_ASSURANCES_ET_DE_REASSURANCES_-_COMAR_-_Etats_financiers_au_31_12_2024.pdf"
    page_num = 2   # <-- toujours 2

    data = extract_actif(pdf_path, page_num)

    if data:
        df = pd.DataFrame(data)
        output_csv = "output_actif_test.csv"
        df.to_csv(output_csv, index=False, encoding="utf-8-sig")
        print(f":file_folder: CSV généré : {output_csv}")
    else:
        print(":x: Aucun résultat exporté") '''