import camelot
import pandas as pd
import re

def fix_ac_header_shift(values):
    designation, brut, amort, net_n, net_n1 = values

    if re.match(r"^AC\d+$", brut.strip().upper()):
        brut = amort
        amort = net_n
        net_n = net_n1
        net_n1 = ""

    return [designation, brut, amort, net_n, net_n1]

def clean_text(x):
    if x is None:
        return ""
    return str(x).replace("\u00a0", " ").strip()


def is_actif_line(row):
    text = " ".join(str(x) for x in row)
    return bool(re.search(r"\bAC\s*\d+|\bTOTAL\b", text.upper()))


def normalize_columns(row):
    row = [clean_text(x) for x in row]

    while len(row) < 5:
        row.append("")

    return row[:5]


def shift_if_needed(values):
    """
    Corrige le décalage:
    si BRUT vide et AMORT_PROV contient un chiffre => shift à gauche
    """
    designation, brut, amort, net_n, net_n1 = values

    if brut == "" and re.search(r"\d", amort):
        brut = amort
        amort = net_n
        net_n = net_n1
        net_n1 = ""

    return [designation, brut, amort, net_n, net_n1]


def extract_total_code(row):
    """
    Detecte les lignes TOTAL du type:
    ["", "AC1", "30,522,402", "23,494,299", "7,028,103"]
    """
    for cell in row:
        cell_clean = str(cell).strip().upper()
        if re.match(r"^AC\d+$", cell_clean):
            return cell_clean
    return None


def extract_title_code(designation):
    """
    Extrait AC1 depuis "AC1 Actifs incorporels"
    """
    if designation is None:
        return None
    m = re.match(r"^(AC\d+)\b", str(designation).strip().upper())
    return m.group(1) if m else None


def extract_actif(pdf_path, page_num, is_scanned=False):

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
                    norm = normalize_columns(row.tolist())
                    norm = shift_if_needed(norm)
                    norm = fix_ac_header_shift(norm)
                    rows.append(norm)

        if not rows:
            print(":x: Aucune ligne ACTIF détectée")
            return None

        df = pd.DataFrame(rows, columns=["DESIGNATION", "BRUT", "AMORT_PROV", "NET_N", "NET_N1"])

        # =============================
        # Detecter lignes TOTAL (AC1 seul)
        # =============================
        df["TOTAL_CODE"] = df.apply(lambda r: extract_total_code(r.tolist()), axis=1)

        total_rows = df[df["TOTAL_CODE"].notna()].copy()
        normal_rows = df[df["TOTAL_CODE"].isna()].copy()

        # =============================
        # Transfert valeurs total vers ligne titre
        # =============================
        normal_rows["TITLE_CODE"] = normal_rows["DESIGNATION"].apply(extract_title_code)

        for _, trow in total_rows.iterrows():
            code = trow["TOTAL_CODE"]

            # trouver ligne titre correspondante (AC1 Actifs incorporels)
            mask = (normal_rows["TITLE_CODE"] == code)

            if mask.any():
                idx = normal_rows[mask].index[0]

                # transférer valeurs seulement si la ligne titre est vide
                if (
                    normal_rows.at[idx, "BRUT"] == "" and
                    normal_rows.at[idx, "AMORT_PROV"] == "" and
                    normal_rows.at[idx, "NET_N"] == "" and
                    normal_rows.at[idx, "NET_N1"] == ""
                ):
                    normal_rows.at[idx, "BRUT"] = trow["BRUT"]
                    normal_rows.at[idx, "AMORT_PROV"] = trow["AMORT_PROV"]
                    normal_rows.at[idx, "NET_N"] = trow["NET_N"]
                    normal_rows.at[idx, "NET_N1"] = trow["NET_N1"]

        # supprimer colonne helper
        normal_rows = normal_rows.drop(columns=["TITLE_CODE"])

        df_final = normal_rows.drop(columns=["TOTAL_CODE"])

        print(f":white_check_mark: {len(df_final)} lignes ACTIF extraites")

        return df_final.to_dict("records")

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