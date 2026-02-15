# B.py
# ------------------------------------------------------------
# Annexe 13 (COMAR) - Normalisation + validation UNIQUEMENT via :
#   C1 = TOTAL - somme(des autres colonnes)
# - Conserve le style du tableau extrait (openpyxl, modifications "in place")
# - Ajoute une colonne C1 juste après TOTAL (si absente)
# - Met en ORANGE la cellule C1 si la ligne est invalide
# - Met en ROUGE la cellule la plus "contributrice" (|valeur| max) sur la ligne invalide
# - Boucle Excel: ouvre le fichier, attend Ctrl+S (via modification mtime), ferme Excel, revalide
# - Option: forcer Excel 2010 via EXCEL2010_PATH (si tu veux)
# ------------------------------------------------------------

import os
import re
import sys
import time
import copy
import subprocess
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# =========================
# CONFIG
# =========================

EXPECTED_COLUMNS = [
    "CATEGORIES", "INCENDIE", "A.TRAVAIL", "RC", "AUTOMOBILE", "TRANSPORT", "GROUPE",
    "DOMMAGES AUX BIENS", "RISQUES AGRICOLES", "CONSTRUCTION", "PERTE D'EXPLOITATION",
    "CAUTION", "ASSISTANCE", "A.CORPOREL", "ACCEPTATION", "TOTAL"
]

EXPECTED_ROWS = [
    "PRIMES EMISES", "VARIATION DES PRIMES NON ACQUISES",
    "PRESTATIONS ET FRAIS PAYES",
    "CHARGES DES PROVISIONS POUR PRESTATIONS DIVERSE",
    "SOLDE DE SOUSCRIPTION",
    "FRAIS D'ACQUISITION", "AUTRES CHARGES DE GESTION NETTES",
    "CHARGES D'ACQUISITION ET DE GESTION NETTES",
    "PRODUITS NETS DE PLACEMENTS", "AUTRE PRODUITS TECHNIQUES",
    "SOLDE FINANCIER",
    "PART REASSUREURS DANS LES PRIMES ACQUISES",
    "PART REASSUREURS DANS LES PRIMES NON ACQUISES",
    "PART REASSUREURS DANS LES PRESTATIONS PAYEES",
    "PART REASSUREURS DANS LES CHARGES DE PROVISIONS",
    "COMMISSIONS REÇUES DES REASSUREURS",
    "PART REASSUREURS DANS LA PARTICIPATION AUX RESULTATS",
    "PART REASSUREURS DANS LES FRAIS REPORTES",
    "SOLDE DE REASSURANCE",
    "RESULTAT TECHNIQUE NON VIE",
    "INFORMATIONS COMPLEMENTAIRES",
    "PROVISIONS POUR PRIMES NON ACQUISES - ANNEE N",
    "PROVISIONS POUR PRIMES NON ACQUISES - ANNEE N-1",
    "PROVISIONS POUR SINSITRES A PAYER - ANNEE N",
    "PROVISIONS POUR SINSITRES A PAYER - ANNEE N-1",
    "PREVISIONS DE RECOURS A ENCAISSER - ANNEE N",
    "PREVISIONS DE RECOURS A ENCAISSER - ANNEE N-1",
    "PROVISIONS POUR PARTICIPATIONS AUX BENEFICES - ANNEE N",
    "PROVISIONS POUR PARTICIPATIONS AUX BENEFICES - ANNEE N-1",
    "PROVISIONS POUR EGALISATION ET EQUILIBRAGE - ANNEE N",
    "PROVISIONS POUR EGALISATION ET EQUILIBRAGE - ANNEE N-1",
    "PROVISIONS MATHEMATIQUES DE RENTE - ANNEE N",
    "PROVISIONS MATHEMATIQUES DE RENTE - ANNEE N-1",
    "PROVISIONS POUR RISQUES EN COURS - ANNEE N",
    "PROVISIONS POUR RISQUES EN COURS - ANNEE N-1",
]

# Tolérance sur C1 (différences de somme)
TOL = 5.0  # ajuste si besoin

# Couleurs (fill)
FILL_RED = PatternFill("solid", fgColor="FF0000")
FILL_ORANGE = PatternFill("solid", fgColor="FFA500")
FILL_GREEN = PatternFill("solid", fgColor="00B050")  # vert Excel
FILL_NONE = PatternFill(fill_type=None)

# Excel 2010 (optionnel) : mets ici le chemin EXACT si tu veux forcer Office 2010
# Exemple:
# EXCEL2010_PATH = r"C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE"
EXCEL2010_PATH = None


# =========================
# UTILS: texte / matching
# =========================

def _norm_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().upper()
    # accents simples
    s = (
        s.replace("É", "E").replace("È", "E").replace("Ê", "E")
        .replace("À", "A").replace("Â", "A")
        .replace("Î", "I").replace("Ï", "I")
        .replace("Ô", "O")
        .replace("Ù", "U")
        .replace("Ç", "C")
        .replace("’", "'").replace("`", "'")
    )
    # espaces multiples
    s = re.sub(r"\s+", " ", s)
    # nettoyage ponctuation non utile
    s = re.sub(r"[•\u2010\u2011\u2012\u2013\u2014\u2212]", "-", s)  # tirets
    return s


def _similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio() * 100.0


def best_match(value: str, expected: List[str], min_score: float = 78.0) -> Tuple[Optional[str], float]:
    k = _norm_key(value)
    best = None
    best_score = -1.0
    for e in expected:
        sc = _similarity(k, _norm_key(e))
        if sc > best_score:
            best_score = sc
            best = e
    if best_score >= min_score:
        return best, best_score
    return None, best_score


# =========================
# UTILS: nombres
# =========================

_NUM_RE = re.compile(r"[-+]?\d[\d\s.,]*")

def parse_number(x) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return None

    # extraire la première séquence num-like
    m = _NUM_RE.search(s.replace("\u00A0", " "))
    if not m:
        return None
    token = m.group(0)

    # enlever espaces
    token = token.replace(" ", "")

    # gérer virgule/point
    # si contient , et . => on suppose . = milliers, , = décimal (ou l'inverse) : on prend dernière occurrence comme décimal
    if "," in token and "." in token:
        if token.rfind(",") > token.rfind("."):
            token = token.replace(".", "")
            token = token.replace(",", ".")
        else:
            token = token.replace(",", "")
    else:
        # si juste virgule -> décimal
        if "," in token and "." not in token:
            token = token.replace(",", ".")
        # si juste point -> décimal
        # sinon ok

    try:
        return float(token)
    except:
        return None


# =========================
# UTILS: styles (évite StyleProxy unhashable)
# =========================

import copy

def copy_cell_style(dst, src):
    """
    Copie le style d'une cellule en cassant les StyleProxy (sinon erreurs unhashable / immutable).
    """
    dst._style = copy.copy(src._style)
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy.copy(src.protection)
    dst.comment = src.comment



def ensure_c1_column_inplace(ws, header_row: int = 1):
    # trouve TOTAL
    headers = {}
    max_col = ws.max_column
    for c in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        headers[_norm_key(v)] = c

    if "C1" in headers:
        return  # déjà là

    total_col = headers.get("TOTAL")
    if not total_col:
        # tente "Total"
        total_col = headers.get("TOTAL ") or headers.get("TOTAL:")
    if not total_col:
        raise ValueError("Colonne TOTAL introuvable (ligne d'entête).")

    insert_at = total_col + 1
    ws.insert_cols(insert_at)

    # header C1 en copiant style de TOTAL header
    src_header = ws.cell(row=header_row, column=total_col)
    dst_header = ws.cell(row=header_row, column=insert_at)
    dst_header.value = "C1"
    copy_cell_style(dst_header, src_header)

    # copier largeur
    total_letter = get_column_letter(total_col)
    c1_letter = get_column_letter(insert_at)
    ws.column_dimensions[c1_letter].width = ws.column_dimensions[total_letter].width

    # copier styles des cellules de la colonne TOTAL vers C1 (même format)
    for r in range(header_row + 1, ws.max_row + 1):
        src = ws.cell(row=r, column=total_col)
        dst = ws.cell(row=r, column=insert_at)
        copy_cell_style(dst, src)


# =========================
# Normalisation colonnes/lignes (sans casser le style)
# =========================

def normalize_columns_inplace(ws, header_row: int = 1):
    """
    Normalise les noms de colonnes pour correspondre à EXPECTED_COLUMNS.
    Règles fixes demandées :
      - COL_12 -> CONSTRUCTION
      - COL_16 -> ASSISTANCE
    + mapping de variantes fréquentes (COMAR)
    """
    # mapping direct (clé normalisée -> cible)
    direct_map = {
        _norm_key("Incendie"): "INCENDIE",
        _norm_key("Accident Travail"): "A.TRAVAIL",
        _norm_key("A.TRAVAIL"): "A.TRAVAIL",
        _norm_key("RC"): "RC",
        _norm_key("Automobile"): "AUTOMOBILE",
        _norm_key("Transport"): "TRANSPORT",
        _norm_key("Groupe"): "GROUPE",
        _norm_key("Biens"): "DOMMAGES AUX BIENS",
        _norm_key("Dommages aux biens"): "DOMMAGES AUX BIENS",
        _norm_key("Risques Agricoles"): "RISQUES AGRICOLES",
        _norm_key("d'Eploitation"): "PERTE D'EXPLOITATION",
        _norm_key("d'Exploitation"): "PERTE D'EXPLOITATION",
        _norm_key("Perte d'exploitation"): "PERTE D'EXPLOITATION",
        _norm_key("Caution"): "CAUTION",
        _norm_key("Corporel"): "A.CORPOREL",
        _norm_key("Accident corporel"): "A.CORPOREL",
        _norm_key("Acceptation"): "ACCEPTATION",
        _norm_key("Total"): "TOTAL",
        _norm_key("TOTAL"): "TOTAL",
        _norm_key("CATEGORIES"): "CATEGORIES",
        _norm_key("Categories"): "CATEGORIES",
        _norm_key("COL_12"): "CONSTRUCTION",
        _norm_key("COL_16"): "ASSISTANCE",
    }

    max_col = ws.max_column
    for c in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=c)
        if cell.value is None:
            continue

        k = _norm_key(cell.value)
        if k in direct_map:
            cell.value = direct_map[k]
            continue

        # fuzzy match vers EXPECTED_COLUMNS (si proche)
        match, score = best_match(str(cell.value), EXPECTED_COLUMNS, min_score=78.0)
        if match:
            cell.value = match


def normalize_rows_inplace(ws, category_col: int = 1, start_row: int = 2):
    """
    Normalise les libellés de lignes (colonne CATEGORIES) par fuzzy match.
    N'altère pas la structure / style.
    """
    for r in range(start_row, ws.max_row + 1):
        cell = ws.cell(row=r, column=category_col)
        if cell.value is None:
            continue
        raw = str(cell.value).strip()
        if not raw:
            continue

        match, score = best_match(raw, EXPECTED_ROWS, min_score=75.0)
        if match:
            cell.value = match


# =========================
# Auto-size colonnes (largeur)
# =========================

from openpyxl.utils import get_column_letter

def autosize_columns(ws, min_w: float = 10.0, max_w: float = 80.0, padding: float = 2.5):
    """
    Ajuste la largeur des colonnes selon le contenu.
    Donne plus de poids aux entêtes (ligne 1).
    """
    header_row = 1

    for col_cells in ws.columns:
        first_cell = col_cells[0]
        col_letter = get_column_letter(first_cell.column)

        max_len = 0

        # header (poids fort)
        hv = ws.cell(row=header_row, column=first_cell.column).value
        if hv is not None:
            hs = str(hv)
            hs = max(hs.splitlines(), key=len) if "\n" in hs else hs
            max_len = max(max_len, int(len(hs) * 1.35))

        # contenu
        for cell in col_cells:
            v = cell.value
            if v is None:
                continue
            s = str(v)
            s = max(s.splitlines(), key=len) if "\n" in s else s
            max_len = max(max_len, len(s))

        ws.column_dimensions[col_letter].width = max(min_w, min(max_w, max_len + padding))


def autosize_rows(ws, min_h: float = 15.0, max_h: float = 140.0):
    """
    Ajuste la hauteur des lignes selon le contenu (wrap).
    Heuristique simple et stable.
    """
    for r in range(1, ws.max_row + 1):
        max_lines = 1
        max_chars = 0

        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            s = str(v)
            lines = s.splitlines() if "\n" in s else [s]
            max_lines = max(max_lines, len(lines))
            max_chars = max(max_chars, max(len(x) for x in lines))

        h = 15.0 * max_lines
        if max_chars > 45:
            h += 6.0

        ws.row_dimensions[r].height = max(min_h, min(max_h, h))




# =========================
# Validation C1: TOTAL - somme(autres colonnes)
# =========================

@dataclass
class InvalidCell:
    excel_row: int
    c1_value: float



def _find_header_map(ws, header_row: int = 1) -> Dict[str, int]:
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        m[_norm_key(v)] = c
    return m


def clear_previous_greens(ws):
    """
    Compatibilité avec l'ancien nom.
    On efface les verts/oranges/rouges appliqués par le script,
    pour que le vert disparaisse au cycle suivant.
    """
    clear_previous_marks(ws)



import copy
from openpyxl.styles import PatternFill

FILL_NONE = PatternFill(fill_type=None)

def clear_previous_marks(ws):
    """
    Efface uniquement les couleurs qu'on applique pendant la validation des LIGNES (C1):
    - vert + orange
    (on évite d'effacer d'éventuels styles d'origine)
    """
    target_rgbs = {"FFA500", "00B050"}  # ORANGE, GREEN
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            fill = getattr(cell, "fill", None)
            rgb = getattr(getattr(fill, "fgColor", None), "rgb", None)
            if fill and fill.patternType == "solid" and rgb in target_rgbs:
                cell.fill = copy.copy(FILL_NONE)


import copy
from dataclasses import dataclass
from typing import List, Dict, Optional

TOL = 5.0  # tolérance [-5, +5]

FILL_ORANGE = PatternFill("solid", fgColor="FFA500")
FILL_GREEN  = PatternFill("solid", fgColor="00B050")

@dataclass
class InvalidCell:
    excel_row: int
    c1_value: float

def validate_c1_inplace(ws, header_row: int = 1, data_start_row: int = 2) -> List[InvalidCell]:
    """
    Validation UNIQUEMENT par lignes:
      C1 = TOTAL - somme(des autres colonnes)
    Coloration:
      - C1 vert si abs(C1) <= TOL
      - C1 orange si abs(C1) > TOL
    AUCUNE autre cellule n'est coloriée.
    """
    headers = _find_header_map(ws, header_row=header_row)
    cat_col = headers.get("CATEGORIES")
    total_col = headers.get("TOTAL")
    c1_col = headers.get("C1")

    if not cat_col or not total_col or not c1_col:
        raise ValueError("Il faut CATEGORIES, TOTAL et C1 dans l'entête.")

    # colonnes numériques = toutes sauf CATEGORIES et C1
    numeric_cols = [c for k, c in headers.items() if k not in ("CATEGORIES", "C1")]

    invalids: List[InvalidCell] = []

    for r in range(data_start_row, ws.max_row + 1):
        cat = ws.cell(row=r, column=cat_col).value
        if cat is None or str(cat).strip() == "":
            continue

        total_val = parse_number(ws.cell(row=r, column=total_col).value)
        if total_val is None:
            continue

        row_sum = 0.0
        for c in numeric_cols:
            if c == total_col:
                continue
            v = parse_number(ws.cell(row=r, column=c).value)
            if v is None:
                continue
            row_sum += v

        c1 = total_val - row_sum

        c1_cell = ws.cell(row=r, column=c1_col)
        c1_cell.value = c1

        if abs(c1) <= TOL:
            c1_cell.fill = copy.copy(FILL_GREEN)
        else:
            c1_cell.fill = copy.copy(FILL_ORANGE)
            invalids.append(InvalidCell(excel_row=r, c1_value=c1))

    return invalids


import copy

def format_header_row(ws, header_row: int = 1, min_h: float = 34.0):
    """
    Agrandit la ligne des entêtes + wrap_text SANS modifier un objet style immutable.
    """
    current_h = ws.row_dimensions[header_row].height or 0
    ws.row_dimensions[header_row].height = max(current_h, min_h)

    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=c)
        if cell.value is None:
            continue

        # IMPORTANT: openpyxl styles sont immutables => on réassigne une COPIE
        if cell.alignment is not None:
            al = copy.copy(cell.alignment)
        else:
            from openpyxl.styles import Alignment
            al = Alignment()

        al.wrap_text = True
        al.vertical = al.vertical or "center"
        al.horizontal = al.horizontal or "center"
        cell.alignment = al




# =========================
# Excel loop (CTRL+S)
# =========================

def _open_excel_2010_or_default(file_path: str):
    """
    Ouvre le fichier dans Excel (idéalement 2010).
    Si EXCEL2010_PATH est défini, on lance cet exe.
    Sinon on tente via COM standard (Excel.Application).
    """
    try:
        import win32com.client  # pywin32
        import pythoncom
    except Exception:
        win32com = None
        pythoncom = None

    if EXCEL2010_PATH and os.path.exists(EXCEL2010_PATH):
        # Lance Excel 2010 directement
        subprocess.Popen([EXCEL2010_PATH, file_path], close_fds=True)
        return None  # pas de handle COM
    else:
        if win32com is None:
            # fallback: ouverture simple (explorer)
            os.startfile(file_path)
            return None

        pythoncom.CoInitialize()
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = True
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(os.path.abspath(file_path))
        return (xl, wb)


def _close_excel_handle(handle):
    if not handle:
        return
    xl, wb = handle
    try:
        wb.Close(SaveChanges=True)
    except:
        pass
    try:
        xl.Quit()
    except:
        pass


def wait_for_ctrl_s(file_path: str, poll: float = 0.8, timeout_sec: Optional[int] = None) -> bool:
    """
    On attend que le fichier soit sauvegardé (= mtime change).
    """
    try:
        last_mtime = os.path.getmtime(file_path)
    except:
        last_mtime = None

    start = time.time()
    while True:
        time.sleep(poll)

        if timeout_sec is not None and (time.time() - start) > timeout_sec:
            return False

        try:
            cur = os.path.getmtime(file_path)
        except:
            continue

        if last_mtime is None:
            last_mtime = cur
            continue

        if cur != last_mtime:
            return True


def save_with_retries(wb, path: str, tries: int = 10, sleep_sec: float = 0.6):
    last_err = None
    for _ in range(tries):
        try:
            wb.save(path)
            return
        except Exception as e:
            last_err = e
            time.sleep(sleep_sec)
    raise last_err


# =========================
# Pipeline Annexe 13
# =========================

def normalize_excel_annexe13_keep_style(in_path: str, out_path: str) -> str:
    wb = openpyxl.load_workbook(in_path)
    ws = wb.active  # tableau principal

    # 1) normaliser colonnes / lignes
    normalize_columns_inplace(ws, header_row=1)

    # detecter col CATEGORIES
    headers = _find_header_map(ws, header_row=1)
    cat_col = headers.get("CATEGORIES", 1)

    normalize_rows_inplace(ws, category_col=cat_col, start_row=2)

    # 2) assurer C1
    ensure_c1_column_inplace(ws, header_row=1)

    # 3) auto-size
    autosize_columns(ws)
    format_header_row(ws, header_row=1, min_h=28.0)
    autosize_columns(ws)

    save_with_retries(wb, out_path)
    return out_path


def validate_excel_loop_annexe13_keep_style(xlsx_path: str) -> int:
    """
    Boucle:
      - ouvre workbook
      - efface les marques (vert/orange/rouge) du cycle précédent
      - calcule C1 partout + coloration
      - si invalide => ouvre Excel et attend Ctrl+S, ferme, relance validation
      - si valide => termine
    """
    while True:
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb.active

        # effacer les couleurs appliquées par le script (vert/orange/rouge)
        clear_previous_marks(ws)

        # recalcul + coloration
        invalids = validate_c1_inplace(ws, header_row=1, data_start_row=2)

        # auto-size (largeur + hauteur)
        autosize_columns(ws)
        autosize_rows(ws)

        # sauver
        save_with_retries(wb, xlsx_path)

        # libérer le fichier (sinon Permission denied)
        wb.close()
        del wb

        if not invalids:
            print("STATUT: Valide ✅ (Annexe 13)")
            return 0

        print(f"STATUT: Invalide ❌ (Annexe 13) | lignes invalides = {len(invalids)} | Excel rows: {[x.excel_row for x in invalids]}")
        print("🟡 Corrige dans Excel puis fais Ctrl+S (ensuite reviens ici).")

        # ouvrir excel + attendre ctrl+s
        handle = None
        try:
            handle = _open_excel_2010_or_default(xlsx_path)
            ok = wait_for_ctrl_s(xlsx_path, poll=0.8)
            if ok:
                print("CTRL+S détecté, fermeture du fichier Excel...")
            _close_excel_handle(handle)
        except Exception:
            try:
                _close_excel_handle(handle)
            except:
                pass
            try:
                subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print("Excel fermé via terminaison du processus EXCEL.EXE.")
            except:
                pass

        print("🔁 Relance validation...")


# =========================
# MAIN
# =========================

def _make_default_out_path(in_path: str) -> str:
    """
    Si l’utilisateur lance: py B.py "13E2024.xlsx"
    -> on travaille sur le même fichier (in-place) pour garder style + boucle.
    Mais tu peux décider de sortir un NV séparé au début.
    Ici: si le nom contient '13E2024' -> on produit '13NV2024.xlsx' une seule fois,
    puis on boucle sur ce NV.
    """
    folder = os.path.dirname(os.path.abspath(in_path))
    base = os.path.basename(in_path)
    name, ext = os.path.splitext(base)

    if "13E2024" in name.upper():
        return os.path.join(folder, "13NV2024.xlsx")
    if name.upper().endswith("E2024"):
        return os.path.join(folder, name[:-5] + "NV2024.xlsx")
    # fallback
    return os.path.join(folder, f"{name}_NV.xlsx")


def main() -> int:
    if len(sys.argv) < 2:
        print('Usage: py B.py "13E2024.xlsx"')
        return 2

    in_path = sys.argv[1]
    if not os.path.isabs(in_path):
        in_path = os.path.abspath(in_path)

    if not os.path.exists(in_path):
        print(f"Fichier introuvable: {in_path}")
        return 2

    out_path = _make_default_out_path(in_path)

    # 1) normalisation (copie style)
    if os.path.abspath(out_path) != os.path.abspath(in_path):
        print(f"➡️ Normalisation Annexe 13: {in_path}")
        out_path = normalize_excel_annexe13_keep_style(in_path, out_path)
        print(f"✅ Fichier normalisé: {out_path}")
    else:
        # Si tu veux vraiment travailler in-place sans créer NV:
        # assure-toi quand même que colonnes/lignes/C1 sont prêts
        print(f"➡️ Normalisation Annexe 13 (in-place): {in_path}")
        normalize_excel_annexe13_keep_style(in_path, in_path)
        out_path = in_path
        print(f"✅ Fichier normalisé: {out_path}")

    # 2) boucle validation C1
    return validate_excel_loop_annexe13_keep_style(out_path)


if __name__ == "__main__":
    raise SystemExit(main())
