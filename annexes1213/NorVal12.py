# -*- coding: utf-8 -*-
"""
C.py - Normalisation + Validation Annexe 12 (COMAR) (post-extraction A.py)

- Input:  Annexe_12_..._extracted.xlsx (produit par A.py)
- Output: output_<input>.xlsx (même dossier)
- Normalisation:
    * 1ère colonne -> CATEGORIES
    * Colonnes montants -> VIE, TOTAL (si possible)
    * Nettoyage cellules numériques: garder uniquement le montant (chiffres)
    * Ne JAMAIS ajouter des lignes "VIE" / "TOTAL" (et les supprimer si présentes)
- Validation (boucle):
    * Pour chaque ligne: C1 = TOTAL - VIE
    * Si C1 != 0: cellule VIE en rouge
    * Tant qu'il existe du rouge: on attend un Ctrl+S (modif du fichier) puis on revalide.
"""

   
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import re
import sys
import time
import logging
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ----------------------------- CONFIG -----------------------------
# Tes libellés canoniques COMAR (Annexe 12)
CANONICAL_ROWS = [
    "PRIMES EMISES",
    "PRESTATIONS ET FRAIS PAYES",
    "CHARGES DES PROVISIONS POUR PRESTATIONS DIVERSES",
    "SOLDE DE SOUSCRIPTION",
    "FRAIS D'ACQUISITION",
    "AUTRES CHARGES DE GESTION NETTES",
    "CHARGES D'ACQUISITION ET DE GESTION NETTES",
    "PRODUITS NETS DE PLACEMENTS",
    "PARTICIPATION AUX RESULTATS",
    "SOLDE FINANCIER",
    "PART DES REASSUREURS DANS LES PRIMES ACQUISES",
    "PART DES REASSUREURS DANS LES PRESTATIONS PAYEES",
    "PART DES REASSUREURS DANS LES CHARGES DE PROVISIONS",
    "COMMISSIONS RECUES DES REASSUREURS /RETROCESSIONNAIRES",
    "SOLDE DE REASSURANCE / RETROCESSION",
    "RESULTAT TECHNIQUE",
    "INFORMATIONS COMPLEMENTAIRES",
    "PROVISIONS POUR SINISTRES A PAYER - ANNEE N",
    "PROVISIONS POUR SINISTRES A PAYER - ANNEE N-1",
    "PROVISIONS POUR PARTICIPATIONS AUX BENEFICES - ANNEE N",
    "PROVISIONS POUR PARTICIPATIONS AUX BENEFICES - ANNEE N-1",
    "PROVISIONS POUR EGALISATION ET EQUILIBRAGE - ANNEE N",
    "PROVISIONS POUR EGALISATION ET EQUILIBRAGE - ANNEE N-1",
    "PROVISIONS MATHEMATIQUES VIE - ANNEE N",
    "PROVISIONS MATHEMATIQUES VIE - ANNEE N-1",
]

CANONICAL_COLS = ["VIE", "TOTAL"]
MATCH_THRESHOLD = 0.78

HEADER_COLOR = "0070C0"

logging.basicConfig(
    filename="C_script.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


# ----------------------------- UTILITAIRES -----------------------------
def _now_ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _safe_save_path(path: str) -> str:
    """
    Si le fichier est ouvert (PermissionError), sauvegarde sous un autre nom avec timestamp.
    """
    base, ext = os.path.splitext(path)
    try:
        with open(path, "ab"):
            pass
        return path
    except Exception:
        return f"{base}_{_now_ts()}{ext}"


def _normalize_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[’`´]", "'", s)
    s = re.sub(r"[^A-Z0-9\s'\-\/]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _similar(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def _extract_number_only(x):
    """
    Garde uniquement le montant:
    - Supporte: espaces, NBSP, virgules, points, parenthèses, signes -, etc.
    - Si pas de nombre -> "" (vide)
    """
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        try:
            if isinstance(x, float) and abs(x - int(x)) < 1e-9:
                return int(x)
            return x
        except Exception:
            return x

    s = str(x).strip()
    if s == "":
        return ""

    s = s.replace("\u00a0", " ").replace(" ", "")

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    s = s.replace("−", "-").replace("–", "-").replace("—", "-").replace("‐", "-")

    m = re.search(r"-?\d[\d\.,]*", s)
    if not m:
        return ""

    num = m.group(0)

    # normaliser séparateurs
    # si beaucoup de points/virgules: on enlève tout sauf chiffres et dernier séparateur
    num = num.replace(",", ".")
    # garder chiffres et points
    num = re.sub(r"[^0-9\.\-]", "", num)

    # si plusieurs points -> c'est sûrement des milliers => enlever tous les points
    if num.count(".") > 1:
        num = num.replace(".", "")

    try:
        if num in ("", "-", "."):
            return ""
        val = float(num)
        if negative:
            val = -val
        # COMAR: montants entiers
        return int(round(val))
    except Exception:
        return ""


def _to_number_or_none(x):
    if x is None:
        return None
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return float(x)
    s = str(x).strip()
    if not s:
        return None
    v = _extract_number_only(s)
    if v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None

def _open_in_excel(path: str):
    try:
        os.startfile(os.path.abspath(path))
        return True
    except Exception:
        try:
            import subprocess
            subprocess.Popen(['cmd', '/c', 'start', '', os.path.abspath(path)], shell=False)
            return True
        except Exception:
            return False



def _wait_for_ctrl_s(file_path: str, last_mtime: float, poll_sec: float = 1.0):
    print("Corrige dans Excel puis fais Ctrl+S (puis reviens ici).")
    while True:
        time.sleep(poll_sec)
        try:
            mt = os.path.getmtime(file_path)
            if mt > last_mtime:
                return
        except Exception:
            pass


# ----------------------------- NORMALISATION -----------------------------
def _choose_amount_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Objectif: ne garder que 2 colonnes montants -> VIE et TOTAL
    On tente de détecter automatiquement:
      - une colonne qui ressemble à "VIE"
      - une colonne qui ressemble à "TOTAL"
    Fallback:
      - prendre les 2 dernières colonnes non CATEGORIES.
    """
    df = df.copy()
    cols = list(df.columns)

    if "CATEGORIES" not in cols:
        df = df.rename(columns={cols[0]: "CATEGORIES"})
        cols = list(df.columns)

    non_cat = [c for c in cols if c != "CATEGORIES"]
    if not non_cat:
        df["VIE"] = ""
        df["TOTAL"] = ""
        return df[["CATEGORIES", "VIE", "TOTAL"]]

    norm_map = {c: _normalize_key(c) for c in non_cat}

    vie_candidates = [c for c in non_cat if "VIE" == norm_map[c] or norm_map[c].endswith(" VIE") or "VIE" in norm_map[c]]
    total_candidates = [c for c in non_cat if "TOTAL" in norm_map[c]]

    vie_col = vie_candidates[0] if vie_candidates else None
    total_col = total_candidates[0] if total_candidates else None

    # fallback: prendre dernières colonnes
    if vie_col is None or total_col is None or vie_col == total_col:
        if len(non_cat) >= 2:
            vie_col = non_cat[-2]
            total_col = non_cat[-1]
        elif len(non_cat) == 1:
            vie_col = non_cat[0]
            total_col = non_cat[0]
        else:
            vie_col, total_col = None, None

    out = pd.DataFrame()
    out["CATEGORIES"] = df["CATEGORIES"]
    out["VIE"] = df[vie_col] if vie_col in df.columns else ""
    out["TOTAL"] = df[total_col] if total_col in df.columns else ""
    return out


def _clean_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ["VIE", "TOTAL"]:
        if c in df.columns:
            df[c] = df[c].apply(_extract_number_only)
            df[c] = df[c].replace("", pd.NA)
    return df


def _normalize_rows_to_canonical(df: pd.DataFrame) -> pd.DataFrame:
    """
    STRICT: la sortie contient UNIQUEMENT les lignes CANONICAL_ROWS (dans cet ordre),
    une seule fois chacune. Aucune ligne extra. Pas de doublons.
    """
    if df is None or df.empty:
        return pd.DataFrame({"CATEGORIES": CANONICAL_ROWS, "VIE": [pd.NA]*len(CANONICAL_ROWS), "TOTAL": [pd.NA]*len(CANONICAL_ROWS)})

    df = df.copy()
    df["CATEGORIES"] = df["CATEGORIES"].astype(str).map(lambda x: x.strip())

    # enlever lignes "VIE"/"TOTAL" si présentes (exact)
    key = df["CATEGORIES"].map(_normalize_key)
    df = df.loc[~key.isin(["VIE", "TOTAL"])].reset_index(drop=True)

    # indexer lignes existantes
    existing = []
    for idx, raw in enumerate(df["CATEGORIES"].tolist()):
        k = _normalize_key(raw)
        if k:
            existing.append((idx, raw, k))

    used_idx = set()
    out_rows = []

    for target in CANONICAL_ROWS:
        tkey = _normalize_key(target)
        best_idx = None
        best_score = -1.0

        for idx, raw, k in existing:
            if idx in used_idx:
                continue
            sc = _similar(tkey, k)
            if sc > best_score:
                best_score = sc
                best_idx = idx

        if best_idx is not None and best_score >= MATCH_THRESHOLD:
            used_idx.add(best_idx)
            row = df.loc[best_idx].copy()
            row["CATEGORIES"] = target
            out_rows.append(row)
        else:
            out_rows.append(pd.Series({"CATEGORIES": target, "VIE": pd.NA, "TOTAL": pd.NA}))

    out = pd.DataFrame(out_rows).reset_index(drop=True)
    return out


def _get_col_index(ws, header_name: str):
    hn = str(header_name).strip().upper()
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is None:
            continue
        if str(v).strip().upper() == hn:
            return c
    return None


def _delete_rows_by_label(ws, col_cat: int):
    """
    Supprime dans la feuille les lignes dont CATEGORIES est exactement VIE ou TOTAL.
    """
    to_del = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, col_cat).value
        if _normalize_key(v) in ("VIE", "TOTAL"):
            to_del.append(r)
    for r in reversed(to_del):
        ws.delete_rows(r, 1)


def _ensure_c1_column_right_after_total(ws, col_total: int) -> int:
    """
    Place C1 juste après TOTAL et applique un style type B.py:
      - header bleu + texte blanc
      - bordures
      - align center
      - number format "0"
    """
    existing_c1 = _get_col_index(ws, "C1")
    insert_at = col_total + 1

    if existing_c1 is not None and existing_c1 != insert_at:
        ws.delete_cols(existing_c1, 1)
        ws.insert_cols(insert_at, 1)
    elif existing_c1 is None:
        ws.insert_cols(insert_at, 1)

    # header
    h = ws.cell(1, insert_at)
    h.value = "C1"

    header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    h.fill = header_fill
    h.font = header_font
    h.border = border
    h.alignment = center

    # style colonne C1 (toutes lignes)
    for r in range(2, ws.max_row + 1):
        c = ws.cell(r, insert_at)
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.number_format = "0"

    # largeur type B.py
    ws.column_dimensions[get_column_letter(insert_at)].width = 12

    return insert_at

def _close_excel():
    """
    Ferme toutes les instances Excel ouvertes (Windows).
    """
    try:
        import subprocess
        subprocess.call(["taskkill", "/F", "/IM", "EXCEL.EXE"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.5)
    except Exception:
        pass


def validate_excel_loop_vie_total(out_path: str, annexe_num: str = "12") -> int:
    """
    Validation demandée:
      - Colonne C1 insérée juste après TOTAL
      - C1 = TOTAL - VIE
      - si C1 != 0 :
            * VIE rouge
            * C1 orange
        sinon :
            * si la cellule VIE était rouge au cycle précédent -> vert sur CE cycle
            * sinon -> normal (sans couleur)
      - Nettoyage auto: lettres/symboles supprimés (on garde uniquement le nombre)
      - Boucle: tant que invalide -> ouvrir Excel 2010, attendre Ctrl+S, fermer l'Excel lancé, revalider
    """
    import os
    import re
    import time
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    THIN = Side(style="thin")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

    RED_FILL = PatternFill(start_color="FF4040", end_color="FF4040", fill_type="solid")
    ORANGE_FILL = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    GREEN_FILL = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    NO_FILL = PatternFill(fill_type=None)

    WHITE_BOLD = Font(color="FFFFFF", bold=True)
    BLACK_NORMAL = Font(color="000000", bold=False)

    HEADER_FILL = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    HEADER_FONT = Font(color="FFFFFF", bold=True)

    def _norm(x):
        return (str(x).strip().upper() if x is not None else "")

    def _get_col(ws, name: str):
        target = _norm(name)
        for c in range(1, ws.max_column + 1):
            if _norm(ws.cell(1, c).value) == target:
                return c
        return None

    def _to_num_aggressive(x):
        """
        Nettoyage agressif: garde uniquement le montant numérique.
        Ex: "12 300 DT" -> 12300 ; "(1 000)" -> -1000 ; "abc" -> None
        """
        if x is None:
            return None
        if isinstance(x, (int, float)) and not isinstance(x, bool):
            return float(x)

        s = str(x).strip()
        if not s:
            return None

        s = s.replace("\u00a0", " ").replace(" ", "")
        s = s.replace("−", "-").replace("–", "-").replace("—", "-").replace("‐", "-")

        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1]

        m = re.search(r"-?\d[\d\.,]*", s)
        if not m:
            return None

        num = m.group(0).replace(",", ".")
        num = re.sub(r"[^0-9\.\-]", "", num)

        if num.count(".") > 1:
            num = num.replace(".", "")

        try:
            v = float(num)
            return -v if neg else v
        except Exception:
            return None

    def _file_writable(path: str) -> bool:
        try:
            with open(path, "a+b"):
                pass
            return True
        except Exception:
            return False

    def _wait_unlock(path: str, timeout: int = 60) -> bool:
        t0 = time.time()
        while time.time() - t0 < timeout:
            if _file_writable(path):
                return True
            time.sleep(0.8)
        return False

    def _wait_for_ctrl_s(path: str, last_mtime: float):
        print("🟡 Corrige dans Excel puis fais Ctrl+S (ensuite reviens ici).")
        while True:
            time.sleep(1.0)
            try:
                mt = os.path.getmtime(path)
                if mt > last_mtime:
                    return
            except Exception:
                pass

    def _delete_rows_vie_total(ws, col_cat: int):
        to_del = []
        for r in range(2, ws.max_row + 1):
            if _norm(ws.cell(r, col_cat).value) in ("VIE", "TOTAL"):
                to_del.append(r)
        for r in reversed(to_del):
            ws.delete_rows(r, 1)

    def _ensure_c1_after_total(ws, col_total: int) -> int:
        existing_c1 = _get_col(ws, "C1")
        insert_at = col_total + 1

        if existing_c1 is not None and existing_c1 != insert_at:
            ws.delete_cols(existing_c1, 1)
            ws.insert_cols(insert_at, 1)
        elif existing_c1 is None:
            ws.insert_cols(insert_at, 1)

        h = ws.cell(1, insert_at)
        h.value = "C1"
        h.fill = HEADER_FILL
        h.font = HEADER_FONT
        h.alignment = CENTER
        h.border = BORDER

        ws.column_dimensions[get_column_letter(insert_at)].width = 12
        return insert_at

    if not os.path.exists(out_path):
        raise FileNotFoundError(out_path)

    while True:
        _wait_unlock(out_path, timeout=60)

        wb = load_workbook(out_path)
        ws = wb.active

        col_cat = _get_col(ws, "CATEGORIES") or 1
        col_vie = _get_col(ws, "VIE")
        col_total = _get_col(ws, "TOTAL")

        if col_vie is None or col_total is None:
            wb.close()
            raise ValueError("Colonnes VIE/TOTAL introuvables.")

        _delete_rows_vie_total(ws, col_cat)

        col_cat = _get_col(ws, "CATEGORIES") or 1
        col_vie = _get_col(ws, "VIE")
        col_total = _get_col(ws, "TOTAL")

        col_c1 = _ensure_c1_after_total(ws, col_total)

        invalid_rows = []
        invalid_count = 0

        for r in range(2, ws.max_row + 1):
            cat = ws.cell(r, col_cat).value
            if cat is None or str(cat).strip() == "":
                continue

            vie_cell = ws.cell(r, col_vie)
            tot_cell = ws.cell(r, col_total)
            c1_cell = ws.cell(r, col_c1)

            # was red last cycle?
            was_red = (
                vie_cell.fill is not None
                and getattr(vie_cell.fill, "fill_type", None) == "solid"
                and getattr(vie_cell.fill, "start_color", None) is not None
                and str(vie_cell.fill.start_color.rgb).upper().endswith("FF4040")
            )

            # ✅ nettoyage auto (supprime lettres/symboles)
            vie = _to_num_aggressive(vie_cell.value)
            tot = _to_num_aggressive(tot_cell.value)

            # écrire les valeurs nettoyées dans la feuille
            # (comme ça si c'est juste "DT" ou autre, ça se corrige automatiquement)
            vie_cell.value = int(round(vie)) if vie is not None else None
            tot_cell.value = int(round(tot)) if tot is not None else None
            vie_num = float(vie_cell.value or 0)
            tot_num = float(tot_cell.value or 0)

            c1 = int(round(tot_num - vie_num))

            # styles/align
            ws.cell(r, col_cat).alignment = LEFT
            ws.cell(r, col_cat).border = BORDER

            for c in (col_vie, col_total, col_c1):
                ws.cell(r, c).alignment = CENTER
                ws.cell(r, c).border = BORDER

            c1_cell.value = c1
            c1_cell.number_format = "0"

            if c1 != 0:
                invalid_count += 1
                invalid_rows.append(r)

                vie_cell.fill = RED_FILL
                vie_cell.font = WHITE_BOLD

                c1_cell.fill = ORANGE_FILL
                c1_cell.font = BLACK_NORMAL
            else:
                # OK
                if was_red:
                    vie_cell.fill = GREEN_FILL
                else:
                    vie_cell.fill = NO_FILL

                vie_cell.font = BLACK_NORMAL
                c1_cell.fill = NO_FILL
                c1_cell.font = BLACK_NORMAL

        wb.save(out_path)
        wb.close()

        if invalid_count == 0:
            print(f"STATUT: Valide ✅ (Annexe {annexe_num})")
            _open_in_excel_2010(out_path)
            return 0

        print(f"STATUT: Invalide ❌ (Annexe {annexe_num}) | lignes invalides = {invalid_count} | Excel rows: {invalid_rows}")

        proc = _open_in_excel_2010(out_path)
        last_mtime = os.path.getmtime(out_path)
        _wait_for_ctrl_s(out_path, last_mtime)

        _close_excel_process(proc)
        _wait_unlock(out_path, timeout=60)

        print("🔁 Relance validation...")




def _infer_annexe_and_year(input_path: str, default_year: int = 2024):
    """
    Déduit (annexe, année) à partir du nom de fichier.
    Ex:
      12E2024.xlsx -> ("12", 2024)
      ...Annexe_13_2023... -> ("13", 2023)
    """
    import os, re
    name = os.path.basename(input_path).upper()

    m = re.search(r"(20\d{2})", name)
    year = int(m.group(1)) if m else int(default_year)

    if ("ANNEXE_12" in name) or ("ANNEXE 12" in name) or ("12E" in name) or ("12NV" in name):
        ann = "12"
    elif ("ANNEXE_13" in name) or ("ANNEXE 13" in name) or ("13E" in name) or ("13NV" in name):
        ann = "13"
    else:
        ann = "12"

    return ann, year

def _output_nv_path_from_input(in_path: str, ann: str, year: int) -> str:
    """
    Force la sortie NV dans le même dossier que le fichier E (dossier société).
    Ex: ...\COMAR\12E2024.xlsx -> ...\COMAR\12NV2024.xlsx
    """
    folder = os.path.abspath(os.path.dirname(in_path))
    return os.path.join(folder, f"{ann}NV{year}.xlsx")





def _drop_rows_vie_total(df: pd.DataFrame) -> pd.DataFrame:
    """
    Supprime les lignes dont le libellé est exactement 'VIE' ou 'TOTAL'
    (tu ne veux jamais ces 2 lignes en bas).
    """
    if df is None or df.empty:
        return df
    df = df.copy()
    key = df["CATEGORIES"].astype(str).map(lambda x: _normalize_key(x))
    mask = ~key.isin(["VIE", "TOTAL"])
    return df.loc[mask].reset_index(drop=True)


def _write_excel_with_style(df: pd.DataFrame, out_path: str) -> str:
    """
    Ecrit un Excel propre:
      - header bleu
      - bordures
      - montants centrés
      - CATEGORIES à gauche (pas centré)
    """
    out_path = _safe_save_path(out_path)

    # écrire via pandas
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Annexe_12")

    wb = load_workbook(out_path)
    ws = wb.active

    header_fill = PatternFill(start_color=HEADER_COLOR, end_color=HEADER_COLOR, fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    max_row = ws.max_row
    max_col = ws.max_column

    # header style
    for c in range(1, max_col + 1):
        cell = ws.cell(1, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = border

    # cells style
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.border = border
            if c == 1:
                cell.alignment = align_left
            else:
                cell.alignment = align_center

    # auto width
    for c in range(1, max_col + 1):
        max_len = 0
        for r in range(1, max_row + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(c)].width = min(60, max(10, int((max_len + 2) * 1.15)))

    for r in range(1, max_row + 1):
        ws.row_dimensions[r].height = 18

    wb.save(out_path)
    wb.close()
    return out_path


def normalize_excel(input_file: str, out_path: str) -> str:
    """
    Normalise et écrit directement vers out_path (ex: 12NV2024.xlsx dans dossier COMAR).
    """
    df = pd.read_excel(input_file, sheet_name=0, dtype=str)
    if df is None or df.empty:
        raise ValueError("Excel vide / non lisible.")

    cols = list(df.columns)
    if not cols:
        raise ValueError("Aucune colonne détectée.")

    # forcer première colonne
    df = df.rename(columns={cols[0]: "CATEGORIES"})

    # 1) colonnes montants -> VIE/TOTAL
    df = _choose_amount_columns(df)

    # 2) nettoyer montants (enlève symboles/lettres automatiquement)
    df = _clean_numeric_columns(df)

    # 3) normaliser lignes
    df = _normalize_rows_to_canonical(df)

    # 4) supprimer définitivement lignes VIE/TOTAL si présentes
    df = _drop_rows_vie_total(df)

    # écrire avec style (et safe save si ouvert)
    out_path = _write_excel_with_style(df, out_path)
    return out_path




# ----------------------------- VALIDATION (C1 = TOTAL - VIE) -----------------------------
def validate_excel_vie_total_c1(xlsx_path: str) -> tuple[str, str]:
    """
    Ajoute/écrase une colonne C1.
    Pour chaque ligne: C1 = TOTAL - VIE
    Si C1 != 0 -> cellule VIE rouge.
    Retourne: (xlsx_path, "Valide"/"Invalide")
    """
    wb = load_workbook(xlsx_path)
    ws = wb.active

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    headers_norm = [_normalize_key(h) for h in headers]

    def _find_col(name_norm: str):
        for idx, hn in enumerate(headers_norm, start=1):
            if hn == name_norm:
                return idx
        return None

    col_cat = _find_col("CATEGORIES") or 1
    col_vie = _find_col("VIE")
    col_total = _find_col("TOTAL")

    if col_vie is None or col_total is None:
        wb.close()
        raise ValueError("Colonnes VIE/TOTAL introuvables pour validation.")

    # supprimer les lignes VIE/TOTAL si jamais elles existent encore (au niveau Excel)
    rows_to_delete = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, col_cat).value
        if _normalize_key(v) in ("VIE", "TOTAL"):
            rows_to_delete.append(r)
    for r in reversed(rows_to_delete):
        ws.delete_rows(r, 1)

    # refresh headers after deletes
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    headers_norm = [_normalize_key(h) for h in headers]
    col_cat = _find_col("CATEGORIES") or 1
    col_vie = _find_col("VIE")
    col_total = _find_col("TOTAL")

    # trouver / créer col C1
    col_c1 = None
    for c in range(1, ws.max_column + 1):
        if _normalize_key(ws.cell(1, c).value) == "C1":
            col_c1 = c
            break
    if col_c1 is None:
        col_c1 = ws.max_column + 1
        ws.cell(1, col_c1).value = "C1"

    # styles
    header_fill = PatternFill(start_color="0077CC", end_color="0077CC", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    red_fill = PatternFill(start_color="FF4040", end_color="FF4040", fill_type="solid")
    normal_font = Font(color="000000", bold=False)

    # header C1 style
    h = ws.cell(1, col_c1)
    h.fill = header_fill
    h.font = header_font
    h.alignment = align_center

    invalid_count = 0

    for r in range(2, ws.max_row + 1):
        vie = _to_number_or_none(ws.cell(r, col_vie).value) or 0
        total = _to_number_or_none(ws.cell(r, col_total).value) or 0

        c1 = int(round(total - vie))
        ws.cell(r, col_c1).value = c1
        ws.cell(r, col_c1).number_format = "0"

        # align: CATEGORIES left, autres center
        ws.cell(r, col_cat).alignment = align_left
        ws.cell(r, col_vie).alignment = align_center
        ws.cell(r, col_total).alignment = align_center
        ws.cell(r, col_c1).alignment = align_center

        # règle: C1 doit être exactement 0
        if c1 != 0:
            ws.cell(r, col_vie).fill = red_fill
            invalid_count += 1
        else:
            ws.cell(r, col_vie).fill = PatternFill(fill_type=None)
        ws.cell(r, col_vie).font = normal_font

    # ajuster largeur C1
    ws.column_dimensions[get_column_letter(col_c1)].width = 12

    status = "Valide" if invalid_count == 0 else "Invalide"
    print(f"STATUT: {status} (lignes invalides = {invalid_count})")

    # safe save
    safe_path = _safe_save_path(xlsx_path)
    if safe_path != xlsx_path:
        try:
            os.remove(xlsx_path)
        except Exception:
            pass
        xlsx_path = safe_path

    wb.save(xlsx_path)
    wb.close()
    return xlsx_path, status



def _excel_2010_exe_path():
    """
    Chemin Excel 2010 (Office14) le plus probable.
    """
    import os
    candidates = [
        r"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _open_in_excel_2010(path: str):
    """
    Ouvre le fichier avec Excel 2010 explicitement.
    Retourne un objet Popen si succès (pour fermer via PID), sinon None.
    """
    import os
    import subprocess

    excel = _excel_2010_exe_path()
    abspath = os.path.abspath(path)

    if excel:
        try:
            # /e = open in existing instance if possible, else new
            # On garde Popen pour récupérer pid
            return subprocess.Popen([excel, "/e", abspath], shell=False)
        except Exception:
            pass

    # fallback association Windows
    try:
        os.startfile(abspath)
        return None
    except Exception:
        return None


def _close_excel_process(proc):
    """
    Ferme uniquement l'instance Excel lancée via _open_in_excel_2010 (si proc != None).
    """
    import subprocess
    import time

    if proc is None:
        return

    try:
        # ferme uniquement le PID (pas toutes les versions Excel)
        subprocess.call(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.0)
    except Exception:
        pass



# ----------------------------- MAIN (boucle) -----------------------------
def _extract_year_from_name(path: str, default_year: int = 2024) -> int:
    m = re.search(r"(20\d{2})", os.path.basename(path))
    return int(m.group(1)) if m else default_year

def main():
    import os
    import sys

    if len(sys.argv) < 2:
        print("Usage: py C.py <12E2024.xlsx ou 13E2024.xlsx>")
        return 1

    in_path = sys.argv[1].strip('"').strip()
    if not os.path.exists(in_path):
        print(f"❌ Fichier introuvable: {in_path}")
        return 1

    ann, year = _infer_annexe_and_year(in_path, default_year=2024)

    # sortie NV dans le même dossier que le fichier E
    out_path = _output_nv_path_from_input(in_path, ann, year)

    # normalisation -> écrit directement 12NV2024.xlsx / 13NV2024.xlsx
    out_path = normalize_excel(in_path, out_path)
    print(f"✅ Fichier normalisé: {out_path}")

    # validation (boucle Ctrl+S)
    return validate_excel_loop_vie_total(out_path, annexe_num=ann)







if __name__ == "__main__":
    sys.exit(main())
