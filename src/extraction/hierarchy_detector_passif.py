import re
from config.document_structure import get_subcategory, CP_SUBCATEGORIES, PA_SUBCATEGORIES, PARENT_CODES
from src.utils.helpers import clean_number, extract_trailing_numbers

_PARENT_CODES_SET = set(PARENT_CODES)


def _is_numeric(value):
    """Check if a string represents a number, including space-separated like '45 000 000'."""
    s = str(value).strip()
    if not s or s == '-':
        return False
    s = s.replace('\u00a0', '').replace(' ', '').replace(',', '.').replace('\xa0', '')
    try:
        float(s)
        return True
    except ValueError:
        return False


# ══════════════════════════════════════════════════════════════════
# RAW DATA PRE-PROCESSING
# ══════════════════════════════════════════════════════════════════

def _preprocess_raw_data(raw_data):
    """
    Clean up raw Camelot output. Handles:
      A) Cross-row code splits: ['P', '482797'...] + ['A2 Provisions...'] → ['PA2 Provisions...', '482797'...]
      B) Backward merge: value-only row + text-only row → merged (PA14 gets values from preceding row)
      C) Forward merge: text-only row + value-only row → merged
      D) Within-row code splits: ['P', 'A2 Prov...'] → ['PA2 Prov...', '']
      E) Consecutive duplicate removal (e.g. CP5 appearing twice)
    """
    if not raw_data:
        return raw_data

    rows = [[str(cell) if cell is not None else "" for cell in row] for row in raw_data]

    merged = []
    i = 0
    while i < len(rows):
        row = rows[i]

        if not any(cell.strip() for cell in row):
            i += 1
            continue

        # ── Fix A: Cross-row code split ──
        # ROW N:   ['P', '482 797', '518 176', '', '']
        # ROW N+1: ['A2 Provisions pour autres risques et Charges', '', '', '', '']
        # → ['PA2 Provisions pour autres risques et Charges', '482 797', '518 176', '', '']
        if i + 1 < len(rows):
            col0 = row[0].strip()
            next_row = rows[i + 1]
            next_col0 = next_row[0].strip()

            if (len(col0) <= 2
                    and col0 in ('P', 'C', 'PA', 'CP', 'IM')
                    and next_col0
                    and re.match(r'^(PA\d+|CP\d+|IMCP\d*|A\d+|P\d+)',
                                 col0 + next_col0.split()[0])):
                new_row = list(row)
                new_row[0] = col0 + next_col0
                for j in range(1, len(next_row)):
                    ns = next_row[j].strip()
                    if ns and j < len(new_row) and not new_row[j].strip():
                        new_row[j] = next_row[j]
                row = new_row
                i += 1

        # ── Fix D: Within-row code split ──
        row = _rejoin_split_codes_in_row(row)

        # ── Fix C: Forward merge (text-only + value-only) ──
        if i + 1 < len(rows):
            next_row = rows[i + 1]
            fwd = _try_merge_forward(row, next_row)
            if fwd:
                row = fwd
                i += 1

        # ── Fix B: Backward merge (value-only + text-only with code) ──
        if i + 1 < len(rows):
            next_row = rows[i + 1]
            bwd = _try_merge_backward(row, next_row)
            if bwd:
                rows[i + 1] = bwd
                i += 1
                continue  # Skip current value-only row

        merged.append(row)
        i += 1

    # ── Fix E: Remove consecutive duplicates ──
    deduped = []
    for row in merged:
        if not any(cell.strip() for cell in row):
            continue
        if deduped and _rows_are_duplicate(deduped[-1], row):
            continue
        deduped.append(row)

    return deduped


def _rejoin_split_codes_in_row(row):
    """Fix codes split across adjacent cells within a row."""
    if len(row) < 2:
        return row
    row = list(row)
    for j in range(len(row) - 1):
        cell = row[j].strip()
        next_cell = row[j + 1].strip()
        if not cell or not next_cell:
            continue
        if _is_numeric(next_cell):
            continue
        candidate = cell + next_cell.split()[0]
        if re.match(r'^(PA\d+[A-Z]?\d*|CP\d+|IMCP\d*)', candidate):
            row[j] = cell + next_cell
            row[j + 1] = ""
            break
    return row


def _try_merge_forward(current_row, next_row):
    """Merge if current=text-only, next=value-only."""
    curr_text, curr_val = False, False
    for cell in current_row:
        s = cell.strip()
        if not s: continue
        if _is_numeric(s): curr_val = True
        else: curr_text = True

    next_text, next_val = False, False
    for cell in next_row:
        s = cell.strip()
        if not s: continue
        if _is_numeric(s): next_val = True
        elif s not in _PARENT_CODES_SET: next_text = True

    if curr_text and not curr_val and next_val and not next_text:
        merged = list(current_row)
        for j, cell in enumerate(next_row):
            s = cell.strip()
            if s and j < len(merged):
                if not merged[j].strip():
                    merged[j] = cell
                else:
                    merged[j] = merged[j] + " " + cell
            elif s:
                merged.append(cell)
        return merged
    return None


def _try_merge_backward(current_row, next_row):
    """
    Merge if current=value-only, next=text-only with a code.
    ROW N:   ['', '28 000 000', '15 000 000', '', '']
    ROW N+1: ['PA14 Dettes Etab...', '', '', '', '']
    → ['PA14 Dettes Etab...', '28 000 000', '15 000 000', '', '']
    """
    curr_has_text, curr_has_value = False, False
    for cell in current_row:
        s = cell.strip()
        if not s: continue
        if _is_numeric(s): curr_has_value = True
        elif s not in _PARENT_CODES_SET: curr_has_text = True

    next_has_text, next_has_value = False, False
    next_col0 = next_row[0].strip() if next_row else ""
    for cell in next_row:
        s = cell.strip()
        if not s: continue
        if _is_numeric(s): next_has_value = True
        else: next_has_text = True

    has_code = bool(re.match(r'^(PA|CP|IMCP)', next_col0))
    if curr_has_value and not curr_has_text and next_has_text and not next_has_value and has_code:
        merged = list(next_row)
        for j, cell in enumerate(current_row):
            s = cell.strip()
            if s and _is_numeric(s) and j < len(merged) and not merged[j].strip():
                merged[j] = cell
        return merged
    return None


def _rows_are_duplicate(row1, row2):
    cells1 = [str(c).strip() for c in row1 if str(c).strip()]
    cells2 = [str(c).strip() for c in row2 if str(c).strip()]
    return cells1 == cells2 and len(cells1) > 0


# ══════════════════════════════════════════════════════════════════
# TEXT / VALUE SEPARATION
# ══════════════════════════════════════════════════════════════════

def _get_text_and_values(row_data):
    """
    Separate text columns from value columns.
    Stops at the first numeric cell.

    ['CP1 Capital social', '45 000 000', '45 000 000', '', '']
     ^ TEXT               ^ NUMERIC → STOP

    Returns ("CP1 Capital social", 1)
    """
    text_parts = []
    first_value_idx = len(row_data)
    for i, cell in enumerate(row_data):
        cell_str = str(cell).strip()
        if not cell_str:
            continue
        if _is_numeric(cell_str):
            first_value_idx = i
            break
        if cell_str in _PARENT_CODES_SET and i > 0:
            continue
        text_parts.append(cell_str)
    return " ".join(text_parts), first_value_idx


def _is_purely_numeric_row(row_data):
    has_value = False
    for cell in row_data:
        cell_str = str(cell).strip()
        if not cell_str: continue
        if cell_str in _PARENT_CODES_SET: continue
        if _is_numeric(cell_str): has_value = True
        else: return False
    return has_value


def _combine_all_text(row_data):
    parts = []
    for cell in row_data:
        cell_str = str(cell).strip()
        if not cell_str: continue
        if not _is_numeric(cell_str):
            parts.append(cell_str)
    return ' '.join(parts)


# ══════════════════════════════════════════════════════════════════
# TOTAL CLASSIFICATION
# ══════════════════════════════════════════════════════════════════

def _classify_total(text_lower):
    if ('total des capitaux propres et du passif' in text_lower
            or ('total des capitaux' in text_lower and 'passif' in text_lower)):
        return 'TOTAL_GENERAL'
    stripped = re.sub(r'[^a-zéèêà\s]', '', text_lower).strip()
    if stripped == 'total':
        return 'TOTAL_GENERAL'
    if ('total capitaux propres avant affectation' in text_lower
            or 'total cp av affectation' in text_lower):
        return 'CP_AFFECTATION'
    if ('total capitaux propres avant r' in text_lower
            or 'total cp av r' in text_lower):
        return 'CP_RESULTAT'
    if 'total capitaux propres consolid' in text_lower:
        return 'CP_CONSOLIDE'
    if 'total du passif' in text_lower:
        return 'TOTAL_PASSIF'
    return 'OTHER'


# ══════════════════════════════════════════════════════════════════
# HIERARCHY DETECTION
# ══════════════════════════════════════════════════════════════════

def detect_hierarchy_level_passif(row_data, current_section=None):
    if not row_data or len(row_data) == 0:
        return None

    combined, first_value_idx = _get_text_and_values(row_data)
    combined_lower = combined.lower()
    full_text = _combine_all_text(row_data)
    full_text_lower = full_text.lower()

    if _is_purely_numeric_row(row_data):
        return None

    # Title
    if ("capitaux propres et" in full_text_lower
            and "passif" in full_text_lower
            and "total" not in full_text_lower):
        return (0, "", full_text, False, "TITRE", "", [])

    # Section headers
    if re.match(r'^(CAPITAUX PROPRES|PASSIF):?$', combined, re.IGNORECASE):
        section = "CAPITAUX PROPRES" if "capitaux propres" in combined_lower else "PASSIF"
        return (1, "", combined, False, "SECTION", section, [])

    if re.match(r'^CP\s+Capitaux\s+Propres$', combined, re.IGNORECASE):
        return (1, "", combined, False, "SECTION", "CAPITAUX PROPRES", [])

    # Totals
    if "total" in full_text_lower:
        clean_desc, extra_vals = extract_trailing_numbers(combined)
        total_type = _classify_total(full_text_lower)
        if total_type == 'TOTAL_GENERAL':
            return (1, "", clean_desc, True, "TOTAL GÉNÉRAL", "", extra_vals)
        elif total_type == 'CP_RESULTAT':
            return (2, "", clean_desc, True, "TOTAL",
                    "Capitaux Propres - Avant Résultat", extra_vals)
        elif total_type == 'CP_CONSOLIDE':
            return (2, "", clean_desc, True, "TOTAL",
                    "Capitaux Propres - Consolidés", extra_vals)
        elif total_type == 'CP_AFFECTATION':
            return (2, "", clean_desc, True, "TOTAL",
                    "Capitaux Propres - Avant Affectation", extra_vals)
        elif total_type == 'TOTAL_PASSIF':
            return (2, "", clean_desc, True, "TOTAL", "Total Passif", extra_vals)
        else:
            category = current_section if current_section else "TOTAL"
            return (3, "", clean_desc, True, "TOTAL", category, extra_vals)

    # PA codes
    if re.match(r'^(PA\d+[A-Z]?\d*)\s+', combined):
        code_match = re.match(r'^(PA\d+[A-Z]?\d*)\s+(.+)', combined)
        if code_match:
            code = code_match.group(1)
            desc_raw = code_match.group(2)
            desc, extra_vals = extract_trailing_numbers(desc_raw)
            subcategory = get_subcategory(code)
            if code in _PARENT_CODES_SET:
                return (2, code, desc, False, "PASSIF", subcategory, extra_vals)
            else:
                return (3, code, desc, False, "PASSIF", subcategory, extra_vals)

    # CP codes (also handles CP2'-, IMCP-, CP6'-)
    cp_match = re.match(r'^((?:CP|IMCP)\d*[\'\'\-]*)\s+(.+)', combined)
    if cp_match:
        code = cp_match.group(1)
        desc_raw = cp_match.group(2)
        desc, extra_vals = extract_trailing_numbers(desc_raw)
        clean_code = code.rstrip("'-")
        subcategory = get_subcategory(clean_code)
        return (2, clean_code, desc, False, "CAPITAUX PROPRES", subcategory, extra_vals)

    # Code alone in first column
    first_col = str(row_data[0]).strip() if row_data[0] else ""
    if re.match(r'^(CP\d+|PA\d+[A-Z]?\d*)$', first_col):
        desc = ""
        for cell in row_data[1:]:
            cell_str = str(cell).strip()
            if cell_str and not _is_numeric(cell_str) and cell_str not in _PARENT_CODES_SET:
                desc = cell_str
                break
        if first_col.startswith('CP'):
            return (2, first_col, desc, False, "CAPITAUX PROPRES", "", [])
        else:
            return (2, first_col, desc, False, "PASSIF", "", [])

    # Description line without code
    if combined and not re.match(r'^(CP|PA)', combined):
        desc, extra_vals = extract_trailing_numbers(combined)
        category = current_section if current_section else "AUTRE"
        return (2, "", desc, False, category, "", extra_vals)

    return None


# ══════════════════════════════════════════════════════════════════
# POST-PROCESSING HELPERS
# ══════════════════════════════════════════════════════════════════

def _extract_numeric_values_from_row(row):
    values = []
    for cell in row:
        cleaned = clean_number(cell)
        if isinstance(cleaned, (int, float)):
            values.append(cleaned)
    return values


def _find_parent_code_in_row(row):
    for cell in row:
        cell_str = str(cell).strip()
        if cell_str in _PARENT_CODES_SET:
            return cell_str
    return None


def _find_parent_for_code(code):
    if not code:
        return None
    for length in range(len(code) - 1, 1, -1):
        prefix = code[:length]
        if prefix in _PARENT_CODES_SET:
            return prefix
    return None


# ══════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ══════════════════════════════════════════════════════════════════

def structure_hierarchical_data_passif(raw_data):
    """
    Structure raw table data into hierarchical format.

    Pipeline:
      0. Pre-process: fix split codes, merge incomplete rows, deduplicate
      1. Parse all rows. Value-only rows collected separately.
      2. Attribute each value-only total row to the correct parent.
    """
    raw_data = _preprocess_raw_data(raw_data)

    hierarchical_rows = []
    unmatched_rows = []
    current_section = None
    last_code_seen = None
    parents_with_children = set()

    for row in raw_data:
        if not any(str(cell).strip() for cell in row):
            continue

        hierarchy_info = detect_hierarchy_level_passif(row, current_section)

        if hierarchy_info:
            level, code, description, is_total, category, subcategory, extra_values = hierarchy_info

            if category == "SECTION":
                current_section = subcategory

            if code and not is_total:
                for parent_code in PARENT_CODES:
                    if code.startswith(parent_code) and code != parent_code:
                        parents_with_children.add(parent_code)

            if code:
                last_code_seen = code

            values = []
            if extra_values:
                values.extend(extra_values)

            _, first_val_idx = _get_text_and_values(row)
            for cell in row[first_val_idx:]:
                cleaned = clean_number(cell)
                if isinstance(cleaned, (int, float)):
                    values.append(cleaned)

            if not values:
                for cell in row:
                    cleaned = clean_number(cell)
                    if isinstance(cleaned, (int, float)):
                        values.append(cleaned)

            hierarchical_rows.append({
                'level': level, 'code': code, 'description': description,
                'is_total': is_total, 'category': category,
                'subcategory': subcategory, 'values': values
            })
        else:
            values = _extract_numeric_values_from_row(row)
            parent_code_in_row = _find_parent_code_in_row(row)
            if values:
                unmatched_rows.append({
                    'values': values,
                    'explicit_parent': parent_code_in_row,
                    'last_code_before': last_code_seen,
                })

    # ── Attribute value-only rows to parent headers ──
    parent_header_indices = {}
    for i, row in enumerate(hierarchical_rows):
        code = row.get('code', '')
        if code in _PARENT_CODES_SET and not row['is_total']:
            parent_header_indices[code] = i

    parents_assigned = set()

    for unmatched in unmatched_rows:
        target_parent = None

        if unmatched['explicit_parent']:
            target_parent = unmatched['explicit_parent']
        else:
            last_code = unmatched['last_code_before']
            if last_code:
                candidate = last_code if last_code in _PARENT_CODES_SET else _find_parent_for_code(last_code)
                while candidate:
                    if candidate in parents_assigned:
                        candidate = _find_parent_for_code(candidate)
                    elif candidate in parent_header_indices:
                        header_idx = parent_header_indices[candidate]
                        has_inline = bool(hierarchical_rows[header_idx]['values'])
                        has_children = candidate in parents_with_children
                        if has_inline and not has_children:
                            candidate = _find_parent_for_code(candidate)
                        else:
                            break
                    else:
                        break
                target_parent = candidate

        if target_parent and target_parent in parent_header_indices:
            header_idx = parent_header_indices[target_parent]
            hierarchical_rows[header_idx]['values'] = unmatched['values']
            parents_assigned.add(target_parent)

    return hierarchical_rows