import re
from config.document_structure import get_subcategory, CP_SUBCATEGORIES, PA_SUBCATEGORIES, PARENT_CODES
from src.utils.helpers import clean_number, extract_trailing_numbers


def _is_purely_numeric_row(row_data):
    """
    Check if ALL non-empty cells in the row are numeric (or a parent code).
    These are parent total rows with just values, no description.
    """
    has_value = False
    for cell in row_data:
        cell_str = str(cell).strip()
        if not cell_str:
            continue
        if cell_str in PARENT_CODES:
            continue
        cleaned = clean_number(cell)
        if isinstance(cleaned, (int, float)):
            has_value = True
        else:
            return False
    return has_value


def _combine_all_text(row_data):
    """
    Combine ALL non-numeric, non-empty cells from a row into a single string.
    Handles text like 'TOTAL DES CAPITAUX PROPRES ET DU PASSIF' split across columns.
    """
    parts = []
    for cell in row_data:
        cell_str = str(cell).strip()
        if not cell_str:
            continue
        cleaned = clean_number(cell)
        if not isinstance(cleaned, (int, float)):
            parts.append(cell_str)
    return ' '.join(parts)


def detect_hierarchy_level_passif(row_data, current_section=None):

    if not row_data or len(row_data) == 0:
        return None

    first_col = str(row_data[0]).strip() if row_data[0] else ""
    second_col = str(row_data[1]).strip() if len(row_data) > 1 and row_data[1] else ""

    # Combine first two columns for code detection
    combined = f"{first_col} {second_col}".strip()
    combined_lower = combined.lower()

    # Combine ALL text columns for keyword matching (totals, sections)
    full_text = _combine_all_text(row_data)
    full_text_lower = full_text.lower()

    # EARLY EXIT: value-only rows (parent totals) - let post-processing handle them
    if _is_purely_numeric_row(row_data):
        return None

    # Level 0: Main title (CAPITAUX PROPRES ET LE PASSIF) but NOT the grand total
    if ("capitaux propres et" in full_text_lower
            and "passif" in full_text_lower
            and "total" not in full_text_lower):
        return (0, "", full_text, False, "TITRE", "", [])

    # Level 1: Main sections (CAPITAUX PROPRES:, PASSIF:)
    if re.match(r'^(CAPITAUX PROPRES|PASSIF):?$', combined, re.IGNORECASE):
        section = "CAPITAUX PROPRES" if "capitaux propres" in combined_lower else "PASSIF"
        return (1, "", combined, False, "SECTION", section, [])

    # Main totals - use full_text for matching to catch multi-column text
    # Check TOTAL GENERAL first since it also contains other total keywords
    if "total" in full_text_lower:
        clean_desc, extra_vals = extract_trailing_numbers(full_text)

        if "total des capitaux propres et du passif" in full_text_lower:
            return (1, "", clean_desc, True, "TOTAL G\u00c9N\u00c9RAL", "", extra_vals)
        elif ("total capitaux propres avant r\u00e9sultat" in full_text_lower
                or "total capitaux propres avant resultat" in full_text_lower):
            return (2, "", clean_desc, True, "TOTAL",
                    "Capitaux Propres - Avant R\u00e9sultat", extra_vals)
        elif "total capitaux propres avant affectation" in full_text_lower:
            return (2, "", clean_desc, True, "TOTAL",
                    "Capitaux Propres - Avant Affectation", extra_vals)
        elif "total du passif" in full_text_lower:
            return (2, "", clean_desc, True, "TOTAL", "Total Passif", extra_vals)
        else:
            category = current_section if current_section else "TOTAL"
            return (3, "", clean_desc, True, "TOTAL", category, extra_vals)

    # Level 2: PA codes (PASSIF subsections)
    if re.match(r'^(PA\d+|PA\d+[A-Z]?\d*)\s+', combined):
        code_match = re.match(r'^(PA\d+[A-Z]?\d*)\s+(.+)', combined)
        if code_match:
            code = code_match.group(1)
            desc_raw = code_match.group(2)
            desc, extra_vals = extract_trailing_numbers(desc_raw)
            subcategory = get_subcategory(code)

            if code in PARENT_CODES:
                return (2, code, desc, False, "PASSIF", subcategory, extra_vals)
            else:
                return (3, code, desc, False, "PASSIF", subcategory, extra_vals)

    # Level 2: CP codes (Capital)
    if re.match(r'^(CP\d+)\s+', combined):
        code_match = re.match(r'^(CP\d+)\s+(.+)', combined)
        if code_match:
            code = code_match.group(1)
            desc_raw = code_match.group(2)
            desc, extra_vals = extract_trailing_numbers(desc_raw)
            subcategory = get_subcategory(code)
            return (2, code, desc, False, "CAPITAUX PROPRES", subcategory, extra_vals)

    # Code alone (CP or PA)
    if re.match(r'^(CP\d+|PA\d+[A-Z]?\d*)$', first_col):
        if first_col.startswith('CP'):
            return (2, first_col, second_col, False, "CAPITAUX PROPRES", "", [])
        else:
            return (2, first_col, second_col, False, "PASSIF", "", [])

    # Description line without code (level 2 by default)
    if first_col and not re.match(r'^(CP|PA)', first_col):
        desc, extra_vals = extract_trailing_numbers(combined)
        category = current_section if current_section else "AUTRE"
        return (2, "", desc, False, category, "", extra_vals)

    return None


def _extract_numeric_values_from_row(row):
    """Extract all numeric values from a row, regardless of position."""
    values = []
    for cell in row:
        cleaned = clean_number(cell)
        if isinstance(cleaned, (int, float)):
            values.append(cleaned)
    return values


def _find_parent_code_in_row(row):
    """Check if any cell contains exactly a parent code."""
    for cell in row:
        cell_str = str(cell).strip()
        if cell_str in PARENT_CODES:
            return cell_str
    return None


def _find_parent_for_code(code):
    """
    Given a code like PA710, find its parent in PARENT_CODES.
    Tries progressively shorter prefixes: PA71 -> PA7.
    """
    if not code:
        return None
    for length in range(len(code) - 1, 1, -1):
        prefix = code[:length]
        if prefix in PARENT_CODES:
            return prefix
    return None


def structure_hierarchical_data_passif(raw_data):
    """
    Structure raw table data into hierarchical format.

    Two-pass approach:
      Pass 1: Parse all rows normally. Value-only rows are collected separately.
              Track which parents actually have children in the data.
      Pass 2: Attribute each value-only total row to the correct parent.
              Key rule: if a parent has inline values but NO children,
              the value-only row belongs to a HIGHER parent (climb up).
              e.g. PA72 has values but no children -> climb to PA7.
    """
    # -- Pass 1: Parse all rows --
    hierarchical_rows = []
    unmatched_rows = []
    current_section = None
    last_code_seen = None
    # Track which parent codes have children appearing in the data
    parents_with_children = set()

    for row in raw_data:
        if not any(str(cell).strip() for cell in row):
            continue

        hierarchy_info = detect_hierarchy_level_passif(row, current_section)

        if hierarchy_info:
            level, code, description, is_total, category, subcategory, extra_values = hierarchy_info

            if category == "SECTION":
                current_section = subcategory

            # Track parent-child relationships:
            # If this code is a child of any parent, mark that parent
            if code and not is_total:
                for parent_code in PARENT_CODES:
                    if code.startswith(parent_code) and code != parent_code:
                        parents_with_children.add(parent_code)

            if code:
                last_code_seen = code

            values = []
            if extra_values:
                values.extend(extra_values)

            for cell in row[2:]:
                cleaned = clean_number(cell)
                if cleaned != '':
                    values.append(cleaned)

            if not values and len(row) >= 2:
                for cell in row:
                    cleaned = clean_number(cell)
                    if isinstance(cleaned, (int, float)):
                        values.append(cleaned)

            hierarchical_rows.append({
                'level': level,
                'code': code,
                'description': description,
                'is_total': is_total,
                'category': category,
                'subcategory': subcategory,
                'values': values
            })
        else:
            # Value-only row. Store with context for Pass 2.
            values = _extract_numeric_values_from_row(row)
            parent_code_in_row = _find_parent_code_in_row(row)

            if values:
                unmatched_rows.append({
                    'values': values,
                    'explicit_parent': parent_code_in_row,
                    'last_code_before': last_code_seen,
                })

    # -- Pass 2: Attribute value-only rows to their parent headers --

    # Build index of parent headers
    parent_header_indices = {}
    for i, row in enumerate(hierarchical_rows):
        code = row.get('code', '')
        if code in PARENT_CODES and not row['is_total']:
            parent_header_indices[code] = i

    # Track which parents get assigned during this pass
    parents_assigned = set()

    for unmatched in unmatched_rows:
        target_parent = None

        if unmatched['explicit_parent']:
            # Explicit parent code in the row (e.g. "PA3" in Note column)
            target_parent = unmatched['explicit_parent']
        else:
            last_code = unmatched['last_code_before']
            if last_code:
                # Find the closest parent for this code
                if last_code in PARENT_CODES:
                    candidate = last_code
                else:
                    candidate = _find_parent_for_code(last_code)

                # Climb up if:
                #   - candidate already assigned a total, OR
                #   - candidate has inline values but NO children
                #     (e.g. PA72 has "3,515,116" but no PA72x children,
                #      so the total row belongs to PA7, not PA72)
                while candidate:
                    if candidate in parents_assigned:
                        candidate = _find_parent_for_code(candidate)
                    elif candidate in parent_header_indices:
                        header_idx = parent_header_indices[candidate]
                        has_inline_values = bool(
                            hierarchical_rows[header_idx]['values'])
                        has_children = candidate in parents_with_children

                        if has_inline_values and not has_children:
                            # Leaf parent with values -> total belongs higher
                            candidate = _find_parent_for_code(candidate)
                        else:
                            break
                    else:
                        break

                target_parent = candidate

        if target_parent and target_parent in parent_header_indices:
            header_idx = parent_header_indices[target_parent]
            # Always overwrite: the total row is authoritative
            hierarchical_rows[header_idx]['values'] = unmatched['values']
            parents_assigned.add(target_parent)

    return hierarchical_rows
