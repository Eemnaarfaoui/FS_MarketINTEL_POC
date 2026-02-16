from typing import Dict, Any, Optional


class ValidatorPassifs:
    """
    Validates PASSIF table by checking that every parent = sum(children).

    Rules checked:
    ─── CAPITAUX PROPRES ───
      Total CP Av Résultat    = CP1 + CP2 + CP3 + CP4 + CP5
      Total CP Av Affectation = CP1 + CP2 + CP3 + CP4 + CP5 + CP6

    ─── PASSIF ───
      PA1  = PA13 + PA14
      PA2  = PA23  (+ any other PA2x children present)
      PA3  = PA310 + PA320 + PA330 + PA331 + PA340 + PA341 + PA350 + PA360 + PA361
      PA5  = standalone (no children to check)
      PA6  = PA61 + PA62 + PA63
      PA62 = PA622
      PA63 = PA631 + PA632 + PA633 + PA634
      PA7  = PA71 + PA73  (+ PA72 if present)
      PA71 = PA710 + PA711 + PA712

    ─── TOTALS ───
      Total Passif         = PA1 + PA2 + PA3 + PA5 + PA6 + PA7
      Total CP et Passif   = (CP1+...+CP6) + (PA1+PA2+PA3+PA5+PA6+PA7)
    """

    # Parent → list of children codes
    PARENT_CHILDREN = {
        'PA1':  ['PA13', 'PA14'],
        'PA2':  ['PA23'],
        'PA3':  ['PA310', 'PA320', 'PA330', 'PA331', 'PA340', 'PA341', 'PA350', 'PA360', 'PA361'],
        # PA5 has no children
        'PA6':  ['PA61', 'PA62', 'PA63'],
        'PA62': ['PA622'],
        'PA63': ['PA631', 'PA632', 'PA633', 'PA634'],
        'PA7':  ['PA71', 'PA72', 'PA73'],
        'PA71': ['PA710', 'PA711', 'PA712'],
    }

    CP_BEFORE_RESULT = ['CP1', 'CP2', 'CP3', 'CP4', 'CP5']
    CP_BEFORE_AFFECTATION = ['CP1', 'CP2', 'CP3', 'CP4', 'CP5', 'CP6']
    TOP_LEVEL_PA = ['PA1', 'PA2', 'PA3', 'PA5', 'PA6', 'PA7']

    def __init__(self, extracted_data_context: Optional[Dict] = None, error_margin: float = 1.0):
        """
        Args:
            extracted_data_context: dict of {code: {'value': float}} from the Excel
            error_margin: allowed rounding difference (default 1.0 for integer rounding)
        """
        self.ctx = extracted_data_context or {}
        self.error_margin = error_margin

    def _get_value(self, code: str) -> Optional[float]:
        """Get value for a code from context. Returns None if not found."""
        entry = self.ctx.get(code)
        if entry is None:
            return None
        val = entry.get('value')
        if val is None:
            return None
        try:
            return float(val)
        except (TypeError, ValueError):
            return None

    def _sum_children(self, children: list) -> Optional[float]:
        """
        Sum values of children codes. 
        Returns None if no children found at all.
        Missing children treated as 0 (they might not exist in this company).
        """
        total = 0.0
        found_any = False
        for code in children:
            val = self._get_value(code)
            if val is not None:
                total += val
                found_any = True
        return total if found_any else None

    def _get_first_numeric_value(self, values) -> Optional[float]:
        """Get first numeric value from a list."""
        if not values:
            return None
        for v in values:
            try:
                f = float(v)
                return f
            except (TypeError, ValueError):
                continue
        return None

    def _close_enough(self, a: float, b: float) -> bool:
        """Check if two values are within error margin."""
        return abs(a - b) <= self.error_margin

    # ══════════════════════════════════════════════════════════════
    # PARENT-CHILD VALIDATION (PA codes)
    # ══════════════════════════════════════════════════════════════

    def _validate_parent_children(self, code: str, parent_value: float) -> bool:
        """
        Check: parent_value == sum(children values from context)
        Returns True if:
          - Code is not a parent (no rule to check)
          - No children found in context (can't validate)
          - Sum matches within margin
        Returns False if sum doesn't match.
        """
        children = self.PARENT_CHILDREN.get(code)
        if not children:
            return True  # Not a parent, nothing to check

        children_sum = self._sum_children(children)
        if children_sum is None:
            return True  # No children data, can't validate

        return self._close_enough(parent_value, children_sum)

    # ══════════════════════════════════════════════════════════════
    # CP TOTAL VALIDATIONS
    # ══════════════════════════════════════════════════════════════

    def _validate_cp_avant_resultat(self, description: str, row_value: float) -> bool:
        """Total CP Av Résultat = CP1 + CP2 + CP3 + CP4 + CP5"""
        desc_lower = description.lower()
        if 'avant r' not in desc_lower and 'av r' not in desc_lower:
            return True  # Not this total row

        expected = self._sum_children(self.CP_BEFORE_RESULT)
        if expected is None:
            return True

        return self._close_enough(row_value, expected)

    def _validate_cp_avant_affectation(self, description: str, row_value: float) -> bool:
        """Total CP Av Affectation = CP1 + CP2 + CP3 + CP4 + CP5 + CP6"""
        desc_lower = description.lower()
        if 'avant affectation' not in desc_lower and 'av affectation' not in desc_lower:
            return True

        expected = self._sum_children(self.CP_BEFORE_AFFECTATION)
        if expected is None:
            return True

        return self._close_enough(row_value, expected)

    def _validate_cp_consolide(self, description: str, row_value: float) -> bool:
        """Total capitaux propres consolidés = CP1 + CP2 + CP3 + CP4 + CP5 + CP6"""
        desc_lower = description.lower()
        if 'consolid' not in desc_lower:
            return True

        expected = self._sum_children(self.CP_BEFORE_AFFECTATION)
        if expected is None:
            return True

        return self._close_enough(row_value, expected)

    # ══════════════════════════════════════════════════════════════
    # PASSIF TOTAL VALIDATIONS
    # ══════════════════════════════════════════════════════════════

    def _validate_total_passif(self, description: str, row_value: float) -> bool:
        """Total du Passif = PA1 + PA2 + PA3 + PA5 + PA6 + PA7"""
        desc_lower = description.lower()
        if 'total du passif' not in desc_lower:
            return True

        expected = self._sum_children(self.TOP_LEVEL_PA)
        if expected is None:
            return True

        return self._close_enough(row_value, expected)

    def _validate_total_general(self, description: str, row_value: float) -> bool:
        """
        Total des capitaux propres et du passif (or standalone "Total")
        = CP total + PA total
        = (CP1+CP2+CP3+CP4+CP5+CP6) + (PA1+PA2+PA3+PA5+PA6+PA7)
        """
        desc_lower = description.lower()
        is_grand_total = (
            ('capitaux propres' in desc_lower and 'passif' in desc_lower and 'total' in desc_lower)
            or desc_lower.strip() == 'total'
        )
        if not is_grand_total:
            return True

        cp_sum = self._sum_children(self.CP_BEFORE_AFFECTATION)
        pa_sum = self._sum_children(self.TOP_LEVEL_PA)

        if cp_sum is None or pa_sum is None:
            return True

        expected = cp_sum + pa_sum
        return self._close_enough(row_value, expected)

    # ══════════════════════════════════════════════════════════════
    # MAIN VALIDATE METHOD
    # ══════════════════════════════════════════════════════════════

    def validate(self, row: Dict[str, Any]) -> bool:
        """
        Validate a single row. The row dict should contain:
          'code': str (e.g. 'PA3', 'CP1')
          'description': str (e.g. 'Provisions techniques brutes')
          'values': list of numeric values [val_2024, val_2023, ...]

        Returns True if all applicable rules pass, False if any fail.
        """
        code = str(row.get('code', '')).strip()
        description = str(row.get('description', ''))
        values = row.get('values', [])
        row_value = self._get_first_numeric_value(values)

        # If no numeric value, can't validate → pass
        if row_value is None:
            return True

        results = []

        # ── Rule 1: Parent-child sum check (PA codes) ──
        # If this row's code is a parent, check value == sum(children)
        if code:
            results.append(self._validate_parent_children(code, row_value))

        # ── Rule 2: CP avant résultat ──
        results.append(self._validate_cp_avant_resultat(description, row_value))

        # ── Rule 3: CP avant affectation ──
        results.append(self._validate_cp_avant_affectation(description, row_value))

        # ── Rule 4: CP consolidés ──
        results.append(self._validate_cp_consolide(description, row_value))

        # ── Rule 5: Total du passif ──
        results.append(self._validate_total_passif(description, row_value))

        # ── Rule 6: Grand total ──
        results.append(self._validate_total_general(description, row_value))

        return all(results)