from typing import Dict, Any, Optional, List, Tuple


class ValidatorPassifs:
    """
    Validates PASSIF table by checking that every parent = sum(children).
    Returns detailed breakdown on failure showing each child's contribution.
    """

    PARENT_CHILDREN = {
        'PA1':  ['PA13', 'PA14'],
        'PA2':  ['PA23'],
        'PA3':  ['PA310', 'PA320', 'PA330', 'PA331', 'PA340', 'PA341', 'PA350', 'PA360', 'PA361'],
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
        self.ctx = extracted_data_context or {}
        self.error_margin = error_margin

    def _get_value(self, code: str) -> Optional[float]:
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

    def _get_first_numeric_value(self, values) -> Optional[float]:
        if not values:
            return None
        for v in values:
            try:
                return float(v)
            except (TypeError, ValueError):
                continue
        return None

    def _close_enough(self, a: float, b: float) -> bool:
        return abs(a - b) <= self.error_margin

    def _build_breakdown(self, children: list) -> Tuple[Optional[float], str]:
        """
        Sum children and build a readable breakdown string.
        Returns (total, breakdown_str)
        Example: (223660142, "PA310=61,840,376 + PA330=677,867 + PA331=151,500,161 + ...")
        """
        total = 0.0
        found_any = False
        parts = []

        for code in children:
            val = self._get_value(code)
            if val is not None:
                total += val
                found_any = True
                parts.append(f"{code}={val:,.0f}")
            # Skip codes with None (not present in this company)

        breakdown = " + ".join(parts) if parts else "no children found"
        return (total if found_any else None), breakdown

    # ══════════════════════════════════════════════════════════════
    # VALIDATION RULES - each returns (passed: bool, detail: str)
    # ══════════════════════════════════════════════════════════════

    def _check_parent_children(self, code: str, parent_value: float) -> Tuple[bool, str]:
        """Check: parent_value == sum(children)"""
        children = self.PARENT_CHILDREN.get(code)
        if not children:
            return True, ""

        children_sum, breakdown = self._build_breakdown(children)
        if children_sum is None:
            return True, ""

        if self._close_enough(parent_value, children_sum):
            return True, ""

        diff = parent_value - children_sum
        return False, (
            f"{code}={parent_value:,.0f} ≠ children sum={children_sum:,.0f} "
            f"(diff={diff:+,.0f}). Breakdown: {breakdown}"
        )

    def _check_cp_avant_resultat(self, description: str, row_value: float) -> Tuple[bool, str]:
        desc_lower = description.lower()
        if 'avant r' not in desc_lower and 'av r' not in desc_lower:
            return True, ""

        expected, breakdown = self._build_breakdown(self.CP_BEFORE_RESULT)
        if expected is None:
            return True, ""

        if self._close_enough(row_value, expected):
            return True, ""

        diff = row_value - expected
        return False, (
            f"Total CP Av Résultat={row_value:,.0f} ≠ expected={expected:,.0f} "
            f"(diff={diff:+,.0f}). Breakdown: {breakdown}"
        )

    def _check_cp_avant_affectation(self, description: str, row_value: float) -> Tuple[bool, str]:
        desc_lower = description.lower()
        if 'avant affectation' not in desc_lower and 'av affectation' not in desc_lower:
            return True, ""

        expected, breakdown = self._build_breakdown(self.CP_BEFORE_AFFECTATION)
        if expected is None:
            return True, ""

        if self._close_enough(row_value, expected):
            return True, ""

        diff = row_value - expected
        return False, (
            f"Total CP Av Affectation={row_value:,.0f} ≠ expected={expected:,.0f} "
            f"(diff={diff:+,.0f}). Breakdown: {breakdown}"
        )

    def _check_cp_consolide(self, description: str, row_value: float) -> Tuple[bool, str]:
        desc_lower = description.lower()
        if 'consolid' not in desc_lower:
            return True, ""

        expected, breakdown = self._build_breakdown(self.CP_BEFORE_AFFECTATION)
        if expected is None:
            return True, ""

        if self._close_enough(row_value, expected):
            return True, ""

        diff = row_value - expected
        return False, (
            f"Total CP Consolidés={row_value:,.0f} ≠ expected={expected:,.0f} "
            f"(diff={diff:+,.0f}). Breakdown: {breakdown}"
        )

    def _check_total_passif(self, description: str, row_value: float) -> Tuple[bool, str]:
        desc_lower = description.lower()
        if 'total du passif' not in desc_lower:
            return True, ""

        expected, breakdown = self._build_breakdown(self.TOP_LEVEL_PA)
        if expected is None:
            return True, ""

        if self._close_enough(row_value, expected):
            return True, ""

        diff = row_value - expected
        return False, (
            f"Total Passif={row_value:,.0f} ≠ expected={expected:,.0f} "
            f"(diff={diff:+,.0f}). Breakdown: {breakdown}"
        )

    def _check_total_general(self, description: str, row_value: float) -> Tuple[bool, str]:
        desc_lower = description.lower()
        is_grand_total = (
            ('capitaux propres' in desc_lower and 'passif' in desc_lower and 'total' in desc_lower)
            or desc_lower.strip() == 'total'
        )
        if not is_grand_total:
            return True, ""

        cp_sum, cp_breakdown = self._build_breakdown(self.CP_BEFORE_AFFECTATION)
        pa_sum, pa_breakdown = self._build_breakdown(self.TOP_LEVEL_PA)

        if cp_sum is None or pa_sum is None:
            return True, ""

        expected = cp_sum + pa_sum
        if self._close_enough(row_value, expected):
            return True, ""

        diff = row_value - expected
        return False, (
            f"Grand Total={row_value:,.0f} ≠ expected={expected:,.0f} "
            f"(diff={diff:+,.0f}). "
            f"CP sum={cp_sum:,.0f} ({cp_breakdown}) + "
            f"PA sum={pa_sum:,.0f} ({pa_breakdown})"
        )

    # ══════════════════════════════════════════════════════════════
    # MAIN VALIDATE METHOD
    # ══════════════════════════════════════════════════════════════

    def validate(self, row: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Validate a single row.

        Args:
            row: dict with 'code', 'description', 'values' keys

        Returns:
            (passed: bool, detail: str)
            - If passed: (True, "")
            - If failed: (False, "PA3=223,660,142 ≠ children sum=261,819,765 ...")
        """
        code = str(row.get('code', '')).strip()
        description = str(row.get('description', ''))
        values = row.get('values', [])
        row_value = self._get_first_numeric_value(values)

        if row_value is None:
            return True, ""

        errors = []

        # Rule 1: Parent-child sum
        if code:
            ok, detail = self._check_parent_children(code, row_value)
            if not ok:
                errors.append(detail)

        # Rule 2-4: CP totals
        for check in [self._check_cp_avant_resultat,
                       self._check_cp_avant_affectation,
                       self._check_cp_consolide]:
            ok, detail = check(description, row_value)
            if not ok:
                errors.append(detail)

        # Rule 5-6: Passif totals
        for check in [self._check_total_passif,
                       self._check_total_general]:
            ok, detail = check(description, row_value)
            if not ok:
                errors.append(detail)

        if errors:
            return False, " | ".join(errors)
        return True, ""