from typing import Dict, Any

class ValidatorPassifs:
    def __init__(self, extracted_data_context=None, error_margin=1.0):
        self.extracted_data_context = extracted_data_context or {}
        self.error_margin = error_margin

    def _get_first_numeric_value(self, values):
        for v in values:
            try:
                f = float(v)
                return f
            except (TypeError, ValueError):
                continue
        return None

    def _validate_cp_avant_resultat(self, row: Dict[str, Any]) -> bool:
        description = row.get('description', '').lower()
        if "avant résultat" not in description:
            return True
        parent_value = self._get_first_numeric_value(row.get('values', []))
        if parent_value is None:
            return False
        components = ['CP1', 'CP2', 'CP3', 'CP4', 'CP5']
        total = 0.0
        found = False
        for code in components:
            if code in self.extracted_data_context:
                value = self.extracted_data_context[code]['value']
                if value is not None:
                    total += float(value)
                    found = True
        if not found:
            return True
        return abs(parent_value - total) <= self.error_margin

    def _validate_cp_avant_affectation(self, row: Dict[str, Any]) -> bool:
        description = row.get('description', '').lower()
        if "avant affectation" not in description:
            return True
        parent_value = self._get_first_numeric_value(row.get('values', []))
        if parent_value is None:
            return False
        components = ['CP1', 'CP2', 'CP3', 'CP4', 'CP5', 'CP6']
        total = 0.0
        found = False
        for code in components:
            if code in self.extracted_data_context:
                value = self.extracted_data_context[code]['value']
                if value is not None:
                    total += float(value)
                    found = True
        if not found:
            return True
        return abs(parent_value - total) <= self.error_margin

    def _validate_total_passif(self, row: Dict[str, Any]) -> bool:
        description = row.get('description', '').lower()
        if "total du passif" not in description:
            return True
        parent_value = self._get_first_numeric_value(row.get('values', []))
        if parent_value is None:
            return False
        total = 0.0
        found = False
        for code in self.extracted_data_context.keys():
            if len(code) == 3 and code.startswith("PA") and code[2].isdigit():
                value = self.extracted_data_context[code]['value']
                if value is not None:
                    total += float(value)
                    found = True
        if not found:
            return True
        return abs(parent_value - total) <= self.error_margin

    def _validate_total_cp_et_passif(self, row: Dict[str, Any]) -> bool:
        description = row.get('description', '').lower()
        if "capitaux propres et du passif" not in description:
            return True
        parent_value = self._get_first_numeric_value(row.get('values', []))
        if parent_value is None:
            return False
        total_passif = 0.0
        for code in self.extracted_data_context:
            if len(code) == 3 and code.startswith("PA") and code[2].isdigit():
                value = self.extracted_data_context[code]['value']
                if value is not None:
                    total_passif += float(value)
        cp_total = 0.0
        for code in ['CP1','CP2','CP3','CP4','CP5','CP6']:
            if code in self.extracted_data_context:
                value = self.extracted_data_context[code]['value']
                if value is not None:
                    cp_total += float(value)
        expected = total_passif + cp_total
        return abs(parent_value - expected) <= self.error_margin

    def validate(self, row: Dict[str, Any]) -> bool:
        # Example: run all rules, return True if all pass (customize as needed)
        return (
            self._validate_cp_avant_resultat(row)
            and self._validate_cp_avant_affectation(row)
            and self._validate_total_passif(row)
            and self._validate_total_cp_et_passif(row)
        )
