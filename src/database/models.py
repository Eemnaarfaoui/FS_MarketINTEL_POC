"""
Data Models
Simple data classes for document and financial data
"""

class Document:
    """Document metadata"""
    def __init__(self, id, societe, nom, annee, url):
        self.id = id
        self.societe = societe
        self.nom = nom
        self.annee = annee
        self.url = url


class FinancialData:
    """Financial data row"""
    def __init__(self, document_id, level, code, description, is_total, category, subcategory, value_n, value_n_1):
        self.document_id = document_id
        self.level = level
        self.code = code
        self.description = description
        self.is_total = is_total
        self.category = category
        self.subcategory = subcategory
        self.value_n = value_n
        self.value_n_1 = value_n_1
