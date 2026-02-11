"""
Document Structure Configuration for CAPITAUX PROPRES ET PASSIF
Defines the hierarchical structure and code mappings
"""

# Code to subcategory mappings for CAPITAUX PROPRES
CP_SUBCATEGORIES = {
    'CP1': 'Capital social',
    'CP2': 'Réserves et primes',
    'CP3': "Rachat d'actions",
    'CP4': 'Autres capitaux propres',
    'CP5': 'Résultat reporté',
    'CP6': "Résultat de l'exercice",
}

# Code to subcategory mappings for PASSIF
PA_SUBCATEGORIES = {
    'PA2': 'Provisions pour risques et charges',
    'PA23': 'Provisions pour risques et charges',
    'PA3': 'Provisions techniques brutes',
    'PA310': 'Provisions techniques brutes',
    'PA320': 'Provisions techniques brutes',
    'PA330': 'Provisions techniques brutes',
    'PA331': 'Provisions techniques brutes',
    'PA340': 'Provisions techniques brutes',
    'PA341': 'Provisions techniques brutes',
    'PA350': 'Provisions techniques brutes',
    'PA360': 'Provisions techniques brutes',
    'PA361': 'Provisions techniques brutes',
    'PA5': 'Dettes pour dépôts',
    'PA6': 'Autres dettes',
    'PA61': 'Autres dettes',
    'PA62': 'Autres dettes',
    'PA63': 'Autres dettes',
    'PA631': 'Autres dettes',
    'PA632': 'Autres dettes',
    'PA633': 'Autres dettes',
    'PA634': 'Autres dettes',
    'PA7': 'Autres passifs',
    'PA71': 'Comptes de régularisation',
    'PA710': 'Comptes de régularisation',
    'PA711': 'Comptes de régularisation',
    'PA712': 'Comptes de régularisation',
    'PA72': 'Écart de conversion',
}

# Parent codes (level 2 - main categories)
PARENT_CODES = ['PA2', 'PA3', 'PA5', 'PA6', 'PA7', 'PA72', 'PA71']

# Total keywords for identification
TOTAL_KEYWORDS = {
    'total capitaux propres avant résultat': ('TOTAL', 'Capitaux Propres - Avant Résultat'),
    'total capitaux propres avant affectation': ('TOTAL', 'Capitaux Propres - Avant Affectation'),
    'total du passif': ('TOTAL', 'Total Passif'),
    'total des capitaux propres et du passif': ('TOTAL GÉNÉRAL', ''),
}

def get_subcategory(code):
    """Get subcategory for a given code"""
    if code.startswith('CP'):
        return CP_SUBCATEGORIES.get(code, 'Capitaux Propres')
    elif code.startswith('PA'):
        return PA_SUBCATEGORIES.get(code, 'PASSIF')
    return ''

def is_parent_code(code):
    """Check if code is a parent (level 2) code"""
    return code in PARENT_CODES or code in CP_SUBCATEGORIES
