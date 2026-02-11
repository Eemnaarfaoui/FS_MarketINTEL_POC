"""
PDF Downloader Module
Handles downloading PDFs from URLs
"""
import os
import re
import requests


def download_pdf(url, societe, nom, annee):
    """Download PDF from URL and save with structured filename"""
    try:
        safe_societe = re.sub(r'[^\w\s-]', '_', societe).replace(' ', '_')
        safe_nom = re.sub(r'[^\w\s-]', '_', nom).replace(' ', '_')
        filename = f"{safe_societe}_{safe_nom}_{annee}.pdf"
        filepath = os.path.join(os.getcwd(), filename)
        
        if os.path.exists(filepath):
            print(f"✓ PDF déjà existant : {filepath}")
            return filepath
        
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()
        with open(filepath, 'wb') as f:
            f.write(response.content)
        print(f"✓ PDF téléchargé : {filepath}")
        return filepath
    except Exception as e:
        print(f"❌ Erreur téléchargement : {str(e)}")
        return None


def get_local_pdf_path(societe, nom, annee):
    """Generate local PDF path without downloading"""
    safe_societe = re.sub(r'[^\w\s-]', '_', societe).replace(' ', '_')
    safe_nom = re.sub(r'[^\w\s-]', '_', nom).replace(' ', '_')
    filename = f"{safe_societe}_{safe_nom}_{annee}.pdf"
    return os.path.join(os.getcwd(), filename)
