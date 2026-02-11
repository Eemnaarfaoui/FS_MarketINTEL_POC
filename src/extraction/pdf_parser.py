"""
PDF Parser Module
Handles PDF text extraction and table detection
"""
import re
import unicodedata
import PyPDF2
import camelot
import pytesseract
import difflib
from pdf2image import convert_from_path
from PIL import Image

from src.extraction.hierarchy_detector_passif import detect_hierarchy_level_passif, structure_hierarchical_data_passif
def search_table_in_pdf(pdf_path, table_type):

    keywords = {
        'passif': ["capitaux propres", "passif"],
        'actif': ["actif","AC1 Actifs incorporels"],
        'ann12': ["annexe 12", "engagements"],
        'ann13': ["annexe 13", "engagements"]
    }

    target_keywords = keywords.get(table_type, keywords['passif'])

    try:
        print(f"\n🔍 Recherche de la table: {table_type.upper()}...")

        pdf_reader = PyPDF2.PdfReader(pdf_path)
        num_pages = len(pdf_reader.pages)

        for page_num in range(num_pages):

            page = pdf_reader.pages[page_num]
            text = page.extract_text()

            # Decide if page is scanned
            if not text or len(text.strip()) < 100:
                # Probably scanned
                print(f"📷 Page {page_num+1} seems scanned → OCR...")
                image = convert_from_path(
                    pdf_path,
                    first_page=page_num+1,
                    last_page=page_num+1,
                    dpi=300
                )[0]

                text = pytesseract.image_to_string(
                    image,
                    lang='fra',
                    config='--oem 3 --psm 6'
                )

                is_scanned = True
            else:
                is_scanned = False

            text_lower = text.lower()
            text_lower = re.sub(r'\s+', ' ', text_lower)

            # Stronger detection rule
            if (
                "capitaux propres" in text_lower
                and "passif" in text_lower
                and "total" in text_lower
            ):
                print(f" {table_type.upper()} trouvé à la page {page_num + 1}")
                return page_num + 1, is_scanned

        print(f" {table_type.upper()} non trouvé")
        return None, None

    except Exception as e:
        print(f" Erreur : {str(e)}")
        return None, None



def extract_table_from_page(pdf_path, page_num, is_scanned, table_type):
    """
    Extract raw table data from a specific page based on table_type
    Returns list of rows (each row is a list of cell values)
    """
    # Define markers for each table type
    markers = {
        'passif': {
            'start': ["capitaux propres et", "passif"],
            'end': ["total des capitaux propres et du passif", "total des capitaux propres et du passifs","total des capitaux "]
        },
        'actif': {
            'start': ["actif", "actifs"],
            'end': ["total de l'actif", "total de l actifs"]
        },
        'ann12': {
            'start': ["annexe 12", "engagements"],
            'end': ["total engagements donnes", "total des engagements donnes"]
        },
        'ann13': {
            'start': ["annexe 13", "engagements"],
            'end': ["total engagements reus", "total des engagements reus"]
        }
    }
    
    current_markers = markers.get(table_type, markers['passif'])

    try:
        print(f"\n Extraction du tableau {table_type.upper()} page {page_num}...")
    

        
        if is_scanned:
            # OCR extraction
            images = convert_from_path(pdf_path, first_page=page_num, 
                                       last_page=page_num, dpi=500, fmt='jpeg')
            if not images:
                return None
            
            image = images[0]
            text = pytesseract.image_to_string(image, lang='fra', config='--oem 3 --psm 6')
            text = text.replace('|', ' ')
            text = text.replace('—', '-')


            # Parse OCR text line by line
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            structured_data = []
            
            for line in lines:
                # Split by multiple spaces
                numbers = re.findall(r'[-+]?\d[\d\s,.]*', line)
                parts = re.split(r'\s{2,}', line)
                if len(parts) >= 2:
                    structured_data.append(parts)
                if numbers:
            # Remove numbers from description
                    desc = line
                    for num in numbers:
                        desc = desc.replace(num, '').strip()

                    row = [desc] + numbers
                    structured_data.append(row)
                else:
            # If no numbers detected, keep full line as one column
                    structured_data.append([line])

        else:
            # Native extraction with Camelot
            tables = camelot.read_pdf(pdf_path, flavor='stream', pages=str(page_num))
            
            if tables.n == 0:
                print(" Aucun tableau détecté")
                return None
            
            # Concatenate all tables detected on this page
            structured_data = []
            for i in range(tables.n):
                structured_data.extend(tables[i].df.values.tolist())
        
        # ==========================================
        #  Filter by Boundary
        # ==========================================
        filtered_data = []
        found_start = False
        
        for row in structured_data:
            combined_row = " ".join(str(cell) for cell in row).lower()
            
            if not found_start:
                # Detect start
                if all(s in combined_row for s in current_markers['start']):
                    print(f"Table start detected: {combined_row[:50]}...")
                    found_start = True
                    filtered_data.append(row)
            else:
                filtered_data.append(row)
                # Detect end
                if any(e in combined_row for e in current_markers['end']):
                    print(f"🏁 Table end detected: {combined_row[:50]}...")
                    break
        
        if filtered_data:
            structured_data = filtered_data
            print(f" Filtered to {len(structured_data)} relevant rows")
        elif found_start:
             print(" Start found but no data appended (this shouldn't happen)")
        else:
             print(f"⚠️ '{table_type.upper()}' header not found in raw data, using all rows")

        print(f" {len(structured_data)} lignes brutes extraites")
        return structured_data
        
    except Exception as e:
        print(f" Erreur extraction : {str(e)}")
        return None

        
    except Exception as e:
        print(f" Erreur extraction : {str(e)}")
        return None



































def extract_passif(pdf_path, page_num, is_scanned):
    """
    Dedicated extraction logic for CAPITAL PROPRES ET PASSIF
    """
    
    
    raw_data = extract_table_from_page(pdf_path, page_num, is_scanned, 'passif')
    if not raw_data:
        return None
        
    return structure_hierarchical_data_passif(raw_data)


def extract_actif(pdf_path, page_num, is_scanned):
    """
    Placeholder for ACTIF extraction logic
    """
    print("Extraction de l'ACTIF (Logique à implémenter)")
    # For now, just extract raw and try basic structuring
    raw_data = extract_table_from_page(pdf_path, page_num, is_scanned, 'actif')
    if not raw_data:
        return None
        
    # We use a placeholder structuring for now
    return [{'level': 2, 'code': '', 'description': 'Logic ACTIF à implémenter', 'is_total': False, 'category': 'ACTIF', 'subcategory': '', 'values': []}]


def extract_ann12(pdf_path, page_num, is_scanned):
    """
    Placeholder for ANNEXE 12 extraction logic
    """
    print("ℹ️ Extraction de l'ANNEXE 12 (Logique à implémenter)")
    return [{'level': 2, 'code': '', 'description': 'Logic ANNEXE 12 à implémenter', 'is_total': False, 'category': 'ANNEXE 12', 'subcategory': '', 'values': []}]


def extract_ann13(pdf_path, page_num, is_scanned):
    """
    Placeholder for ANNEXE 13 extraction logic
    """
    print("ℹ️ Extraction de l'ANNEXE 13 (Logique à implémenter)")
    return [{'level': 2, 'code': '', 'description': 'Logic ANNEXE 13 à implémenter', 'is_total': False, 'category': 'ANNEXE 13', 'subcategory': '', 'values': []}]


