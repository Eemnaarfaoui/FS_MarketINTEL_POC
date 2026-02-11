"""
Main Entry Point for Financial Data Extraction
Interactive CLI for extracting CAPITAUX PROPRES ET PASSIF from CMF documents
"""
import time
import os
import re
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('extraction.log'),
        logging.StreamHandler()
    ]
)

# Import modules
from src.scraper.cmf_scraper import init_driver, get_all_companies, select_company_and_submit, scrape_document_list
from src.scraper.pdf_downloader import download_pdf, get_local_pdf_path
from src.extraction.pdf_parser import search_table_in_pdf, extract_table_from_page, extract_passif, extract_actif, extract_ann12, extract_ann13
from src.extraction.excel_exporter import export_to_excel
from src.database.db_manager import create_database_and_tables, insert_document, insert_financial_data_capitaux_passifs, get_document_by_company_year


def main():
    """Main interactive workflow"""
    start_time = time.time()
    print(f"\n{'='*70}")
    print(f"🚀 EXTRACTION STRUCTURÉE - CAPITAUX PROPRES ET PASSIF (INTERACTIF)")
    print(f"{'='*70}")
    
    driver = None
    connection = None
    cursor = None
    
    try:
        # Step 1: Initialize driver and get companies
        driver = init_driver()
        available_companies = get_all_companies(driver)
        
        if not available_companies:
            print("❌ Impossible de récupérer la liste des sociétés.")
            return
        
        # Step 2: Interactive company selection
        target_societe = None
        while not target_societe:
            search_query = input("\n🏢 Entrez le nom (ou partie du nom) de la société : ").strip().lower()
            if not search_query:
                continue
                
            matches = [c for c in available_companies if search_query in c.lower()]
            
            if not matches:
                print("❌ Aucune société trouvée.")
                continue
            
            print(f"\nSociétés trouvées ({len(matches)}) :")
            for i, match in enumerate(matches, 1):
                print(f"  [{i}] {match}")
            
            choice = input(f"👉 Sélectionnez le numéro de la société (1-{len(matches)}) ou 0 pour rechercher à nouveau : ").strip()
            if choice.isdigit():
                idx = int(choice) - 1
                if 0 <= idx < len(matches):
                    target_societe = matches[idx]
                    print(f"✅ Société sélectionnée : {target_societe}")
                elif idx == -1:
                    continue
            else:
                print("⚠️ Choix invalide.")
        
        # Step 3: Year selection
        target_annee = None
        while not target_annee:
            year_input = input("\n📅 Entrez l'année (ex: 2024) : ").strip()
            if year_input.isdigit() and 2010 <= int(year_input) <= 2030:
                target_annee = int(year_input)
            else:
                print("⚠️ Année invalide.")
        
        # Step 4: Select company and submit form
        if not select_company_and_submit(driver, target_societe):
            return
        
        # Step 5: Scrape documents
        all_documents = scrape_document_list(driver, target_societe)
        
        # Filter by year
        year_documents = [doc for doc in all_documents if str(doc['annee']) == str(target_annee)]
        
        if not year_documents:
            print(f"❌ Aucun document trouvé pour l'année {target_annee}.")
            return
        
        print(f"\n📂 Documents trouvés pour {target_annee} :")
        for i, doc in enumerate(year_documents, 1):
            print(f"  [{i}] {doc['nom']}")
        
        # Step 6: Document selection
        selected_doc = None
        while not selected_doc:
            choice = input("\n👉 Choisissez le numéro du document à télécharger (ou 0 pour annuler) : ").strip()
            if choice == '0':
                return
            if choice.isdigit():
                idx = int(choice) - 1
                if 0 <= idx < len(year_documents):
                    selected_doc = year_documents[idx]
                    print(f"✅ Document sélectionné : {selected_doc['nom']}")
                else:
                    print("⚠️ Choix invalide.")
            else:
                print("⚠️ Choix invalide.")
        
        # Step 7: Download PDF
        pdf_path = download_pdf(
            selected_doc['url'],
            selected_doc['societe'],
            selected_doc['nom'],
            selected_doc['annee']
        )
        
        if not pdf_path:
            print("❌ Erreur de téléchargement.")
            return
        
        # Step 8: Database operations
        connection, cursor = create_database_and_tables()
        if not connection or not cursor:
            print("❌ Échec de la connexion à la base de données")
            return
        
        # Insert document metadata
        insert_document(
            connection,
            cursor,
            target_societe,
            selected_doc['nom'],
            target_annee,
            selected_doc['url']
        )
        
        # Step 9: Select Table Type
        print("\n📊 Sélectionnez le tableau à extraire :")
        print("  [1] CAPITAUX PROPRES ET PASSIF")
        print("  [2] ACTIF")
        print("  [3] ANNEXE 12 (Engagements donnés)")
        print("  [4] ANNEXE 13 (Engagements reçus)")
        
        table_choice = input("\n👉 Sélectionnez le numéro du tableau (1-4) : ").strip()
        table_map = {
            '1': 'passif',
            '2': 'actif',
            '3': 'ann12',
            '4': 'ann13'
        }
        table_type = table_map.get(table_choice, 'passif')
        table_label = table_type.upper()

        # Step 10: Extract data from PDF
        page_num, is_scanned = search_table_in_pdf(pdf_path, table_type)
        
        if not page_num:
            print(f"⚠️ {table_label} non trouvé dans le document")
            return
        
        # Extract table using specialized functions
        if table_type == 'passif':
            hierarchical_data = extract_passif(pdf_path, page_num, is_scanned)
        elif table_type == 'actif':
            hierarchical_data = extract_actif(pdf_path, page_num, is_scanned)
        elif table_type == 'ann12':
            hierarchical_data = extract_ann12(pdf_path, page_num, is_scanned)
        elif table_type == 'ann13':
            hierarchical_data = extract_ann13(pdf_path, page_num, is_scanned)
        else:
            hierarchical_data = None
        
        if not hierarchical_data:
            print(f"❌ Échec de l'extraction ou de la structuration pour {table_label}")
            return
        
        print(f"✅ {len(hierarchical_data)} lignes structurées extraites")
        
        # Step 11: Insert into database
        doc_record = get_document_by_company_year(cursor, target_societe, target_annee)
        if doc_record:
            doc_id = doc_record[0]
            insert_financial_data_capitaux_passifs(cursor, doc_id, hierarchical_data)
        
        # Step 12: Export to Excel
        safe_societe = re.sub(r'[^\w\s-]', '_', target_societe).replace(' ', '_')
        safe_nom = re.sub(r'[^\w\s-]', '_', selected_doc['nom']).replace(' ', '_')
        output_name = f"{safe_societe}_{target_annee}_{table_type}_{safe_nom}.xlsx"
        
        if export_to_excel(hierarchical_data, target_societe,pdf_path, output_name, target_annee, target_annee - 1):
            print(f"\n{'='*70}")
            print(f"✅ EXTRACTION RÉUSSIE")
            print(f"📁 Fichier : {output_name}")
            print(f"📊 Lignes extraites : {len(hierarchical_data)}")
            print(f"{'='*70}")
        else:
            print("Échec de l'exportation vers Excel")
        
        elapsed = time.time() - start_time
        print(f"\n{'='*70}")
        print(f"✅ TERMINÉ en {elapsed:.2f}s")
        print(f"{'='*70}\n")
        
    except Exception as e:
        logging.error(f"ERREUR GLOBALE : {str(e)}")
        print(f"\n❌ ERREUR GLOBALE : {str(e)}")
    
    finally:
        if driver:
            driver.quit()
        if connection:
            cursor.close()
            connection.close()
            print("✅ Connexion fermée")


if __name__ == "__main__":
    try:
        # Check required packages
        import camelot
        import pandas
        import fitz
        import openpyxl
        import pytesseract
        from pdf2image import convert_from_path
        from PIL import Image
        
        main()
    except ImportError as e:
        print(f"\n❌ Packages manquants : {str(e)}")
        print("Installez-les avec : pip install camelot-py pandas pymupdf openpyxl pytesseract pdf2image pillow selenium webdriver-manager pyodbc requests")
        input("Appuyez sur Entrée pour quitter...")
