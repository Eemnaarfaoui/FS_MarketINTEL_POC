import time
import os
import re
import logging
# Import modules
from src.extraction.validate_passif_excel import validate_capitaux_propres_passif
from src.scraper.cmf_scraper import init_driver, get_all_companies, select_company_and_submit, scrape_document_list
from src.scraper.pdf_downloader import download_pdf, get_local_pdf_path
from src.extraction.pdf_parser import search_table_in_pdf, extract_table_from_page, extract_passif 
from src.extraction.excel_exporter import export_to_excel
from src.database.db_manager import create_database_and_tables, insert_document, insert_financial_data_capitaux_passifs, get_document_by_company_year
from src.extraction.extract_actifs import extract_actif
from src.extraction.excel_exporter_actif import export_actif_to_excel
def run_extraction(company: str, year: int):
    """
    Automated narrated extraction workflow for PASSIF.
    """

    start_time = time.time()
    print(f"\n{'='*70}")
    print("🚀 EXTRACTION AUTOMATISÉE - CAPITAUX PROPRES ET PASSIF")
    print(f"{'='*70}")
    print(f"🏢 Société cible : {company}")
    print(f"📅 Année cible   : {year}")
    print(f"{'-'*70}")

    driver = None
    connection = None
    cursor = None

    try:
        # ============================================================
        # 1️⃣ INITIALIZE DRIVER
        # ============================================================
        print("🌐 Initialisation du navigateur...")
        driver = init_driver()

        print("🔎 Récupération des sociétés disponibles...")
        available_companies = get_all_companies(driver)

        matches = [c for c in available_companies if company.lower() in c.lower()]
        if not matches:
            print(f"❌ Société non trouvée : {company}")
            return

        target_societe = matches[0]
        print(f"✅ Société trouvée : {target_societe}")

        # ============================================================
        # 2️⃣ LOAD DOCUMENTS
        # ============================================================
        print("📂 Chargement des documents CMF...")
        if not select_company_and_submit(driver, target_societe):
            print("❌ Échec soumission formulaire")
            return

        all_documents = scrape_document_list(driver, target_societe)

        year_documents = [doc for doc in all_documents if str(doc['annee']) == str(year)]
        if not year_documents:
            print(f"❌ Aucun document trouvé pour {year}")
            return

        selected_doc = year_documents[0]
        print(f"✅ Document sélectionné : {selected_doc['nom']}")

        # ============================================================
        # 3️⃣ DOWNLOAD PDF
        # ============================================================
        print("⬇️ Téléchargement du PDF...")
        pdf_path = download_pdf(
            selected_doc['url'],
            selected_doc['societe'],
            selected_doc['nom'],
            selected_doc['annee']
        )

        if not pdf_path:
            print("❌ Échec téléchargement")
            return

        print(f"✅ PDF téléchargé : {os.path.basename(pdf_path)}")

        # ============================================================
        # 4️⃣ DATABASE CONNECTION
        # ============================================================
        print("🗄️ Connexion à la base de données...")
        connection, cursor = create_database_and_tables()

        if not connection:
            print("❌ Échec connexion DB")
            return

        insert_document(
            connection,
            cursor,
            target_societe,
            selected_doc['nom'],
            year,
            selected_doc['url']
        )

        print("✅ Métadonnées document enregistrées")

        # ============================================================
        # 5️⃣ SEARCH & EXTRACT PASSIF
        # ============================================================
        print("🔍 Recherche du tableau PASSIF dans le PDF...")
        page_num, is_scanned = search_table_in_pdf(pdf_path, "passif")

        if not page_num:
            print("❌ PASSIF non trouvé dans le document")
            return

        print(f"✅ PASSIF trouvé à la page {page_num}")
        print("📊 Extraction et structuration des données...")

        hierarchical_data = extract_passif(pdf_path, page_num, is_scanned)

        if not hierarchical_data:
            print("❌ Échec extraction PASSIF")
            return

        print(f"✅ {len(hierarchical_data)} lignes structurées extraites")

        # ============================================================
        # 6️⃣ EXPORT EXCEL
        # ============================================================

        print("📁 Export vers Excel en cours...")

        safe_societe = re.sub(r'[^\w\s-]', '_', target_societe).replace(' ', '_')
        safe_nom = re.sub(r'[^\w\s-]', '_', selected_doc['nom']).replace(' ', '_')

        output_name = f"{safe_societe}_{year}_passif_{safe_nom}.xlsx"

        result = export_to_excel(
            hierarchical_data,
            target_societe,
            pdf_path,
            output_name,
            year,
            year - 1
        )
        if result is True:
            print(f"✅ Fichier Excel généré : {output_name}")
            excel_path = os.path.join(os.getcwd(), "outputs", safe_societe, output_name)
            safe_societe = "".join(c if c.isalnum() or c in " _-" else "_" for c in target_societe)
            if len(safe_societe) > 30:
                    safe_societe = safe_societe[:27] + "_"

            excel_path = os.path.join(os.getcwd(), "outputs", safe_societe, output_name)
        # Validation du fichier Excel généré
            print("\n🔍 Validation des données extraites PASSIF...")
            validated_file = validate_capitaux_propres_passif(excel_path, target_societe)
            print(f"✅ Validation terminée, fichier sauvegardé : {validated_file}")
        else:
            print("⚠️ Échec export Excel")
            if isinstance(result, str):
                print(f"Détail erreur : {result}")

        # ============================================================
        # 7️⃣ EXTRACTION & VALIDATION DES ACTIFS
        # ============================================================
        print("🔍 Recherche du tableau ACTIF dans le PDF...")
        print("Données trouvées à la page 2 (fixe pour ACTIF)")
        data_actifs = extract_actif(pdf_path, 2, is_scanned)
        if data_actifs:
            print(f"✅ {len(data_actifs)} lignes ACTIF extraites")
            # Export ACTIF to Excel
            print("📁 Export ACTIF vers Excel en cours...")
            export_actif_to_excel(
                data_actifs,   
                
                f"{re.sub(r'[^\w\s-]', '_', target_societe).replace(' ', '_')}_{year}_actif_{re.sub(r'[^\w\s-]', '_', selected_doc['nom']).replace(' ', '_')}.xlsx",
                year,
                year - 1
            )
            print(f"✅ Fichier Excel ACTIF généré : {target_societe}_{year}_actif_{selected_doc['nom']}.xlsx")
        else:
            print("❌ Échec extraction ACTIF")


        


        







        # ============================================================
        # 6️⃣ INSERT FINANCIAL DATA
        # ============================================================
        print("💾 Insertion des données financières en base...")
        doc_record = get_document_by_company_year(cursor, target_societe, year)

        if doc_record:
            doc_id = doc_record[0]
            insert_financial_data_capitaux_passifs(cursor, doc_id, hierarchical_data)
            connection.commit()
            print("✅ Données financières insérées avec succès")

        
  

        elapsed = time.time() - start_time
        print(f"\n{'='*70}")
        print(f"🎉 EXTRACTION TERMINÉE EN {elapsed:.2f} secondes")
        print(f"{'='*70}")


    except Exception as e:
        logging.error(f"ERREUR GLOBALE : {str(e)}")
        print(f"\n❌ ERREUR GLOBALE : {str(e)}")

    finally:
        if driver:
            driver.quit()
        if connection:
            cursor.close()
            connection.close()
            print("🔒 Connexion fermée")


if __name__ == "__main__":
    try:
        run_extraction("Comar", 2024)
    except Exception as e:
        print(f"Erreur : {e}")
