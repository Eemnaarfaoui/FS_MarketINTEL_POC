import time
import sys
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
from src.extraction.validate_actif_excel import validate_actif_from_data
import subprocess
from pathlib import Path


def _build_output_dir(company_name):
    """Build and create the output directory for a company."""
    safe_name = re.sub(r'[^\w\s-]', '_', company_name).replace(' ', '_')
    if len(safe_name) > 50:
        safe_name = safe_name[:50]
    output_dir = os.path.join(os.getcwd(), "outputs", safe_name)
    os.makedirs(output_dir, exist_ok=True)
    return output_dir, safe_name


def run_extraction(company: str, year: int):
    """
    Automated narrated extraction workflow for PASSIF.
    """

    start_time = time.time()
    print(f"\n{'='*70}")
    print("🚀 EXTRACTION AUTOMATISÉE - RAPPORTS ANNUELS CMF")
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
        """
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
        """
        # ============================================================
        # SETUP OUTPUT DIRECTORY (single definition, used everywhere)
        # ============================================================
        output_dir, safe_societe = _build_output_dir(target_societe)
        short_company = re.sub(r'[^\w]', '_', company).strip('_').upper()
        print(f"📂 Dossier de sortie : {output_dir}")

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
        # 6️⃣ EXPORT PASSIF EXCEL
        # ============================================================
        print("📁 Export PASSIF vers Excel en cours...")

        passif_filename = f"{short_company}_{year}_passif.xlsx"
        passif_path = os.path.join(output_dir, passif_filename)

        result = export_to_excel(
            hierarchical_data,
            target_societe,
            pdf_path,
            passif_path,
            year,
            year - 1
        )

        if result is True:
            print(f"✅ Fichier Excel PASSIF généré : {passif_path}")

            # Validation PASSIF
            print("\n🔍 Validation des données extraites PASSIF...")
            validated_file = validate_capitaux_propres_passif(passif_path, company)
            print(f"✅ Validation PASSIF terminée : {validated_file}")
        else:
            print("⚠️ Échec export Excel PASSIF")
            if isinstance(result, str):
                print(f"Détail erreur : {result}")

        # ============================================================
        # 7️⃣ EXTRACTION & EXPORT ACTIF
        # ============================================================
        print("🔍 Recherche du tableau ACTIF dans le PDF...")
        print("Données trouvées à la page 2 (fixe pour ACTIF)")
        data_actifs = extract_actif(pdf_path, 2, is_scanned=is_scanned)

        if data_actifs:
            print(f"✅ {len(data_actifs)} lignes ACTIF extraites")

            # Export ACTIF to Excel
            print("📁 Export ACTIF vers Excel en cours...")

            actif_filename = f"{short_company}_{year}_actif.xlsx"
            actif_path = os.path.join(output_dir, actif_filename)

            export_actif_to_excel(
                data_actifs,
                actif_path,
                year,
                year - 1
            )

            print(f"✅ Fichier Excel ACTIF généré : {actif_path}")

            # Validation ACTIF
            print("\n🔍 Validation des données extraites ACTIF...")
            validated_actif_filename = f"{short_company}_{year}_actif_validated.xlsx"
            validated_actif_path = os.path.join(output_dir, validated_actif_filename)

            validated_file = validate_actif_from_data(
                data_actifs=data_actifs,
                assurance_name=target_societe,
                annee=year,
                output_xlsx=validated_actif_path
            )

            print(f"✅ Validation ACTIF terminée : {validated_file}")
        else:
            print("❌ Échec extraction ACTIF")

               # ============================================================
        # 8️⃣ INSERT FINANCIAL DATA
        # ============================================================
        if cursor is not None and connection is not None:
            print("💾 Insertion des données financières en base...")
            doc_record = get_document_by_company_year(cursor, target_societe, year)

            if doc_record:
                doc_id = doc_record[0]
                insert_financial_data_capitaux_passifs(cursor, doc_id, hierarchical_data)
                connection.commit()
                print("✅ Données financières insérées avec succès")
        else:
            print("ℹ️ DB désactivée (cursor/connection None) → insertion ignorée")

            # ============================================================
        # 9️⃣ LANCER ANNEXES 12/13 (Extraction1213 → NorVal12 → NorVal13)
        # ============================================================

        print(f"\n{'='*70}")
        print("📌 LANCEMENT ANNEXES 12 & 13 (Extraction1213 → NorVal12 → NorVal13)")
        print(f"{'='*70}")

        base_dir = Path(__file__).resolve().parent
        script_path = (base_dir / "annexes1213" / "Extraction1213.py").resolve()

        print("📍 base_dir     =", base_dir)
        print("📍 script_path  =", script_path)

        if not script_path.exists():
         raise FileNotFoundError(f"Extraction1213.py introuvable: {script_path}")

        cmd = [sys.executable, str(script_path), target_societe, str(year)]
        print("➡️ Commande:", " ".join(cmd))

        res = subprocess.run(cmd, capture_output=False)
        print(f"✅ ANNEXES 12/13 terminées, code retour = {res.returncode}")




        

        
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
        run_extraction("LLOYD TUNISIE", 2024)
    except Exception as e:
        print(f"Erreur : {e}")
   
