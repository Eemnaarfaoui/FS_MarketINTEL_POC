"""
Database Manager Module
Handles database connections and operations
"""
import pyodbc
import logging
from src.utils.helpers import normalize_url


def create_database_and_tables():
    """Create CMF database and required tables"""
    try:
        print("\n🔧 Création/mise à jour de la base de données 'cmf'...")
        
        # Connect to SQL Server
        connection = pyodbc.connect(
            'DRIVER={SQL Server};'
            'SERVER=localhost;'
            'DATABASE=master;'
            'Trusted_Connection=yes;'
        )
        connection.autocommit = True
        cursor = connection.cursor()
        
        # Create database if not exists
        cursor.execute("IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'cmf') CREATE DATABASE cmf")
        cursor.close()
        connection.close()
        
        # Connect to cmf database
        connection = pyodbc.connect(
            'DRIVER={SQL Server};'
            'SERVER=localhost;'
            'DATABASE=cmf;'
            'Trusted_Connection=yes;'
        )
        cursor = connection.cursor()
        
        # Create document table
        cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='document' AND xtype='U')
        CREATE TABLE document (
            id INT IDENTITY(1,1) PRIMARY KEY,
            Societe NVARCHAR(255),
            Nom NVARCHAR(255),
            Annee INT,
            URL NVARCHAR(MAX)
        )
        """)
        
        # Create financial_data_capitaux_passifs table
        cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='financial_data_capitaux_passifs' AND xtype='U')
        CREATE TABLE financial_data_capitaux_passifs (
            id INT IDENTITY(1,1) PRIMARY KEY,
            document_id INT,
            level INT,
            code NVARCHAR(50),
            description NVARCHAR(MAX),
            is_total BIT,
            category NVARCHAR(100),
            subcategory NVARCHAR(100),
            value_n BIGINT,
            value_n_1 BIGINT,
            FOREIGN KEY (document_id) REFERENCES document(id)
        )
        """)
        
        connection.commit()
        print("✅ Base de données 'cmf' prête")
        return connection, cursor
        
    except Exception as e:
        logging.error(f"Erreur lors de la création de la base : {e}")
        print(f"❌ Erreur lors de la création de la base : {e}")
        return None, None


def check_document_exists(cursor, societe, nom_document, annee):
    """Check if document already exists in database"""
    try:
        query = "SELECT COUNT(*) FROM document WHERE Societe = ? AND Nom = ? AND Annee = ?"
        cursor.execute(query, (societe, nom_document, annee))
        count = cursor.fetchone()[0]
        return count > 0
    except Exception as e:
        logging.error(f"Erreur check_document_exists : {e}")
        return False


def insert_document(connection, cursor, societe, nom_document, annee, url):
    """Insert document metadata into database"""
    try:
        normalized_url = normalize_url(url)
        
        try:
            annee_int = int(annee)
            if not (2015 <= annee_int <= 2030):
                return False
        except ValueError:
            return False
            
        if check_document_exists(cursor, societe, nom_document, annee_int):
            return False
            
        insert_query = """
        INSERT INTO document (Societe, Nom, Annee, URL)
        VALUES (?, ?, ?, ?)
        """
        cursor.execute(insert_query, (societe, nom_document, annee_int, normalized_url))
        connection.commit()
        return True
        
    except Exception as e:
        logging.error(f"Erreur lors de l'insertion : {e}")
        return False


def insert_financial_data_capitaux_passifs(cursor, doc_id, hierarchical_data):
    """Insert extracted financial data into database"""
    try:
        print(f"\n Insertion des données financières pour le document {doc_id}...")
        
        # Delete old data for this document to avoid duplicates
        cursor.execute("DELETE FROM financial_data_capitaux_passifs WHERE document_id = ?", (doc_id,))
        
        insert_query = """
        INSERT INTO financial_data_capitaux_passifs 
        (document_id, level, code, description, is_total, category, subcategory, value_n, value_n_1)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        
        for item in hierarchical_data:
            values = item.get('values', [])
            value_n = values[0] if len(values) > 0 and isinstance(values[0], int) else None
            value_n_1 = values[1] if len(values) > 1 and isinstance(values[1], int) else None
            
            cursor.execute(insert_query, (
                doc_id,
                item['level'],
                item['code'],
                item['description'],
                item['is_total'],
                item['category'],
                item['subcategory'],
                value_n,
                value_n_1
            ))
        
        cursor.connection.commit()
        print(f" {len(hierarchical_data)} lignes insérées dans la base de données")
        return True
        
    except Exception as e:
        logging.error(f"Erreur insertion financial_data_capitaux_passifs : {e}")
        print(f" Erreur insertion : {e}")
        return False


def get_document_by_company_year(cursor, societe, annee):
    """Get document by company and year"""
    try:
        query = """
        SELECT id, Societe, Nom, Annee, URL 
        FROM document 
        WHERE Societe = ? 
        AND Annee = ?
        AND Nom LIKE ?
        """
        cursor.execute(query, (societe, int(annee), '%Etats financiers%'))
        
        results = cursor.fetchall()
        target_doc = None
        
        # Filter for 31/12 documents
        for res in results:
            if "31/12" in res[2]:  # res[2] is Nom
                target_doc = res
                break
        
        if not target_doc and results:
            target_doc = results[0]  # Fallback to first found
            
        return target_doc
        
    except Exception as e:
        logging.error(f"Erreur get_document : {e}")
        return None
