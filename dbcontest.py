import pyodbc

connection = pyodbc.connect(
    'DRIVER={ODBC Driver 18 for SQL Server};'
    'SERVER=localhost;'
    'Trusted_Connection=yes;'
    'Encrypt=no;'
)
cursor = connection.cursor()
cursor.execute("SELECT @@version")
row = cursor.fetchone()
print("Connected to:", row[0])
connection.close()
