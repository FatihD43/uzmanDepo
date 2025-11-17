import pyodbc

SQL_CONN_STR = (
    "Driver={SQL Server};"
    "Server=10.30.9.14,1433;"
    "Database=UzmanRaporDB;"
    "UID=uzmanrapor_login;"
    "PWD=03114080Ww.;"
)

try:
    conn = pyodbc.connect(SQL_CONN_STR, timeout=5)
    print("SQL Server bağlantısı başarılı!")
    conn.close()
except Exception as e:
    print("Bağlantı hatası:", e)
