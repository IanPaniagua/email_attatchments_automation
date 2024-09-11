import psycopg2

try:
    # Especifica client_encoding como 'utf-8'
    conn = psycopg2.connect(
        host='localhost',
        database='suppliers',
        user='postgresql',
        password='PostgreS',
        port=5432
    )
    
    print("Connected to the PostgreSQL server.")

    conn.close()

except Exception as ex:
    print(f"Error: {ex}")
