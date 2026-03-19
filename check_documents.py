import sqlite3
from pathlib import Path

db_path = Path("data/app.db")
if not db_path.exists():
    print(f"Database not found at {db_path}")
    exit(1)

conn = sqlite3.connect(str(db_path))
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Search for document 843
print("=== Searching for document 843 ===")
cursor.execute("""
    SELECT document_number, year, created_at, output_file_path 
    FROM documents 
    WHERE document_number = 843
    ORDER BY year DESC, document_number DESC
""")
rows = cursor.fetchall()
print(f"Found {len(rows)} documents with number 843:")
for r in rows:
    print(f"  Doc {r['document_number']}/{r['year']}, Created: {r['created_at']}, File: {r['output_file_path']}")

# Search by filename pattern
print("\n=== Searching by filename pattern ===")
cursor.execute("""
    SELECT document_number, year, created_at, output_file_path 
    FROM documents 
    WHERE output_file_path LIKE '%843%' 
       OR output_file_path LIKE '%Грунт-эмаль%'
       OR output_file_path LIKE '%PentriProtect%'
    ORDER BY year DESC, document_number DESC
""")
rows = cursor.fetchall()
print(f"Found {len(rows)} documents matching filename patterns:")
for r in rows:
    print(f"  Doc {r['document_number']}/{r['year']}, Created: {r['created_at']}, File: {r['output_file_path']}")

# Show all documents
print("\n=== All documents (last 30) ===")
cursor.execute("""
    SELECT document_number, year, created_at, output_file_path 
    FROM documents 
    ORDER BY year DESC, document_number DESC
    LIMIT 30
""")
rows = cursor.fetchall()
print(f"Total documents (showing last 30):")
for r in rows:
    print(f"  Doc {r['document_number']}/{r['year']}, Created: {r['created_at']}, File: {r['output_file_path']}")

conn.close()


