import sqlite3
import json
from pathlib import Path

db_path = Path("data/app.db")
conn = sqlite3.connect(str(db_path))
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Search for document 843 and check its data
print("=== Document 843 details ===")
cursor.execute("""
    SELECT id, document_number, year, created_at, output_file_path, data_json
    FROM documents 
    WHERE document_number = 843
""")
row = cursor.fetchone()

if row:
    print(f"ID: {row['id']}")
    print(f"Number: {row['document_number']}")
    print(f"Year: {row['year']}")
    print(f"Created: {row['created_at']}")
    print(f"File path: {row['output_file_path']}")
    
    # Try to parse the JSON to see the document data
    try:
        data = json.loads(row['data_json'])
        # Look for material names in the data
        if 'part_changes' in data:
            print("\n=== Materials in document ===")
            for part_change in data.get('part_changes', []):
                if 'materials' in part_change:
                    for material in part_change['materials']:
                        if material.get('is_changed'):
                            catalog_entry = material.get('catalog_entry', {})
                            before_name = catalog_entry.get('before_name', '')
                            after_name = material.get('after_name', '')
                            if 'PentriProtect' in before_name or 'PentriProtect' in after_name or 'Грунт-эмаль' in before_name or 'Грунт-эмаль' in after_name:
                                print(f"  Found: {before_name} -> {after_name}")
    except Exception as e:
        print(f"Error parsing JSON: {e}")
else:
    print("Document 843 not found!")

# Also search for any document with PentriProtect in the filename
print("\n=== Searching for PentriProtect in filenames ===")
cursor.execute("""
    SELECT document_number, year, output_file_path
    FROM documents 
    WHERE output_file_path LIKE '%PentriProtect%'
       OR output_file_path LIKE '%PANTONE19-4055%'
       OR output_file_path LIKE '%Грунт-эмаль%'
""")
rows = cursor.fetchall()
print(f"Found {len(rows)} documents with matching filenames:")
for r in rows:
    print(f"  Doc {r['document_number']}/{r['year']}: {r['output_file_path']}")

conn.close()


