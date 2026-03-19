import sys
from pathlib import Path

# Add the project root to the path
sys.path.insert(0, str(Path(__file__).parent))

from app.document_store import DocumentStore
from app.database import DatabaseManager
from app.catalog_loader import CatalogLoader

# Initialize components
db_manager = DatabaseManager()
catalog_loader = CatalogLoader(db_manager)
document_store = DocumentStore(db_manager, catalog_loader)

# Get all documents
documents = document_store.get_all_documents()

print(f"Total documents returned: {len(documents)}")
print("\n=== All documents ===")
for doc_number, year, created_at, file_path in documents:
    file_name = Path(file_path).name if file_path else ""
    print(f"Doc {doc_number}/{year}, Created: {created_at}, File: {file_name}")
    
    # Check specifically for 843
    if doc_number == 843:
        print(f"  *** FOUND DOCUMENT 843 ***")


