"""
Скрипт для визуализации данных из базы в древовидной структуре
"""
import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QLabel, QSplitter, QTextEdit
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from app.database import DatabaseManager
from app.product_store import ProductStore
from app.catalog_loader import CatalogLoader


class DataTreeViewer(QMainWindow):
    """Окно для просмотра данных в древовидной структуре"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Просмотр данных - Древовидная структура")
        self.setGeometry(100, 100, 1200, 800)
        
        # Инициализация БД
        self.db_manager = DatabaseManager()
        self.db_manager.initialize()
        
        # Менеджеры данных
        self.product_store = ProductStore(self.db_manager)
        self.catalog_loader = CatalogLoader(self.db_manager)
        
        self.init_ui()
        self.load_data()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        
        # Заголовок
        title = QLabel("Структура данных")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # Создаём сплиттер для дерева и детальной информации
        splitter = QSplitter(Qt.Horizontal)
        
        # Дерево данных
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Элемент", "Дополнительная информация"])
        self.tree.setColumnWidth(0, 400)
        self.tree.setColumnWidth(1, 600)
        self.tree.itemExpanded.connect(self.on_item_expanded)
        self.tree.itemSelectionChanged.connect(self.on_item_selected)
        splitter.addWidget(self.tree)
        
        # Область для детальной информации
        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setFont(QFont("Consolas", 10))
        splitter.addWidget(self.details_text)
        
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        
        layout.addWidget(splitter)
        
        # Статистика
        self.stats_label = QLabel()
        layout.addWidget(self.stats_label)
    
    def load_data(self):
        """Загрузить данные из БД и построить дерево"""
        self.tree.clear()
        
        # Загружаем данные
        products = self.product_store.get_all_products()
        catalog_entries = self.catalog_loader.load()
        all_parts = self.catalog_loader.get_all_parts()
        
        # Группируем материалы по деталям
        materials_by_part = {}
        for entry in catalog_entries:
            if entry.part not in materials_by_part:
                materials_by_part[entry.part] = []
            materials_by_part[entry.part].append(entry)
        
        # Группируем детали по изделиям
        parts_by_product = {}
        for product_id, product_name in products:
            parts = self.product_store.get_parts_by_product(product_id)
            if parts:
                parts_by_product[product_id] = (product_name, parts)
        
        # Создаём корневой элемент
        root = self.tree.invisibleRootItem()
        
        # 1. Изделия с деталями
        products_item = QTreeWidgetItem(root, ["📦 Изделия", f"Всего: {len(products)}"])
        products_item.setExpanded(True)
        
        for product_id, product_name in products:
            product_item = QTreeWidgetItem(products_item, [f"🔧 {product_name}", f"ID: {product_id}"])
            
            # Получаем детали для этого изделия
            parts = self.product_store.get_parts_by_product(product_id)
            if parts:
                for part_code in sorted(parts):
                    part_item = QTreeWidgetItem(product_item, [f"⚙️ {part_code}", ""])
                    
                    # Добавляем материалы для этой детали
                    if part_code in materials_by_part:
                        materials = materials_by_part[part_code]
                        for entry in materials:
                            material_text = f"{entry.workshop} | {entry.role} | {entry.before_name}"
                            details = f"Ед.изм.: {entry.unit}, Норма: {entry.norm}"
                            if entry.comment:
                                details += f", Прим.: {entry.comment}"
                            if entry.is_part_of_set:
                                details += f" [В наборе: {entry.replacement_set_id}]"
                            
                            material_item = QTreeWidgetItem(part_item, [material_text, details])
                            material_item.setData(0, Qt.UserRole, entry)
        
        # 2. Детали без привязки к изделиям
        all_linked_parts = set()
        for product_id, product_name in products:
            parts = self.product_store.get_parts_by_product(product_id)
            all_linked_parts.update(parts)
        
        unlinked_parts = [p for p in all_parts if p not in all_linked_parts]
        if unlinked_parts:
            unlinked_item = QTreeWidgetItem(root, ["📋 Детали без привязки к изделиям", f"Всего: {len(unlinked_parts)}"])
            for part_code in sorted(unlinked_parts):
                part_item = QTreeWidgetItem(unlinked_item, [f"⚙️ {part_code}", ""])
                
                if part_code in materials_by_part:
                    materials = materials_by_part[part_code]
                    for entry in materials:
                        material_text = f"{entry.workshop} | {entry.role} | {entry.before_name}"
                        details = f"Ед.изм.: {entry.unit}, Норма: {entry.norm}"
                        if entry.comment:
                            details += f", Прим.: {entry.comment}"
                        if entry.is_part_of_set:
                            details += f" [В наборе: {entry.replacement_set_id}]"
                        
                        material_item = QTreeWidgetItem(part_item, [material_text, details])
                        material_item.setData(0, Qt.UserRole, entry)
        
        # 3. Наборы материалов для замены
        self.load_replacement_sets(root)
        
        # 4. Документы
        self.load_documents(root)
        
        # Статистика
        total_materials = len(catalog_entries)
        total_parts = len(all_parts)
        stats_text = (
            f"Статистика: Изделий: {len(products)} | "
            f"Деталей: {total_parts} | "
            f"Материалов: {total_materials}"
        )
        self.stats_label.setText(stats_text)
    
    def load_replacement_sets(self, parent):
        """Загрузить наборы материалов для замены"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            
            # Получаем наборы
            cursor.execute("""
                SELECT id, part_code, set_type, set_name, created_at
                FROM material_replacement_sets
                ORDER BY part_code, set_type
            """)
            sets = cursor.fetchall()
            
            if sets:
                sets_item = QTreeWidgetItem(parent, ["🔄 Наборы материалов для замены", f"Всего: {len(sets)}"])
                
                for set_row in sets:
                    set_id = set_row['id']
                    part_code = set_row['part_code']
                    set_type = set_row['set_type']
                    set_name = set_row['set_name'] or ""
                    created_at = set_row['created_at'] or ""
                    
                    set_type_text = "Заменить" if set_type == 'from' else "Заменить на"
                    set_item_text = f"{set_type_text} | {part_code}"
                    if set_name:
                        set_item_text += f" ({set_name})"
                    
                    set_item = QTreeWidgetItem(sets_item, [set_item_text, f"ID: {set_id}, Создан: {created_at}"])
                    
                    # Получаем элементы набора
                    cursor.execute("""
                        SELECT msi.order_index, ce.workshop, ce.role, ce.before_name, 
                               ce.unit, ce.norm, ce.comment
                        FROM material_set_items msi
                        JOIN catalog_entries ce ON msi.catalog_entry_id = ce.id
                        WHERE msi.set_id = ?
                        ORDER BY msi.order_index
                    """, (set_id,))
                    
                    items = cursor.fetchall()
                    for item_row in items:
                        material_text = (
                            f"{item_row['workshop']} | {item_row['role']} | "
                            f"{item_row['before_name']}"
                        )
                        details = f"Ед.изм.: {item_row['unit']}, Норма: {item_row['norm']}"
                        if item_row['comment']:
                            details += f", Прим.: {item_row['comment']}"
                        
                        item = QTreeWidgetItem(set_item, [material_text, details])
    
    def load_documents(self, parent):
        """Загрузить документы"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT id, document_number, year, created_at, updated_at
                FROM documents
                ORDER BY year DESC, document_number DESC
            """)
            documents = cursor.fetchall()
            
            if documents:
                docs_item = QTreeWidgetItem(parent, ["📄 Документы (Извещения)", f"Всего: {len(documents)}"])
                
                for doc in documents:
                    doc_text = f"№{doc['document_number']}/{doc['year']}"
                    doc_details = f"ID: {doc['id']}, Создан: {doc['created_at']}"
                    if doc['updated_at'] != doc['created_at']:
                        doc_details += f", Обновлён: {doc['updated_at']}"
                    
                    doc_item = QTreeWidgetItem(docs_item, [doc_text, doc_details])
    
    def on_item_expanded(self, item):
        """Обработчик раскрытия элемента"""
        pass
    
    def on_item_selected(self):
        """Обработчик выбора элемента"""
        current_item = self.tree.currentItem()
        if not current_item:
            self.details_text.clear()
            return
        
        # Получаем данные из UserRole, если есть
        entry = current_item.data(0, Qt.UserRole)
        if entry:
            # Это запись каталога
            details = f"""Запись каталога:
            
ID: {entry.id}
Деталь: {entry.part}
Цех: {entry.workshop}
Роль: {entry.role}
Наименование: {entry.before_name}
Единица измерения: {entry.unit}
Норма: {entry.norm}
Примечание: {entry.comment or '(нет)'}
Входит в набор: {'Да' if entry.is_part_of_set else 'Нет'}
ID набора: {entry.replacement_set_id or '(нет)'}
"""
            self.details_text.setPlainText(details)
        else:
            # Показываем информацию об элементе
            text = current_item.text(0)
            info = current_item.text(1)
            details = f"Элемент: {text}\n\n{info}" if info else f"Элемент: {text}"
            self.details_text.setPlainText(details)


def main():
    """Главная функция"""
    app = QApplication(sys.argv)
    
    window = DataTreeViewer()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()



