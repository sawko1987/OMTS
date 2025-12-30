"""
Диалог массового управления привязками деталей к изделиям
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QListWidget, QListWidgetItem, QLabel, QMessageBox,
    QLineEdit, QCheckBox, QGroupBox, QSplitter
)
from PySide6.QtCore import Qt
from typing import List

from app.product_store import ProductStore
from app.catalog_loader import CatalogLoader


class ProductPartsBindingDialog(QDialog):
    """Диалог для массового управления привязками деталей к изделиям"""
    
    def __init__(self, product_store: ProductStore, catalog_loader: CatalogLoader, parent=None):
        super().__init__(parent)
        self.product_store = product_store
        self.catalog_loader = catalog_loader
        self.setWindowTitle("Привязки деталей к изделиям")
        self.setMinimumSize(900, 700)
        self.init_ui()
        self.load_data()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Создаём сплиттер для разделения на две панели
        splitter = QSplitter(Qt.Horizontal)
        
        # Левая панель: Изделия (машины)
        products_group = QGroupBox("Изделия (машины)")
        products_layout = QVBoxLayout()
        
        # Кнопка "Выбрать все"
        select_all_btn = QPushButton("Выбрать все")
        select_all_btn.clicked.connect(self.select_all_products)
        products_layout.addWidget(select_all_btn)
        
        # Кнопка "Снять выбор"
        deselect_all_btn = QPushButton("Снять выбор")
        deselect_all_btn.clicked.connect(self.deselect_all_products)
        products_layout.addWidget(deselect_all_btn)
        
        # Список изделий с чекбоксами
        self.products_list = QListWidget()
        self.products_list.setSelectionMode(QListWidget.NoSelection)
        products_layout.addWidget(self.products_list)
        
        products_group.setLayout(products_layout)
        splitter.addWidget(products_group)
        
        # Правая панель: Детали
        parts_group = QGroupBox("Детали")
        parts_layout = QVBoxLayout()
        
        # Поиск деталей
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Поиск:"))
        self.parts_search = QLineEdit()
        self.parts_search.setPlaceholderText("Введите код детали для поиска...")
        self.parts_search.textChanged.connect(self.filter_parts)
        search_layout.addWidget(self.parts_search)
        parts_layout.addLayout(search_layout)
        
        # Список деталей с чекбоксами
        self.parts_list = QListWidget()
        self.parts_list.setSelectionMode(QListWidget.ExtendedSelection)
        parts_layout.addWidget(self.parts_list)
        
        parts_group.setLayout(parts_layout)
        splitter.addWidget(parts_group)
        
        # Устанавливаем пропорции сплиттера (50/50)
        splitter.setSizes([450, 450])
        layout.addWidget(splitter)
        
        # Кнопки действий
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.btn_link = QPushButton("Привязать выбранные")
        self.btn_link.clicked.connect(self.link_selected)
        buttons_layout.addWidget(self.btn_link)
        
        self.btn_unlink = QPushButton("Отвязать выбранные")
        self.btn_unlink.clicked.connect(self.unlink_selected)
        buttons_layout.addWidget(self.btn_unlink)
        
        buttons_layout.addStretch()
        
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)
        buttons_layout.addWidget(btn_close)
        
        layout.addLayout(buttons_layout)
        
        # Информационная метка
        self.info_label = QLabel("Выберите изделия и детали, затем нажмите 'Привязать' или 'Отвязать'")
        self.info_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(self.info_label)
    
    def load_data(self):
        """Загрузить данные из БД"""
        # Загружаем изделия
        products = self.product_store.get_all_products()
        self.products_list.clear()
        self.product_items = {}  # {product_id: QListWidgetItem}
        
        for product_id, product_name in products:
            item = QListWidgetItem(product_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            item.setData(Qt.UserRole, product_id)
            self.products_list.addItem(item)
            self.product_items[product_id] = item
        
        # Загружаем детали
        parts = self.catalog_loader.get_all_parts()
        self.all_parts = parts
        self.update_parts_list(parts)
    
    def update_parts_list(self, parts: List[str]):
        """Обновить список деталей"""
        self.parts_list.clear()
        for part_code in sorted(parts):
            item = QListWidgetItem(part_code)
            self.parts_list.addItem(item)
    
    def filter_parts(self, search_text: str):
        """Фильтровать детали по поисковому запросу"""
        search_text = search_text.strip().upper()
        
        if not search_text:
            self.update_parts_list(self.all_parts)
            return
        
        # Фильтруем детали (регистронезависимый поиск)
        filtered = [
            part for part in self.all_parts
            if search_text in part.upper()
        ]
        
        self.update_parts_list(filtered)
    
    def select_all_products(self):
        """Выбрать все изделия"""
        for i in range(self.products_list.count()):
            item = self.products_list.item(i)
            if item:
                item.setCheckState(Qt.Checked)
    
    def deselect_all_products(self):
        """Снять выбор со всех изделий"""
        for i in range(self.products_list.count()):
            item = self.products_list.item(i)
            if item:
                item.setCheckState(Qt.Unchecked)
    
    def get_selected_products(self) -> List[int]:
        """Получить список ID выбранных изделий"""
        selected = []
        for i in range(self.products_list.count()):
            item = self.products_list.item(i)
            if item and item.checkState() == Qt.Checked:
                product_id = item.data(Qt.UserRole)
                if product_id:
                    selected.append(product_id)
        return selected
    
    def get_selected_parts(self) -> List[str]:
        """Получить список выбранных деталей"""
        selected = []
        for item in self.parts_list.selectedItems():
            part_code = item.text()
            if part_code:
                selected.append(part_code)
        return selected
    
    def link_selected(self):
        """Привязать выбранные детали к выбранным изделиям"""
        product_ids = self.get_selected_products()
        part_codes = self.get_selected_parts()
        
        if not product_ids:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одно изделие")
            return
        
        if not part_codes:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одну деталь")
            return
        
        # Подтверждение
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            f"Привязать {len(part_codes)} деталей к {len(product_ids)} изделиям?\n\n"
            f"Это создаст {len(part_codes) * len(product_ids)} связей.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Выполняем привязку
        created_count = self.product_store.bulk_link_parts_to_products(product_ids, part_codes)
        
        QMessageBox.information(
            self,
            "Успех",
            f"Создано связей: {created_count}\n"
            f"(Некоторые связи могли уже существовать)"
        )
        
        # Обновляем информацию
        self.update_info_label()
    
    def unlink_selected(self):
        """Отвязать выбранные детали от выбранных изделий"""
        product_ids = self.get_selected_products()
        part_codes = self.get_selected_parts()
        
        if not product_ids:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одно изделие")
            return
        
        if not part_codes:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одну деталь")
            return
        
        # Подтверждение
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            f"Отвязать {len(part_codes)} деталей от {len(product_ids)} изделий?\n\n"
            f"Это удалит до {len(part_codes) * len(product_ids)} связей.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Выполняем отвязку
        deleted_count = self.product_store.bulk_unlink_parts_from_products(product_ids, part_codes)
        
        QMessageBox.information(
            self,
            "Успех",
            f"Удалено связей: {deleted_count}"
        )
        
        # Обновляем информацию
        self.update_info_label()
    
    def update_info_label(self):
        """Обновить информационную метку"""
        selected_products = self.get_selected_products()
        selected_parts = self.get_selected_parts()
        
        if selected_products and selected_parts:
            self.info_label.setText(
                f"Выбрано: {len(selected_products)} изделий, {len(selected_parts)} деталей. "
                f"Можно создать/удалить до {len(selected_products) * len(selected_parts)} связей."
            )
        else:
            self.info_label.setText("Выберите изделия и детали, затем нажмите 'Привязать' или 'Отвязать'")

