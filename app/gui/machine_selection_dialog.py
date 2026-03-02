"""
Диалог выбора машин перед генерацией документа.
"""
from typing import List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)


class MachineSelectionDialog(QDialog):
    """Компактный диалог обязательного выбора машин."""

    def __init__(self, products: List[str], parent=None):
        super().__init__(parent)
        self._selected_products: List[str] = []
        self.setWindowTitle("Выбор машин")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)

        header = QLabel("Выберите хотя бы одну машину для генерации документа:")
        header.setWordWrap(True)
        layout.addWidget(header)

        self.products_list = QListWidget()
        self.products_list.setMaximumHeight(220)
        for product_name in products:
            item = QListWidgetItem(product_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            item.setData(Qt.UserRole, product_name)
            self.products_list.addItem(item)
        layout.addWidget(self.products_list)

        buttons = QHBoxLayout()
        buttons.addStretch()

        cancel_button = QPushButton("Отмена")
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(cancel_button)

        confirm_button = QPushButton("Подтвердить")
        confirm_button.clicked.connect(self._confirm_selection)
        confirm_button.setDefault(True)
        buttons.addWidget(confirm_button)

        layout.addLayout(buttons)

    def _confirm_selection(self) -> None:
        selected_products = []
        for i in range(self.products_list.count()):
            item = self.products_list.item(i)
            if item and item.checkState() == Qt.Checked:
                product_name = item.data(Qt.UserRole)
                if product_name:
                    selected_products.append(product_name)

        if not selected_products:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одну машину")
            return

        self._selected_products = selected_products
        self.accept()

    def get_selected_products(self) -> List[str]:
        """Вернуть подтверждённый список выбранных машин."""
        return list(self._selected_products)
