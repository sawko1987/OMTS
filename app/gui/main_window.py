"""
Главное окно приложения
"""
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QStackedWidget,
    QPushButton, QLabel, QMessageBox, QFileDialog, QMenuBar, QMenu,
    QFrame,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence, QShortcut
from pathlib import Path
from datetime import date, datetime
from typing import Optional
import re
import logging

from app.models import DocumentData
from app.gui.document_info_widget import DocumentInfoWidget
from app.gui.changes_table_widget import ChangesTableWidget
from app.gui.machine_selection_dialog import MachineSelectionDialog
from app.gui.replacement_sets_editor_widget import ReplacementSetsEditorWidget
from app.gui.settings_dialog import SettingsDialog
from app.gui.product_parts_binding_dialog import ProductPartsBindingDialog
from app.gui.database_restore_dialog import DatabaseRestoreDialog
from app.gui.status_bar import StatusBarWidget
from app.gui.wizard_sidebar import WizardSidebar, StepStatus
from app.gui.preview_widget import PreviewWidget
from app.gui.banned_replacements_dialog import BannedReplacementsDialog
from app.gui.audit_log_dialog import AuditLogDialog
from app.draft_manager import DraftManager
from app.audit_log import AuditLog
from app.config import PROJECT_ROOT, DATABASE_PATH, MONTHS
from app.database import DatabaseManager
from app.product_store import ProductStore
from app.catalog_loader import CatalogLoader
from app.history_store import HistoryStore
from app.banned_replacements_store import BannedReplacementsStore  # [BAN_FEATURE]
from app.migrate_to_sqlite import migrate_all
from app.settings_manager import SettingsManager
from app.document_store import DocumentStore
from app.parsing_importer import import_parsing_file

class MainWindow(QMainWindow):
    """Главное окно приложения"""
    
    def __init__(self):
        super().__init__()
        self.document_data = DocumentData()
        
        # Счётчик дополнительных страниц (0 = первая страница, 1 = 1+, 2 = 2+ и т.д.)
        self.current_additional_page = None  # None = данные идут на первую страницу
        # Следующий номер доп. страницы, который будет назначен при нажатии кнопки
        self.next_additional_page = 1
        
        # Сохраняем последнюю использованную дату для использования при создании нового документа
        self._last_used_date = None
        self._last_used_validity_period = None
        self._last_used_year = None  # Год для нумерации
        
        # Инициализация БД (до создания UI)
        self.init_database()
        
        # Создаём менеджеры данных
        self.db_manager = DatabaseManager()
        self.product_store = ProductStore(self.db_manager)
        self.catalog_loader = CatalogLoader(self.db_manager)
        self.history_store = HistoryStore(self.db_manager)
        self.banned_store = BannedReplacementsStore(self.db_manager)  # [BAN_FEATURE]
        self.settings_manager = SettingsManager()
        self.document_store = DocumentStore(self.db_manager, self.catalog_loader)

        # Draft manager (autosave) — инициализируем ДО init_ui,
        # чтобы можно было предложить восстановление при старте
        self.draft_manager = DraftManager()
        self._autosave_timer = None  # создаётся в init_ui

        # Audit log — журнал действий пользователя (в памяти)
        self.audit_log = AuditLog()

        self.init_ui()
        self.load_catalog()

    def _apply_last_used_date(self):
        """Установить сохранённые реквизиты в новый документ текущего сеанса."""
        if self._last_used_date:
            self.document_data.implementation_date = self._last_used_date
        if self._last_used_validity_period:
            self.document_data.validity_period = self._last_used_validity_period
        if self._last_used_year:
            self.document_data.document_year = self._last_used_year

    def _on_implementation_date_changed(self, qdate):
        """Сохранять последнюю выбранную дату до перезапуска приложения."""
        if qdate and qdate.isValid():
            self._last_used_date = qdate.toPython()

    # ----- Навигация по визарду -----

    def _on_wizard_step_selected(self, index: int) -> None:
        """Переключить активную страницу визарда."""
        if not hasattr(self, "pages_stack"):
            return
        if 0 <= index < self.pages_stack.count():
            cur = self.pages_stack.currentIndex()
            if cur == index:
                return
            if cur == 2 and hasattr(self, "replacement_sets_widget") and self.replacement_sets_widget:
                if self.replacement_sets_widget.has_unsaved_changes():
                    if not self.replacement_sets_widget.prompt_unsaved_changes():
                        if hasattr(self, "sidebar"):
                            self.sidebar.set_current(cur)
                        return
            self.pages_stack.setCurrentIndex(index)
            self.sidebar.set_current(index)
            self._update_nav_buttons()
            self._update_wizard_step_status()
            # Обновляем preview при заходе на шаг 4
            if index == 3 and hasattr(self, "preview_widget"):
                self.preview_widget.refresh()

    def _wizard_prev(self) -> None:
        """Перейти на предыдущий шаг."""
        cur = self.pages_stack.currentIndex()
        if cur > 0:
            self._on_wizard_step_selected(cur - 1)

    def _wizard_next(self) -> None:
        """Перейти на следующий шаг."""
        cur = self.pages_stack.currentIndex()
        last = self.pages_stack.count() - 1
        if cur < last:
            self._on_wizard_step_selected(cur + 1)

    def _update_nav_buttons(self) -> None:
        """Включить/выключить кнопки Назад/Далее в зависимости от позиции."""
        if not hasattr(self, "pages_stack"):
            return
        cur = self.pages_stack.currentIndex()
        last = self.pages_stack.count() - 1
        self.btn_prev_step.setEnabled(cur > 0)
        self.btn_next_step.setEnabled(cur < last)

    def _update_wizard_step_status(self) -> None:
        """Обновить статус шагов визарда (готово/предупреждение) на основе данных документа."""
        if not hasattr(self, "sidebar"):
            return
        data = self.document_data

        # Шаг 0: Шапка — готово, если заполнена дата
        if data.implementation_date:
            self.sidebar.set_step_status(0, StepStatus.DONE)
        else:
            self.sidebar.set_step_status(0, StepStatus.WARNING)

        # Шаг 1: Материалы — готово, если есть хотя бы одна замена
        has_replacement = any(
            m.is_changed and (m.after_name or "").strip()
            for pc in data.part_changes
            for m in pc.materials
        )
        if has_replacement:
            self.sidebar.set_step_status(1, StepStatus.DONE)
        elif data.part_changes:
            self.sidebar.set_step_status(1, StepStatus.WARNING)
        else:
            self.sidebar.set_step_status(1, StepStatus.PENDING)

        # Шаг 2: Наборы — всегда optional, помечаем done если были изменения в этом сеансе
        # (пока просто pending — здесь сложно отследить "был ли пользователь здесь")
        # Шаг 3: Предпросмотр — готово к генерации если всё заполнено
        can_generate = bool(data.implementation_date and data.part_changes and data.products)
        active_bans = 0
        if self.banned_store:
            for pc in data.part_changes:
                for m in pc.materials:
                    if m.is_changed and (m.after_name or "").strip():
                        if self.banned_store.is_banned(m.catalog_entry, (m.after_name or "").strip()):
                            active_bans += 1
        if active_bans > 0:
            self.sidebar.set_step_status(3, StepStatus.ERROR)
            self.sidebar.set_hint(f"⚠ {active_bans} активных запрет(ов) — нужно снять или убрать замену")
        elif can_generate:
            self.sidebar.set_step_status(3, StepStatus.DONE)
            self.sidebar.set_hint("✓ Готово к генерации")
        else:
            self.sidebar.set_step_status(3, StepStatus.WARNING)
            if not data.implementation_date:
                self.sidebar.set_hint("⚠ Укажите дату внедрения (Шаг 1)")
            elif not data.part_changes:
                self.sidebar.set_hint("⚠ Добавьте хотя бы одну деталь (Шаг 2)")
            elif not data.products:
                self.sidebar.set_hint("⚠ Выберите изделия (Шаг 2)")

    def _on_validity_period_changed(self, text):
        """Сохранять последний введённый срок действия до перезапуска приложения."""
        value = text.strip()
        if value:
            self._last_used_validity_period = value

    def _on_year_changed(self, year):
        """Сохранять последний выбранный год и записывать в настройки."""
        self._last_used_year = year
        self.settings_manager.set_default_year(year)

    def _setup_shortcuts(self):
        """Назначить горячие клавиши главному окну."""
        # Ctrl+N — новый документ
        QShortcut(QKeySequence("Ctrl+N"), self, activated=self.new_document)
        # Ctrl+O — открыть документ
        QShortcut(QKeySequence("Ctrl+O"), self, activated=self.open_document)
        # Ctrl+S — сохранить черновик (пока no-op, будет в Фазе 3)
        QShortcut(QKeySequence("Ctrl+S"), self, activated=self._shortcut_save_draft)
        # Ctrl+G — сгенерировать документ
        QShortcut(QKeySequence("Ctrl+G"), self, activated=self.generate_document)
        # Ctrl+Shift+Del (рус. "Очистить") — очистить данные
        QShortcut(QKeySequence("Ctrl+Shift+Delete"), self, activated=self.clear_data)
        # F1 — справка по горячим клавишам
        QShortcut(QKeySequence("F1"), self, activated=self._show_help)
        # Esc — закрыть активный диалог (если есть) или сбросить доп. страницу
        QShortcut(QKeySequence("Escape"), self, activated=self._shortcut_escape)

    def _shortcut_save_draft(self):
        """Ctrl+S — сохранить черновик вручную."""
        self.save_draft(source="manual")

    def save_draft(self, source: str = "manual") -> None:
        """Сохранить текущий документ как черновик.

        Args:
            source: "auto" для автосохранения, "manual" для Ctrl+S
        """
        try:
            self.draft_manager.save(self.document_data, source=source)
            self._document_dirty = False
            self._last_saved_at = datetime.now().isoformat(timespec="seconds")
            self._refresh_draft_status_indicator()
            if source == "manual":
                from app.gui.toast import show_toast, ToastLevel
                show_toast(self, "Черновик сохранён", ToastLevel.SUCCESS)
                self.audit_log.info("save_draft", "Черновик сохранён вручную (Ctrl+S)")
            logger.info(f"Черновик сохранён (source={source})")
        except Exception as e:
            logger.error(f"Ошибка сохранения черновика: {e}")
            from app.gui.toast import show_toast, ToastLevel
            show_toast(self, f"Не удалось сохранить черновик: {e}", ToastLevel.ERROR, timeout_ms=5000)
            self.audit_log.error("save_draft", f"Ошибка: {e}")

    def _autosave_tick(self) -> None:
        """Тик таймера автосохранения: сохраняем только если есть изменения."""
        if not self._document_dirty:
            return
        if not self.document_data.part_changes and not self.document_data.products:
            # Пустой документ — не сохраняем, чтобы не замусоривать черновик
            return
        logger.debug("Автосохранение: есть изменения, сохраняю черновик")
        self.save_draft(source="auto")

    def _on_document_dirty(self, *_args) -> None:
        """Пометить документ как содержащий несохранённые изменения."""
        self._document_dirty = True
        self._refresh_draft_status_indicator()

    def _refresh_draft_status_indicator(self) -> None:
        """Обновить индикатор черновика в статус-баре."""
        if not hasattr(self, "status_bar_widget"):
            return
        if self._last_saved_at is None:
            self.status_bar_widget.set_draft_status("не сохранён", dirty=self._document_dirty)
        else:
            try:
                saved_dt = datetime.fromisoformat(self._last_saved_at)
                time_str = saved_dt.strftime("%H:%M:%S")
            except ValueError:
                time_str = "??:??:??"
            self.status_bar_widget.set_draft_status(
                f"сохранён {time_str}",
                dirty=self._document_dirty,
            )

    def _check_draft_on_startup(self) -> None:
        """При старте: если есть свежий черновик, предложить восстановить."""
        info = self.draft_manager.get_draft_info()
        if info is None:
            return
        # Не предлагаем восстановление, если в черновике пусто
        if info.part_count == 0 and not self.document_data.products:
            self.draft_manager.discard()
            return

        age = info.age_human()
        reply = QMessageBox.question(
            self,
            "Восстановить черновик?",
            f"Найден несохранённый черновик ({info.part_count} деталей, "
            f"сохранён {age}).\n\n"
            f"Восстановить его вместо пустого документа?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Yes:
            loaded = self.draft_manager.load(catalog_loader=self.catalog_loader)
            if loaded:
                self._replace_document_data(loaded)
                from app.gui.toast import show_toast, ToastLevel
                show_toast(self, "Черновик восстановлен", ToastLevel.SUCCESS, timeout_ms=4000)
                self.audit_log.info(
                    "restore_draft",
                    f"Восстановлен черновик ({info.part_count} деталей, {info.age_human()})",
                )
                logger.info("Черновик восстановлен при старте")
        else:
            # Пользователь отказался — удаляем черновик
            self.draft_manager.discard()

    def _replace_document_data(self, new_data: DocumentData) -> None:
        """Заменить текущий документ на новый (используется при восстановлении черновика)."""
        self.document_data = new_data
        self._last_used_date = new_data.implementation_date
        self._last_used_validity_period = new_data.validity_period
        self._last_used_year = new_data.document_year
        # Состояние доп. страниц
        max_page = 0
        for pc in new_data.part_changes:
            if pc.additional_page_number is not None:
                max_page = max(max_page, pc.additional_page_number)
        self.current_additional_page = None
        self.next_additional_page = max_page + 1 if max_page > 0 else 1
        # Применяем ко всем виджетам
        self.doc_info_widget.document_data = self.document_data
        self.changes_widget.document_data = self.document_data
        self.replacement_sets_widget.document_data = self.document_data
        self.doc_info_widget.refresh()
        self.changes_widget.refresh()
        self.changes_widget.update_product_filter()
        self._document_dirty = False
        self._last_saved_at = datetime.now().isoformat(timespec="seconds")
        self._refresh_draft_status_indicator()

    # ===== Импорт/экспорт черновика =====

    def export_draft(self) -> None:
        """Экспорт текущего документа в JSON-файл."""
        if not self.document_data.part_changes and not self.document_data.products:
            QMessageBox.information(
                self, "Экспорт",
                "Нечего экспортировать — документ пуст."
            )
            return
        default_name = "izveshchenie_draft.json"
        if self.document_data.get_display_number():
            default_name = f"izveshchenie_{self.document_data.get_display_number()}.json"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Экспорт черновика",
            str(PROJECT_ROOT / default_name),
            "JSON файлы (*.json);;Все файлы (*)",
        )
        if not file_path:
            return
        try:
            self.draft_manager.export_to(self.document_data, Path(file_path))
            from app.gui.toast import show_toast, ToastLevel
            show_toast(self, f"Экспортировано: {Path(file_path).name}", ToastLevel.SUCCESS)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка экспорта", f"Не удалось экспортировать:\n{e}")

    def import_draft(self) -> None:
        """Импорт документа из JSON-файла."""
        if hasattr(self, "replacement_sets_widget") and self.replacement_sets_widget:
            if self.replacement_sets_widget.has_unsaved_changes():
                if not self.replacement_sets_widget.prompt_unsaved_changes():
                    return
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Импорт черновика",
            str(PROJECT_ROOT),
            "JSON файлы (*.json);;Все файлы (*)",
        )
        if not file_path:
            return
        try:
            loaded = self.draft_manager.import_from(Path(file_path), catalog_loader=self.catalog_loader)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка импорта", f"Не удалось импортировать:\n{e}")
            return
        if loaded is None:
            QMessageBox.warning(self, "Ошибка", "Файл повреждён или имеет неверный формат.")
            return
        # Подтверждение замены текущего документа
        if self.document_data.part_changes or self.document_data.products:
            reply = QMessageBox.question(
                self, "Импорт черновика",
                "Импорт заменит текущие данные документа. Продолжить?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
        self._replace_document_data(loaded)
        from app.gui.toast import show_toast, ToastLevel
        show_toast(self, f"Импортировано: {Path(file_path).name}", ToastLevel.SUCCESS)
        logger.info(f"Документ импортирован из {file_path}")

    def _shortcut_escape(self):
        """Escape: сбросить активную доп. страницу (вернуться к основной)."""
        if self.current_additional_page is not None:
            self.current_additional_page = None
            self.status_bar_widget.set_ready_text("Возврат к основной странице")

    def _show_help(self):
        """Показать окно со списком горячих клавиш."""
        QMessageBox.information(
            self,
            "Горячие клавиши",
            "Глобальные:\n"
            "  Ctrl+N — Новый документ\n"
            "  Ctrl+O — Открыть документ\n"
            "  Ctrl+S — Сохранить черновик (в разработке)\n"
            "  Ctrl+G — Сгенерировать документ\n"
            "  Ctrl+Shift+Del — Очистить данные\n"
            "  Esc — Сбросить активную доп. страницу\n"
            "  F1 — Эта справка\n\n"
            "Навигация по визарду:\n"
            "  Кнопки «← Назад» / «Далее →» внизу окна\n"
            "  Шаги в сайдбаре слева\n\n"
            "В таблицах:\n"
            "  Delete — удалить выбранные строки (в наборах)\n"
            "  Правый клик — групповые операции (в материалах)",
        )

    def refresh_status_bar(self):
        """Пересчитать счётчики статус-бара. Вызывать после изменения данных."""
        if hasattr(self, "status_bar_widget"):
            self.status_bar_widget.refresh()
    
    def init_database(self):
        """Инициализировать базу данных"""
        db_manager = DatabaseManager()
        
        # Проверяем, существует ли БД
        if not DATABASE_PATH.exists():
            # Первый запуск - создаём схему и мигрируем данные
            db_manager.initialize()
            migrate_all()
            # Показываем сообщение после создания UI
            # (будет показано в init_ui, если нужно)
        else:
            # БД существует, просто инициализируем схему (создаст таблицы, если их нет)
            db_manager.initialize()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        self.setWindowTitle("Извещение на замену материалов")
        self.setMinimumSize(1200, 800)
        
        # Меню
        menubar = self.menuBar()
        
        # Меню "Документ"
        document_menu = menubar.addMenu("Документ")
        new_doc_action = document_menu.addAction("Новый документ")
        new_doc_action.setShortcut(QKeySequence("Ctrl+N"))
        new_doc_action.triggered.connect(self.new_document)
        open_doc_action = document_menu.addAction("Открыть документ...")
        open_doc_action.setShortcut(QKeySequence("Ctrl+O"))
        open_doc_action.triggered.connect(self.open_document)
        document_menu.addSeparator()
        save_draft_action = document_menu.addAction("Сохранить черновик")
        save_draft_action.setShortcut(QKeySequence("Ctrl+S"))
        save_draft_action.triggered.connect(self.save_draft)
        document_menu.addSeparator()
        import_draft_action = document_menu.addAction("Импорт черновика из JSON...")
        import_draft_action.triggered.connect(self.import_draft)
        export_draft_action = document_menu.addAction("Экспорт черновика в JSON...")
        export_draft_action.triggered.connect(self.export_draft)
        
        # Меню "Данные"
        data_menu = menubar.addMenu("Данные")
        import_parsing_action = data_menu.addAction("Импорт пакетов из Парсинг.xlsx...")
        import_parsing_action.triggered.connect(self.import_parsing_file)
        binding_action = data_menu.addAction("Привязки деталей к изделиям...")
        binding_action.triggered.connect(self.show_product_parts_binding)
        banned_action = data_menu.addAction("⛔ Запреты замен материалов…")
        banned_action.triggered.connect(self.show_banned_replacements)
        data_menu.addSeparator()
        restore_action = data_menu.addAction("Восстановить базу данных...")
        restore_action.triggered.connect(self.show_database_restore)
        
        settings_menu = menubar.addMenu("Настройки")
        settings_action = settings_menu.addAction("Настройки...")
        settings_action.triggered.connect(self.show_settings)

        # Меню "Журнал"
        log_menu = menubar.addMenu("Журнал")
        log_action = log_menu.addAction("📋 Журнал действий…")
        log_action.triggered.connect(self.show_audit_log)
        
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Основной layout
        main_layout = QVBoxLayout(central_widget)

        # === Wizard: сайдбар + QStackedWidget ===
        wizard_layout = QHBoxLayout()
        wizard_layout.setSpacing(0)

        self.sidebar = WizardSidebar([
            "Шапка документа",
            "Материалы",
            "Наборы",
            "Предпросмотр",
        ])
        self.sidebar.step_selected.connect(self._on_wizard_step_selected)
        wizard_layout.addWidget(self.sidebar)

        # Контейнер страниц
        self.pages_stack = QStackedWidget()
        self.pages_stack.setFrameShape(QFrame.NoFrame)
        wizard_layout.addWidget(self.pages_stack, 1)

        main_layout.addLayout(wizard_layout, 1)

        # === Страница 1: Реквизиты документа ===
        self.doc_info_widget = DocumentInfoWidget(self.document_data, self.product_store, self.db_manager)
        self.doc_info_widget.date_edit.dateChanged.connect(self._on_implementation_date_changed)
        self.doc_info_widget.validity_combo.currentTextChanged.connect(self._on_validity_period_changed)
        self.doc_info_widget.year_spin.valueChanged.connect(self._on_year_changed)
        self.doc_info_widget.reason_edit.textChanged.connect(self._on_document_dirty)
        self.pages_stack.addWidget(self.doc_info_widget)

        # === Страница 2: Изменения материалов ===
        self.changes_widget = ChangesTableWidget(
            self.document_data,
            self.catalog_loader,
            self.history_store,
            self.product_store,
            banned_store=self.banned_store  # [BAN_FEATURE]
        )
        # Передаём функцию для получения текущего номера доп. страницы
        self.changes_widget.get_current_additional_page = lambda: self.current_additional_page
        # Колбэк из breadcrumb: установить активную доп. страницу (или None для основной)
        self.changes_widget.set_current_additional_page = self._set_current_additional_page
        # Колбэк из breadcrumb: пользователь нажал "+ Добавить страницу"
        self.changes_widget.request_new_additional_page = self.add_additional_page_data
        self.pages_stack.addWidget(self.changes_widget)

        # === Страница 3: Наборы ===
        self.replacement_sets_widget = ReplacementSetsEditorWidget(
            self.catalog_loader,
            document_data=self.document_data,
            document_store=self.document_store,
            history_store=self.history_store,
            product_store=self.product_store,
            banned_store=self.banned_store,  # [BAN_FEATURE]
            get_current_additional_page=lambda: self.current_additional_page,
            changes_widget=self.changes_widget
        )
        self.changes_widget.replacement_sets_widget = self.replacement_sets_widget
        self.pages_stack.addWidget(self.replacement_sets_widget)

        # === Страница 4: Предпросмотр ===
        self.preview_widget = PreviewWidget(
            self.document_data, banned_store=self.banned_store
        )
        self.preview_widget.goto_materials.connect(lambda: self._on_wizard_step_selected(1))
        self.preview_widget.request_generate.connect(self.generate_document)
        self.pages_stack.addWidget(self.preview_widget)

        # Синхронизация выбора детали: страница «Материалы» -> «Наборы»
        self.changes_widget.part_selected.connect(self.replacement_sets_widget.set_part)
        # Обновление статус-бара и предпросмотра при изменении данных документа
        self.changes_widget.document_data_changed.connect(self.refresh_status_bar)
        self.replacement_sets_widget.document_data_changed.connect(self.refresh_status_bar)
        self.changes_widget.document_data_changed.connect(self.preview_widget.refresh)
        self.replacement_sets_widget.document_data_changed.connect(self.preview_widget.refresh)
        self.changes_widget.document_data_changed.connect(self._update_wizard_step_status)
        self.replacement_sets_widget.document_data_changed.connect(self._update_wizard_step_status)
        # Автосохранение — помечаем документ как «грязный» при любом изменении
        self.changes_widget.document_data_changed.connect(self._on_document_dirty)
        self.replacement_sets_widget.document_data_changed.connect(self._on_document_dirty)
        
        # Кнопки внизу — навигация по визарду + действия
        nav_layout = QHBoxLayout()

        self.btn_prev_step = QPushButton("← Назад")
        self.btn_prev_step.clicked.connect(self._wizard_prev)
        nav_layout.addWidget(self.btn_prev_step)

        self.btn_next_step = QPushButton("Далее →")
        self.btn_next_step.clicked.connect(self._wizard_next)
        nav_layout.addWidget(self.btn_next_step)

        nav_layout.addSpacing(20)

        self.btn_generate = QPushButton("📄 Сгенерировать документ")
        self.btn_generate.setProperty("primary", True)
        self.btn_generate.setStyleSheet(
            "QPushButton {"
            "  background-color: #2a6da8; color: #ffffff; "
            "  border: 1px solid #1f568a; border-radius: 4px;"
            "  padding: 6px 18px; font-weight: 600;"
            "}"
            "QPushButton:hover { background-color: #3480c2; }"
            "QPushButton:pressed { background-color: #1f568a; }"
        )
        self.btn_generate.clicked.connect(self.generate_document)
        nav_layout.addWidget(self.btn_generate)

        self.btn_add_additional_page = QPushButton("+ Доп. страница")
        self.btn_add_additional_page.setToolTip(
            "Добавить данные для следующей дополнительной страницы документа "
            "(Esc — вернуться на основную)"
        )
        self.btn_add_additional_page.clicked.connect(self.add_additional_page_data)
        nav_layout.addWidget(self.btn_add_additional_page)

        nav_layout.addStretch()

        self.btn_clear = QPushButton("🧹 Очистить")
        self.btn_clear.clicked.connect(self.clear_data)
        nav_layout.addWidget(self.btn_clear)

        main_layout.addLayout(nav_layout)

        # Статус-бар внизу окна (счётчики документа, состояние БД)
        self.status_bar_widget = StatusBarWidget(
            self.document_data, banned_store=self.banned_store
        )
        main_layout.addWidget(self.status_bar_widget)

        # Горячие клавиши
        self._setup_shortcuts()

        # Автосохранение черновика каждые 30 сек (если есть изменения)
        from PySide6.QtCore import QTimer
        self._autosave_timer = QTimer(self)
        self._autosave_timer.setInterval(30_000)
        self._autosave_timer.timeout.connect(self._autosave_tick)
        self._autosave_timer.start()
        self._document_dirty = False  # есть несохранённые изменения
        self._last_saved_at: Optional[str] = None  # ISO-строка последнего сохранения

        # Проверяем наличие черновика при старте
        self._check_draft_on_startup()
    
    def load_catalog(self):
        """Загрузить каталог и обновить виджеты"""
        try:
            if hasattr(self, "replacement_sets_widget") and self.replacement_sets_widget:
                if self.replacement_sets_widget.has_unsaved_changes():
                    if not self.replacement_sets_widget.prompt_unsaved_changes():
                        return
            self.catalog_loader.load()
            self.changes_widget.load_catalog()
            self.replacement_sets_widget.load_catalog()

            # Подхватываем текущую деталь из комбобокса (если есть)
            current_part = self.changes_widget.part_combo.currentText().strip()
            if current_part:
                self.replacement_sets_widget.set_part(current_part)
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить справочник:\n{e}")
    
    def show_settings(self):
        """Показать диалог настроек"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            # Настройки сохранены в диалоге
            # Обновляем номер документа, если он был изменен
            self.doc_info_widget.refresh_number()
    
    def import_parsing_file(self):
        """Импортировать данные из Парсинг.xlsx"""
        # Запрашиваем файл
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Парсинг.xlsx",
            str(PROJECT_ROOT),
            "Excel файлы (*.xlsx);;Все файлы (*)"
        )
        
        if not file_path:
            return
        
        file_path_obj = Path(file_path)
        
        # Подтверждение
        reply = QMessageBox.question(
            self,
            "Подтверждение импорта",
            f"Импортировать данные из файла:\n{file_path_obj.name}\n\n"
            "Это создаст наборы материалов для деталей из файла и привяжет их ко всем изделиям.\n"
            "Перед импортом будет создана резервная копия базы данных.\n\n"
            "Продолжить?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Выполняем импорт
        try:
            imported_parts, created_sets, errors = import_parsing_file(
                file_path=file_path_obj,
                catalog_loader=self.catalog_loader,
                product_store=self.product_store,
                db_manager=self.db_manager
            )
            
            # Обновляем каталог в интерфейсе
            self.load_catalog()
            
            # Принудительно обновляем список деталей в виджете изменений
            self.changes_widget.load_parts_list()
            
            # Проверяем, сколько деталей действительно в БД
            all_parts_after = self.catalog_loader.get_all_parts()
            
            # Показываем результат
            message = (
                f"Импорт завершён!\n\n"
                f"Успешно импортировано деталей: {imported_parts}\n"
                f"Создано наборов: {created_sets}\n"
                f"Всего деталей в БД: {len(all_parts_after)}"
            )
            
            if errors:
                message += f"\n\nОшибок при импорте: {len(errors)}"
                for error in errors[:5]:  # Показываем первые 5 ошибок
                    message += f"\n  - {error}"
                if len(errors) > 5:
                    message += f"\n  ... и ещё {len(errors) - 5} ошибок"
            
            QMessageBox.information(self, "Импорт завершён", message)
        
        except Exception as e:
            QMessageBox.critical(
                self,
                "Ошибка импорта",
                f"Произошла ошибка при импорте:\n{e}"
            )
    
    def show_product_parts_binding(self):
        """Показать диалог управления привязками деталей к изделиям"""
        dialog = ProductPartsBindingDialog(
            self.product_store,
            self.catalog_loader,
            self
        )
        dialog.exec()
        
# Обновляем каталог, так как могли измениться привязки
        self.load_catalog()

    def show_banned_replacements(self):
        """Показать диалог управления всеми запретами замены."""
        dialog = BannedReplacementsDialog(self.banned_store, self)
        dialog.exec()
        # После закрытия диалога могли снять запреты — обновляем статус-бар и предпросмотр
        self.refresh_status_bar()
        if hasattr(self, "preview_widget"):
            self.preview_widget.refresh()

    def show_audit_log(self):
        """Показать диалог журнала действий."""
        dialog = AuditLogDialog(self.audit_log, self)
        dialog.exec()

    def show_database_restore(self):
        """Показать диалог восстановления базы данных"""
        reply = QMessageBox.warning(
            self,
            "Восстановление базы данных",
            "Восстановление базы данных из резервной копии приведёт к потере всех изменений, "
            "сделанных после создания резервной копии.\n\n"
            "Текущая БД будет сохранена перед восстановлением.\n\n"
            "После восстановления необходимо перезапустить приложение.\n\n"
            "Продолжить?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        dialog = DatabaseRestoreDialog(self)
        if dialog.exec():
            # Если восстановление прошло успешно, показываем сообщение о перезапуске
            QMessageBox.warning(
                self,
                "Перезапуск необходим",
                "База данных восстановлена.\n\n"
                "⚠ ВАЖНО: Закройте приложение и запустите его снова для применения изменений."
            )
    
    def get_first_material_name(self) -> str:
        """Получить название первого заменяемого материала"""
        for part_change in self.document_data.part_changes:
            for material in part_change.materials:
                if material.is_changed and material.catalog_entry.before_name and material.catalog_entry.before_name.strip():
                    return material.catalog_entry.before_name.strip()
        return "материал"
    
    def get_month_year_folder(self) -> str:
        """Сформировать название папки в формате 'Январь_2024' на основе даты внедрения"""
        if not self.document_data.implementation_date:
            # Если дата не указана, используем текущую дату
            impl_date = date.today()
        else:
            impl_date = self.document_data.implementation_date
        
        month_name = MONTHS.get(impl_date.month, "Неизвестно")
        year = impl_date.year
        
        return f"{month_name}_{year}"
    
    def sanitize_filename(self, filename: str, max_length: int = 200) -> str:
        """Очистить название от недопустимых символов для имени файла"""
        # Удаляем недопустимые символы для Windows
        invalid_chars = r'[<>:"/\\|?*]'
        sanitized = re.sub(invalid_chars, '_', filename)
        
        # Удаляем ведущие и завершающие пробелы и точки
        sanitized = sanitized.strip(' .')
        
        # Ограничиваем длину
        if len(sanitized) > max_length:
            sanitized = sanitized[:max_length]
        
        # Если после очистки строка пустая, возвращаем значение по умолчанию
        if not sanitized:
            sanitized = "материал"
        
        return sanitized
    
    def generate_document(self):
        """Генерация документа"""
        # Обновляем данные из виджетов
        self.doc_info_widget.update_document_data()
        self.changes_widget.update_document_data()
        
        # Сохраняем дату при обновлении данных (на случай, если пользователь изменил дату)
        if self.document_data.implementation_date:
            self._last_used_date = self.document_data.implementation_date
        if self.document_data.validity_period:
            self._last_used_validity_period = self.document_data.validity_period
        if self.document_data.document_year:
            self._last_used_year = self.document_data.document_year
            self.settings_manager.set_default_year(self.document_data.document_year)
        
        # Валидация
        if not self.document_data.implementation_date:
            QMessageBox.warning(self, "Ошибка", "Укажите дату внедрения замены")
            self._on_wizard_step_selected(0)
            return
        
        if not self.document_data.part_changes:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну деталь с изменениями")
            self._on_wizard_step_selected(1)
            return
        
        # Проверяем, что есть хотя бы одно изменение
        has_changes = False
        missing_after = []
        missing_checkbox = []
        
        for part_change in self.document_data.part_changes:
            for material in part_change.materials:
                if material.is_changed:
                    if not material.after_name or not material.after_name.strip():
                        missing_after.append(f"{part_change.part}: {material.catalog_entry.before_name}")
                    else:
                        has_changes = True
        
        # Более детальное сообщение об ошибке
        if not has_changes:
            error_msg = "Не найдено ни одного материала для замены.\n\n"
            if missing_after:
                error_msg += "Следующие материалы отмечены для замены, но не указан материал 'после':\n"
                for item in missing_after[:5]:  # Показываем первые 5
                    error_msg += f"  - {item}\n"
                if len(missing_after) > 5:
                    error_msg += f"  ... и ещё {len(missing_after) - 5}\n"
            else:
                error_msg += "Отметьте чекбокс 'Меняем' для хотя бы одного материала и укажите материал 'после'."
            
            QMessageBox.warning(self, "Ошибка", error_msg)
            self._on_wizard_step_selected(1)
            return

        # [BAN_FEATURE] Проверка на запрещённые замены
        banned_found = []
        for part_change in self.document_data.part_changes:
            for material in part_change.materials:
                if material.is_changed and material.after_name:
                    after_name = material.after_name.strip()
                    if self.banned_store.is_banned(material.catalog_entry, after_name):
                        banned_found.append(
                            f"{part_change.part}: {material.catalog_entry.before_name} → {after_name}"
                        )
        if banned_found:
            error_msg = "Обнаружены запрещённые конструктором замены:\n\n"
            for item in banned_found[:10]:
                error_msg += f"  ⛔ {item}\n"
            if len(banned_found) > 10:
                error_msg += f"  ... и ещё {len(banned_found) - 10}\n"
            error_msg += "\nСнимите запрет во вкладке «Изменения материалов» или уберите замену для указанных материалов."
            QMessageBox.warning(self, "Ошибка", error_msg)
            self._on_wizard_step_selected(1)
            return
        
        if not self.ensure_products_selected_for_generation():
            return
            
        # Получаем папку сохранения из настроек
        selected_dir = self.settings_manager.get_output_directory()
        
        # Если папка не настроена, предлагаем выбрать её
        if not selected_dir:
            output_dir = PROJECT_ROOT / "output"
            selected_dir = QFileDialog.getExistingDirectory(
                self,
                "Выберите папку для сохранения извещений",
                str(output_dir)
            )
            
            if not selected_dir:
                return
            
            # Сохраняем выбранную папку в настройки
            try:
                self.settings_manager.set_output_directory(selected_dir)
            except Exception as e:
                QMessageBox.warning(self, "Предупреждение", 
                                  f"Не удалось сохранить настройки:\n{e}\n\n"
                                  "Папка будет использована только для этого документа.")
        else:
            # Проверяем доступность сохраненного пути
            path_obj = Path(selected_dir)
            if not path_obj.exists():
                # Путь не существует - пытаемся создать его
                try:
                    path_obj.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    # Не удалось создать - предлагаем выбрать другую папку
                    reply = QMessageBox.question(
                        self,
                        "Папка недоступна",
                        f"Сохраненная папка недоступна:\n{selected_dir}\n\n"
                        f"Ошибка: {e}\n\n"
                        "Выбрать другую папку?",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.Yes:
                        output_dir = PROJECT_ROOT / "output"
                        selected_dir = QFileDialog.getExistingDirectory(
                            self,
                            "Выберите папку для сохранения извещений",
                            str(output_dir)
                        )
                        if not selected_dir:
                            return
                        try:
                            self.settings_manager.set_output_directory(selected_dir)
                        except Exception as e2:
                            QMessageBox.warning(self, "Предупреждение", 
                                              f"Не удалось сохранить настройки:\n{e2}")
                    else:
                        return
            elif not path_obj.is_dir():
                # Путь существует, но это не папка
                QMessageBox.warning(self, "Ошибка", 
                                  f"Указанный путь не является папкой:\n{selected_dir}")
                return
        
        # Получаем номер документа заранее (для формирования имени файла)
        # Импортируем здесь, чтобы избежать циклических импортов
        from app.numbering import NumberingManager
        
        numbering = NumberingManager()
        doc_month = self.document_data.document_month or date.today().month
        doc_year = self.document_data.document_year or date.today().year
        
        if not self.document_data.document_number:
            # Получаем следующий номер и устанавливаем его в document_data
            doc_number = numbering.get_next_number(doc_year, doc_month)
            self.document_data.document_number = doc_number
            self.document_data.document_month = doc_month
            self.document_data.document_year = doc_year
        else:
            doc_number = self.document_data.document_number
            doc_month = self.document_data.document_month or date.today().month
            doc_year = self.document_data.document_year or date.today().year
        
        display_number = self.document_data.get_display_number()
        
        # Формируем структуру папок: месяц_год
        month_year_folder = self.get_month_year_folder()
        target_dir = Path(selected_dir) / month_year_folder
        
        # Создаём папку, если её нет
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", 
                               f"Не удалось создать папку для сохранения:\n{target_dir}\n\n"
                               f"Ошибка: {e}")
            return
        
        # Получаем название первого материала
        material_name = self.get_first_material_name()
        sanitized_material = self.sanitize_filename(material_name)
        
        # Формируем имя файла: ММ-ГГ-№_материал.xlsx
        filename = f"{display_number}_{sanitized_material}.xlsx"
        file_path = target_dir / filename
        
        # Импортируем генератор здесь, чтобы избежать циклических импортов
        from app.excel_generator import ExcelGenerator
        
        try:
            # Проверяем, существует ли уже документ с таким номером, годом и месяцем
            document_year = doc_year
            if self.document_store.document_exists(doc_number, document_year, doc_month):
                # Получаем информацию о существующем документе
                existing_docs = self.document_store.get_all_documents(document_year)
                existing_doc = next((d for d in existing_docs if d[0] == doc_number), None)
                
                if existing_doc:
                    existing_file = Path(existing_doc[4]).name if existing_doc[4] else "неизвестно"
                    reply = QMessageBox.warning(
                        self,
                        "Документ уже существует",
                        f"Документ №{display_number} уже существует:\n\n"
                        f"Файл: {existing_file}\n"
                        f"Создан: {existing_doc[3]}\n\n"
                        f"Перезаписать существующий документ?\n\n"
                        f"⚠ ВНИМАНИЕ: Это удалит старый документ без возможности восстановления!",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No
                    )
                    if reply != QMessageBox.Yes:
                        return
            
            generator = ExcelGenerator()
            generator.generate(self.document_data, file_path)
            
            # Сохраняем документ в БД
            self.document_store.save_document(self.document_data, str(file_path))

            # Документ успешно сгенерирован — сбрасываем черновик
            self.draft_manager.discard()
            self._document_dirty = False
            self._last_saved_at = None
            self._refresh_draft_status_indicator()

            self.audit_log.success(
                "generate_document",
                f"Сгенерирован документ {self.document_data.get_display_number()}: {file_path.name}",
            )

            QMessageBox.information(self, "Успех", f"Документ успешно создан:\n{file_path}")
            
            # Автоматически открываем файл, если настройка включена
            if self.settings_manager.get_open_after_generate():
                try:
                    import os
                    os.startfile(str(file_path))
                except Exception as e:
                    # Если не удалось открыть, просто логируем, но не показываем ошибку пользователю
                    import logging
                    logging.warning(f"Не удалось автоматически открыть файл: {e}")
            
            # Сохраняем дату перед созданием нового документа
            if self.document_data.implementation_date:
                self._last_used_date = self.document_data.implementation_date
            
            # Создаём новый документ с новым номером
            self.new_document()
            
        except PermissionError as e:
            # Специальная обработка ошибок доступа
            error_msg = str(e)
            if "открыт" in error_msg.lower() or "open" in error_msg.lower():
                QMessageBox.critical(
                    self, 
                    "Ошибка доступа", 
                    f"Не удалось создать документ:\n\n{error_msg}\n\n"
                    f"Действия:\n"
                    f"1. Закройте файл в Excel, если он открыт\n"
                    f"2. Проверьте, не открыт ли файл в другой программе\n"
                    f"3. Попробуйте сохранить документ снова"
                )
            else:
                QMessageBox.critical(
                    self, 
                    "Ошибка доступа", 
                    f"Не удалось создать документ:\n\n{error_msg}\n\n"
                    f"Проверьте права доступа к папке:\n{target_dir}"
                )
        except FileNotFoundError as e:
            QMessageBox.critical(
                self, 
                "Ошибка", 
                f"Не удалось создать документ:\n\n{str(e)}\n\n"
                f"Проверьте, что все необходимые файлы и папки существуют."
            )
        except Exception as e:
            # Общая обработка других ошибок
            error_msg = str(e)
            if "Permission denied" in error_msg or "Errno 13" in error_msg:
                QMessageBox.critical(
                    self, 
                    "Ошибка доступа", 
                    f"Не удалось создать документ:\n\n"
                    f"Ошибка доступа к файлу: {file_path}\n\n"
                    f"Возможные причины:\n"
                    f"1. Файл открыт в Excel или другой программе\n"
                    f"2. Нет прав на запись в папку\n"
                    f"3. Сетевая папка недоступна\n\n"
                    f"Детали ошибки: {error_msg}"
                )
            else:
                QMessageBox.critical(
                    self, 
                    "Ошибка", 
                    f"Не удалось создать документ:\n\n{error_msg}\n\n"
                    f"Путь: {file_path}"
                )

    def ensure_products_selected_for_generation(self) -> bool:
        """Проверить, что для документа выбрана хотя бы одна машина."""
        if self.document_data.products:
            return True

        products = [name for _, name in self.product_store.get_all_products()]
        if not products:
            QMessageBox.warning(
                self,
                "Ошибка",
                "Нет доступных машин. Сначала создайте хотя бы одну машину на вкладке 'Изменения материалов'."
            )
            self._on_wizard_step_selected(1)
            return False

        dialog = MachineSelectionDialog(products, self)
        if not dialog.exec():
            self._on_wizard_step_selected(1)
            return False

        selected_products = dialog.get_selected_products()
        if not selected_products:
            self._on_wizard_step_selected(1)
            return False

        self.changes_widget.set_selected_products(selected_products)
        return True
    
    def new_document(self):
        """Создать новый документ"""
        if hasattr(self, "replacement_sets_widget") and self.replacement_sets_widget:
            if self.replacement_sets_widget.has_unsaved_changes():
                if not self.replacement_sets_widget.prompt_unsaved_changes():
                    return
        # Проверяем, есть ли несохранённые изменения
        if self.document_data.part_changes or self.document_data.products:
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                "Создать новый документ? Текущие данные будут очищены.",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
        
        # Создаём новый документ
        self.document_data = DocumentData()

        self.audit_log.info("new_document", "Создан новый документ")

        self._apply_last_used_date()
        
        self.current_additional_page = None  # Сбрасываем активную доп. страницу
        self.next_additional_page = 1        # Всегда начинаем с 1+ для новых кликов
        self.doc_info_widget.document_data = self.document_data
        self.changes_widget.document_data = self.document_data
        self.replacement_sets_widget.document_data = self.document_data
        self.doc_info_widget.refresh()
        self.changes_widget.refresh()
        self.changes_widget.update_product_filter()
        self.refresh_status_bar()
        # Сбрасываем индикатор черновика
        self.draft_manager.discard()
        self._document_dirty = False
        self._last_saved_at = None
        self._refresh_draft_status_indicator()

    def open_document(self):
        """Открыть существующий документ"""
        if hasattr(self, "replacement_sets_widget") and self.replacement_sets_widget:
            if self.replacement_sets_widget.has_unsaved_changes():
                if not self.replacement_sets_widget.prompt_unsaved_changes():
                    return
        logger = logging.getLogger(__name__)
        from app.gui.document_selection_dialog import DocumentSelectionDialog
        
        dialog = DocumentSelectionDialog(self.document_store, self)
        if dialog.exec():
            document_number, year, month = dialog.get_selected_document()
            if document_number:
                loaded_data = self.document_store.load_document(document_number, year, month)
                if loaded_data:
                    self.document_data = loaded_data
                    self._last_used_date = self.document_data.implementation_date
                    self._last_used_validity_period = self.document_data.validity_period
                    self._last_used_year = self.document_data.document_year
                    # Определяем максимальный номер доп. страницы для восстановления счётчика
                    max_page = 0
                    pages_found = []
                    all_parts_info = []
                    for part_change in self.document_data.part_changes:
                        all_parts_info.append((part_change.part, part_change.additional_page_number))
                        if part_change.additional_page_number is not None:
                            max_page = max(max_page, part_change.additional_page_number)
                            pages_found.append((part_change.part, part_change.additional_page_number))
                    
                    logger.info(f"[open_document] Загружено деталей: {len(self.document_data.part_changes)}")
                    logger.info(f"[open_document] Все детали с их номерами страниц: {all_parts_info}")
                    logger.info(f"[open_document] Найдено деталей с доп. страницами: {len(pages_found)}")
                    logger.info(f"[open_document] Максимальный номер страницы: {max_page}")
                    logger.info(f"[open_document] Детали по страницам: {pages_found}")
                    
                    # Активную страницу сбрасываем, но следующий клик пойдёт на новую страницу
                    self.current_additional_page = None
                    self.next_additional_page = max_page + 1 if max_page > 0 else 1
                    logger.info(f"[open_document] Установлен next_additional_page = {self.next_additional_page}")
                    self.doc_info_widget.document_data = self.document_data
                    self.changes_widget.document_data = self.document_data
                    self.replacement_sets_widget.document_data = self.document_data
                    self.doc_info_widget.refresh()
                    self.changes_widget.refresh()
                    self.changes_widget.update_product_filter()
                    display_num = self.document_data.get_display_number()
                    self.refresh_status_bar()
                    # Загруженный документ из БД — сбрасываем черновик
                    self.draft_manager.discard()
                    self._document_dirty = False
                    self._last_saved_at = None
                    self._refresh_draft_status_indicator()
                    self.audit_log.info("open_document", f"Открыт документ {display_num}")
                    QMessageBox.information(self, "Успех", f"Документ {display_num} загружен")
                else:
                    QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить документ №{document_number}")

    def _set_current_additional_page(self, page) -> None:
        """Установить активную дополнительную страницу (None = основная)."""
        self.current_additional_page = page
        if hasattr(self, "changes_widget"):
            self.changes_widget.refresh_pages_breadcrumb()
            self.changes_widget.refresh()
        if hasattr(self, "status_bar_widget"):
            if page is None:
                self.status_bar_widget.set_ready_text("Активна основная страница")
            else:
                self.status_bar_widget.set_ready_text(f"Активна доп. страница {page}+")

    def add_additional_page_data(self):
        """Добавить данные для дополнительной страницы"""
        logger = logging.getLogger(__name__)
        
        # Пересчитываем максимальный номер доп. страницы из существующих данных
        # Это гарантирует, что каждая новая партия данных всегда идёт на новую страницу
        max_page = 0
        pages_found = []
        all_parts_info = []
        for part_change in self.document_data.part_changes:
            all_parts_info.append((part_change.part, part_change.additional_page_number))
            if part_change.additional_page_number is not None:
                max_page = max(max_page, part_change.additional_page_number)
                pages_found.append((part_change.part, part_change.additional_page_number))
        
        logger.info(f"[add_additional_page_data] Всего деталей в document_data: {len(self.document_data.part_changes)}")
        logger.info(f"[add_additional_page_data] Все детали с их номерами страниц: {all_parts_info}")
        logger.info(f"[add_additional_page_data] Найдено деталей с доп. страницами: {len(pages_found)}")
        logger.info(f"[add_additional_page_data] Максимальный номер страницы: {max_page}")
        logger.info(f"[add_additional_page_data] Детали по страницам: {pages_found}")
        
        # Устанавливаем следующий номер страницы (max + 1, или 1, если нет доп. страниц)
        self.current_additional_page = max_page + 1 if max_page > 0 else 1
        # Обновляем счётчик для будущих кликов
        self.next_additional_page = self.current_additional_page + 1
        
        logger.info(f"[add_additional_page_data] Установлен current_additional_page = {self.current_additional_page}")
        logger.info(f"[add_additional_page_data] Установлен next_additional_page = {self.next_additional_page}")
        
        page_label = f"{self.current_additional_page}+"
        
        message = (
            f"Режим добавления данных для страницы {page_label} активирован.\n\n"
            f"Все детали и материалы, которые вы добавите сейчас, будут размещены на странице {page_label}.\n\n"
            f"Переключить на вкладку 'Изменения материалов'?"
        )
        
        reply = QMessageBox.question(
            self,
            f"Добавить данные для страницы {page_label}",
            message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            # Переключаемся на вкладку изменений материалов
            self._on_wizard_step_selected(1)
            # Фокусируемся на виджете изменений
            self.changes_widget.setFocus()
    
    def clear_data(self):
        """Очистить все данные"""
        if hasattr(self, "replacement_sets_widget") and self.replacement_sets_widget:
            if self.replacement_sets_widget.has_unsaved_changes():
                if not self.replacement_sets_widget.prompt_unsaved_changes():
                    return

        reply = QMessageBox.question(
            self,
            "Подтверждение",
            "Очистить все введённые данные?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.audit_log.info("clear_data", "Данные документа очищены")
            self.document_data = DocumentData()
            self._apply_last_used_date()
            self.current_additional_page = None  # Сбрасываем активную доп. страницу
            self.next_additional_page = 1        # Следующий клик создаст страницу 1+
            self.doc_info_widget.document_data = self.document_data
            self.changes_widget.document_data = self.document_data
            self.replacement_sets_widget.document_data = self.document_data
            self.doc_info_widget.refresh()
            self.changes_widget.refresh()
            self.draft_manager.discard()
            self._document_dirty = False
            self._last_saved_at = None
            self._refresh_draft_status_indicator()
            self.changes_widget.update_product_filter()

    def closeEvent(self, event):
        """Проверка несохранённых данных перед закрытием окна"""
        if hasattr(self, "replacement_sets_widget") and self.replacement_sets_widget:
            if self.replacement_sets_widget.has_unsaved_changes():
                if not self.replacement_sets_widget.prompt_unsaved_changes():
                    event.ignore()
                    return
        super().closeEvent(event)
