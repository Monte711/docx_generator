import sys
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QMenuBar, QStatusBar, QFileDialog, QMessageBox,
    QListWidget, QSplitter, QListWidgetItem # Добавлен QListWidgetItem
)
from PySide6.QtGui import QAction, QKeySequence, QCloseEvent
from PySide6.QtCore import Qt, Slot

from models.project import Project
from models.docx_handler import DocxHandler
from views.simple_key_editor import SimpleKeyEditorWidget
from views.table_editor import TableEditorWidget # Импортируем редактор таблиц

class MainWindow(QMainWindow):
    """
    Главное окно приложения.
    (Версия 5: Интегрирован редактор таблиц)
    """
    def __init__(self, parent=None):
        super().__init__(parent)

        self.project = Project()
        self.docx_handler = DocxHandler()

        self.setWindowTitle(self._build_window_title())
        self.resize(1100, 750) # Еще немного увеличим

        # --- Центральный виджет с разделителем ---
        self._central_widget = QSplitter(Qt.Orientation.Horizontal)
        self.setCentralWidget(self._central_widget)

        # Левая панель (список ключей)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.addWidget(QLabel("Найденные ключи и таблицы:"))
        self.keys_list_widget = QListWidget()
        self.keys_list_widget.currentItemChanged.connect(self._on_key_selected)
        left_layout.addWidget(self.keys_list_widget)
        self._central_widget.addWidget(left_panel)

        # Правая панель (редактор)
        right_panel = QWidget()
        self.editor_layout = QVBoxLayout(right_panel)
        self.editor_layout.setContentsMargins(5, 0, 0, 0)

        # Создаем экземпляры редакторов
        self.simple_key_editor = SimpleKeyEditorWidget()
        self.simple_key_editor.data_changed.connect(self._on_editor_data_changed)
        self.editor_layout.addWidget(self.simple_key_editor)

        self.table_editor = TableEditorWidget() # Создаем редактор таблиц
        self.table_editor.data_changed.connect(self._on_editor_data_changed) # Подключаем сигнал
        self.editor_layout.addWidget(self.table_editor) # Добавляем в layout

        self.editor_layout.addStretch(1) # Растягивающийся элемент

        self._central_widget.addWidget(right_panel)
        self._central_widget.setSizes([350, 750]) # Немного изменим пропорции

        # --- Меню и Статус-бар ---
        self._create_menus()
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Приложение готово.")

        self._update_ui_state()

    # ... (методы _build_window_title, _update_window_title, _update_ui_state, _create_menus, _check_unsaved_changes без изменений) ...
    def _build_window_title(self) -> str:
        base_title = "Генератор документов DOCX"
        project_name = self.project.get_project_filename()
        modified_marker = "*" if self.project.is_modified else ""
        return f"{project_name}{modified_marker} - {base_title}"

    @Slot()
    def _update_window_title(self):
        self.setWindowTitle(self._build_window_title())

    @Slot()
    def _update_ui_state(self):
        project_active = bool(self.project.template_paths) or self.project.filepath is not None
        can_save = self.project.is_modified
        can_save_as = project_active
        self.save_project_action.setEnabled(can_save)
        self.save_project_as_action.setEnabled(can_save_as)
        self.add_template_action.setEnabled(True)
        self.generate_docs_action.setEnabled(bool(self.project.template_paths and self.project.output_path))
        self._update_window_title()

    def _create_menus(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("&Файл")
        new_project_action = QAction("&Новый проект", self)
        new_project_action.setShortcut(QKeySequence.StandardKey.New)
        new_project_action.triggered.connect(self._on_new_project)
        file_menu.addAction(new_project_action)
        open_project_action = QAction("&Открыть проект...", self)
        open_project_action.setShortcut(QKeySequence.StandardKey.Open)
        open_project_action.triggered.connect(self._on_open_project)
        file_menu.addAction(open_project_action)
        self.save_project_action = QAction("&Сохранить проект", self)
        self.save_project_action.setShortcut(QKeySequence.StandardKey.Save)
        self.save_project_action.triggered.connect(self._on_save_project)
        file_menu.addAction(self.save_project_action)
        self.save_project_as_action = QAction("Сохранить проект &как...", self)
        self.save_project_as_action.setShortcut(QKeySequence.StandardKey.SaveAs)
        self.save_project_as_action.triggered.connect(self._on_save_project_as)
        file_menu.addAction(self.save_project_as_action)
        file_menu.addSeparator()
        exit_action = QAction("&Выход", self)
        exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        project_menu = menu_bar.addMenu("&Проект")
        self.add_template_action = QAction("Добавить &шаблон...", self)
        self.add_template_action.triggered.connect(self._on_add_template)
        project_menu.addAction(self.add_template_action)
        self.generate_docs_action = QAction("&Сгенерировать документы...", self)
        # self.generate_docs_action.triggered.connect(self._on_generate_docs)
        project_menu.addAction(self.generate_docs_action)
        help_menu = menu_bar.addMenu("&Справка")
        about_action = QAction("&О программе", self)
        # about_action.triggered.connect(self._on_about)
        help_menu.addAction(about_action)

    def _check_unsaved_changes(self) -> bool:
        if not self.project.is_modified: return True
        project_name = self.project.get_project_filename()
        reply = QMessageBox.question(
            self, "Несохраненные изменения",
            f"В проекте '{project_name}' есть несохраненные изменения.\nХотите сохранить их перед продолжением?",
            QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save
        )
        if reply == QMessageBox.StandardButton.Save: return self._on_save_project()
        elif reply == QMessageBox.StandardButton.Cancel: return False
        else: return True # Discard

    @Slot()
    def _on_new_project(self):
        if not self._check_unsaved_changes(): return
        self.project.reset()
        self.statusBar().showMessage("Создан новый пустой проект.")
        self.keys_list_widget.clear()
        self.simple_key_editor.clear_editor()
        self.table_editor.clear_editor() # Очищаем и редактор таблиц
        self._update_ui_state()
        print("Действие: Новый проект")

    @Slot()
    def _on_open_project(self):
        if not self._check_unsaved_changes(): return
        start_dir = str(self.project.filepath.parent) if self.project.filepath else str(Path.home())
        file_filter = "Проекты DocxFormatter (*.dfp);;Все файлы (*)"
        filepath_str, _ = QFileDialog.getOpenFileName(self, "Открыть проект", start_dir, file_filter)

        if filepath_str:
            self.simple_key_editor.clear_editor()
            self.table_editor.clear_editor() # Очищаем редакторы перед загрузкой

            if self.project.load(filepath_str):
                self.statusBar().showMessage(f"Проект '{self.project.get_project_filename()}' загружен.")
                self._update_keys_list_widget()
            else:
                QMessageBox.warning(self, "Ошибка загрузки", f"Не удалось загрузить проект из файла:\n{filepath_str}")
                self.statusBar().showMessage("Ошибка загрузки проекта.")
                self.keys_list_widget.clear()
            self._update_ui_state()
            print(f"Действие: Открыть проект - {filepath_str}")
        else:
            self.statusBar().showMessage("Открытие проекта отменено.")

    # ... (методы _on_save_project, _on_save_project_as без изменений) ...
    @Slot()
    def _on_save_project(self) -> bool:
        if not self.project.filepath: return self._on_save_project_as()
        if self.project.save():
            self.statusBar().showMessage(f"Проект сохранен в '{self.project.filepath.name}'.")
            self._update_ui_state()
            print(f"Действие: Сохранить проект - {self.project.filepath}")
            return True
        else:
            QMessageBox.critical(self, "Ошибка сохранения", f"Не удалось сохранить проект в файл:\n{self.project.filepath}")
            self.statusBar().showMessage("Ошибка сохранения проекта.")
            return False

    @Slot()
    def _on_save_project_as(self) -> bool:
        start_dir = str(self.project.filepath.parent) if self.project.filepath else str(Path.home())
        start_filename = self.project.filepath.name if self.project.filepath else "Новый проект.dfp"
        start_path = str(Path(start_dir) / start_filename)
        file_filter = "Проекты DocxFormatter (*.dfp);;Все файлы (*)"
        filepath_str, _ = QFileDialog.getSaveFileName(self, "Сохранить проект как...", start_path, file_filter)
        if filepath_str:
            if self.project.save(filepath_str):
                self.statusBar().showMessage(f"Проект сохранен как '{self.project.filepath.name}'.")
                self._update_ui_state()
                print(f"Действие: Сохранить проект как - {filepath_str}")
                return True
            else:
                QMessageBox.critical(self, "Ошибка сохранения", f"Не удалось сохранить проект в файл:\n{filepath_str}")
                self.statusBar().showMessage("Ошибка сохранения проекта.")
                return False
        else:
            self.statusBar().showMessage("Сохранение проекта отменено.")
            return False

    # ... (метод _on_add_template без изменений) ...
    @Slot()
    def _on_add_template(self):
        start_dir = str(self.project.filepath.parent) if self.project.filepath else str(Path.home())
        file_filter = "Документы Word (*.docx);;Все файлы (*)"
        filepath_str, _ = QFileDialog.getOpenFileName(self, "Добавить шаблон DOCX", start_dir, file_filter)
        if filepath_str:
            template_path = Path(filepath_str)
            if self.project.add_template(filepath_str):
                self.statusBar().showMessage(f"Шаблон '{template_path.name}' добавлен. Идет сканирование...")
                QApplication.processEvents()
                scan_results = self.docx_handler.find_keys_in_template(template_path)
                found_keys = scan_results.get('keys', set())
                found_tables = scan_results.get('dynamic_tables', set())
                keys_added_count = 0
                tables_added_count = 0
                if found_keys:
                    for key in found_keys:
                         if key not in self.project.keys_data:
                             self.project.add_found_key(key)
                             keys_added_count += 1
                if found_tables:
                    for table_id in found_tables:
                         if table_id not in self.project.keys_data:
                            self.project.add_found_table(table_id)
                            tables_added_count += 1
                self.statusBar().showMessage(f"Шаблон '{template_path.name}' добавлен. Найдено новых ключей: {keys_added_count}, таблиц: {tables_added_count}.")
                self._update_keys_list_widget()
                self._update_ui_state()
                print(f"Действие: Добавить шаблон - {filepath_str}")
                print(f"Найденные ключи: {found_keys}")
                print(f"Найденные таблицы: {found_tables}")
            else:
                QMessageBox.warning(self, "Ошибка", f"Не удалось добавить шаблон:\n{filepath_str}\nВозможно, файл уже добавлен или не является DOCX.")
                self.statusBar().showMessage("Ошибка добавления шаблона.")
        else:
            self.statusBar().showMessage("Добавление шаблона отменено.")

    # ... (метод _update_keys_list_widget без изменений) ...
    def _update_keys_list_widget(self):
        current_selection = self.keys_list_widget.currentItem()
        current_text = current_selection.text() if current_selection else None
        self.keys_list_widget.clear()
        if not self.project.keys_data: return
        sorted_key_ids = sorted(self.project.keys_data.keys())
        item_to_select = None
        for key_id in sorted_key_ids:
            item_text = key_id
            key_info = self.project.keys_data[key_id]
            if key_info.get('type') == 'dynamic_table':
                item_text = f"[ТАБЛИЦА] {key_id}"
            # Добавляем QListWidgetItem, чтобы можно было хранить исходный key_id
            list_item = QListWidgetItem(item_text)
            list_item.setData(Qt.ItemDataRole.UserRole, key_id) # Сохраняем ID в данных элемента
            self.keys_list_widget.addItem(list_item)
            if item_text == current_text: # Сравниваем по тексту для восстановления
                item_to_select = list_item
        if item_to_select:
            self.keys_list_widget.setCurrentItem(item_to_select)


    @Slot()
    def _on_key_selected(self, current_item: QListWidgetItem | None, previous_item: QListWidgetItem | None):
        """Обрабатывает выбор элемента в списке ключей и показывает нужный редактор."""
        # Сначала скрываем оба редактора
        self.simple_key_editor.setVisible(False)
        self.table_editor.setVisible(False)

        if current_item:
            item_text = current_item.text()
            # Получаем ID из данных элемента, а не из текста
            key_id = current_item.data(Qt.ItemDataRole.UserRole)
            is_table = item_text.startswith("[ТАБЛИЦА]")

            if key_id: # Убедимся, что ID есть
                key_data = self.project.get_key_data(key_id)

                if is_table:
                    print(f"Выбрана таблица: {key_id}")
                    self.table_editor.set_table_data(key_id, key_data) # Показываем редактор таблиц
                    self.statusBar().showMessage(f"Выбрана таблица: {key_id}")
                else:
                    print(f"Выбран ключ: {key_id}")
                    self.simple_key_editor.set_key_data(key_id, key_data) # Показываем редактор ключей
                    self.statusBar().showMessage(f"Выбран ключ: {key_id}")
            else:
                 self.statusBar().showMessage("Ошибка: Не удалось получить ID для выбранного элемента.")

        else:
            # Если ничего не выбрано
            self.simple_key_editor.clear_editor()
            self.table_editor.clear_editor()
            self.statusBar().showMessage("Ключ не выбран.")

    @Slot()
    def _on_editor_data_changed(self):
        """Сохраняет изменения из активного редактора в модель Project."""
        if self.simple_key_editor.isVisible():
            key_id = self.simple_key_editor.get_current_key_id()
            if key_id:
                edited_data = self.simple_key_editor.get_edited_data()
                self.project.update_key_data(key_id, edited_data)
                self._update_ui_state()
                print(f"Данные ключа '{key_id}' обновлены в модели.")
        elif self.table_editor.isVisible(): # Проверяем видимость редактора таблиц
            table_id = self.table_editor.get_current_table_id()
            if table_id:
                edited_data = self.table_editor.get_edited_data()
                # Используем тот же метод update_key_data, он должен уметь
                # обновлять и данные таблиц (поле 'data' внутри словаря)
                self.project.update_key_data(table_id, edited_data)
                self._update_ui_state()
                print(f"Данные таблицы '{table_id}' обновлены в модели.")


    def closeEvent(self, event: QCloseEvent):
        if self._check_unsaved_changes(): event.accept()
        else: event.ignore()

# ... (остальные заглушки и блок if __name__ == '__main__') ...
# def _on_generate_docs(self): print("Действие: Сгенерировать документы")
# def _on_about(self): print("Действие: О программе")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())