from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QPushButton, QAbstractItemView, QHeaderView, QSizePolicy, QMessageBox
)
from PySide6.QtCore import Qt, Signal, Slot
from PySide6.QtGui import QIcon # Для иконок кнопок (опционально)

# TODO: Рассмотреть возможность использования QStyledItemDelegate для кастомных редакторов ячеек (например, QDateEdit)

class TableEditorWidget(QWidget):
    """
    Виджет для редактирования данных динамической таблицы.
    """
    # Сигнал, испускаемый при изменении данных пользователем
    data_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)

        self._current_table_id: str | None = None
        self._column_keys: list[str] = [] # Внутренние ключи для столбцов (из 'template_keys')
        self._project_data_ref: dict | None = None # Ссылка на данные таблицы в Project

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # --- Заголовок ---
        self.table_id_label = QLabel("Таблица: Не выбрана")
        self.table_id_label.setStyleSheet("font-weight: bold;")
        main_layout.addWidget(self.table_id_label)

        # --- Панель кнопок ---
        button_layout = QHBoxLayout()
        self.add_row_button = QPushButton("Добавить строку")
        # self.add_row_button.setIcon(QIcon("path/to/add_icon.png")) # Опционально
        self.delete_row_button = QPushButton("Удалить строку")
        self.move_up_button = QPushButton("Вверх")
        self.move_down_button = QPushButton("Вниз")

        button_layout.addWidget(self.add_row_button)
        button_layout.addWidget(self.delete_row_button)
        button_layout.addStretch(1) # Пространство между группами кнопок
        button_layout.addWidget(self.move_up_button)
        button_layout.addWidget(self.move_down_button)
        main_layout.addLayout(button_layout)

        # --- Таблица ---
        self.table_widget = QTableWidget()
        # Настройки таблицы
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows) # Выделение строк целиком
        self.table_widget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection) # Только одну строку можно выбрать
        self.table_widget.verticalHeader().setVisible(False) # Скрываем нумерацию строк по умолчанию (будет в первом столбце)
        self.table_widget.horizontalHeader().setStretchLastSection(True) # Последний столбец растягивается
        self.table_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        main_layout.addWidget(self.table_widget)

        # --- Подключение сигналов ---
        self.add_row_button.clicked.connect(self._add_row)
        self.delete_row_button.clicked.connect(self._delete_row)
        self.move_up_button.clicked.connect(self._move_row_up)
        self.move_down_button.clicked.connect(self._move_row_down)
        self.table_widget.itemChanged.connect(self._on_item_changed) # Сигнал изменения ячейки

        # Изначально виджет скрыт
        self.setVisible(False)

    def set_table_data(self, table_id: str, data: dict | None):
        """
        Загружает данные выбранной таблицы в QTableWidget.

        Args:
            table_id: Идентификатор таблицы (например, "HardwareList").
            data: Словарь с данными таблицы из модели Project. Ожидается:
                  {'type': 'dynamic_table', 'columns': [...],
                   'template_keys': [...], 'data': [{'col1': val1,...}, ...]}
        """
        self._current_table_id = table_id
        self._project_data_ref = data # Сохраняем ссылку для get_edited_data
        self.table_id_label.setText(f"Таблица: {table_id}")

        self.table_widget.blockSignals(True) # Блокируем сигналы на время заполнения
        self.table_widget.clear() # Очищаем все содержимое и заголовки

        if data and data.get('type') == 'dynamic_table':
            self._column_keys = data.get('template_keys', [])
            table_data_rows = data.get('data', [])

            # --- Настройка столбцов ---
            # Первый столбец для нумерации "№ п/п"
            column_count = len(self._column_keys) + 1
            self.table_widget.setColumnCount(column_count)

            column_labels = ["№ п/п"]
            # Пытаемся получить заголовки из 'columns', если нет - генерируем
            provided_columns = data.get('columns', [])
            if len(provided_columns) == len(self._column_keys):
                 # Используем предоставленные имена, если они есть и совпадают по кол-ву
                 column_labels.extend(provided_columns)
            else:
                # Генерируем заголовки из ключей или просто "Столбец N"
                for i, key in enumerate(self._column_keys):
                    # Убираем скобки {{}} и пробелы для заголовка
                    header = key.strip('{} ') or f"Столбец {i+1}"
                    column_labels.append(header)

            self.table_widget.setHorizontalHeaderLabels(column_labels)

            # --- Заполнение строк ---
            self.table_widget.setRowCount(len(table_data_rows))
            for row_idx, row_data_dict in enumerate(table_data_rows):
                # 1. Номер строки
                num_item = QTableWidgetItem(str(row_idx + 1))
                num_item.setFlags(num_item.flags() & ~Qt.ItemFlag.ItemIsEditable) # Делаем нередактируемым
                num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table_widget.setItem(row_idx, 0, num_item)

                # 2. Данные из словаря
                for col_idx, col_key in enumerate(self._column_keys):
                    value = str(row_data_dict.get(col_key, '')) # Получаем значение по ключу
                    item = QTableWidgetItem(value)
                    self.table_widget.setItem(row_idx, col_idx + 1, item) # +1 из-за столбца "№ п/п"

            # Настраиваем ширину столбцов после заполнения
            self.table_widget.resizeColumnsToContents()
            self.table_widget.horizontalHeader().setStretchLastSection(True)

            self.setVisible(True) # Показываем редактор
        else:
            # Если данных нет или тип неверный
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            self.setVisible(False)

        self.table_widget.blockSignals(False) # Разблокируем сигналы

    def get_edited_data(self) -> dict:
        """
        Собирает данные из QTableWidget и возвращает их в формате для Project.
        """
        if not self._current_table_id or self._project_data_ref is None:
            return {}

        edited_data_list = []
        row_count = self.table_widget.rowCount()
        col_count = self.table_widget.columnCount()

        if col_count != len(self._column_keys) + 1:
            print("Ошибка: несоответствие количества столбцов и ключей!")
            return {} # Или вернуть старые данные?

        for row_idx in range(row_count):
            row_data_dict = {}
            for col_idx, col_key in enumerate(self._column_keys):
                item = self.table_widget.item(row_idx, col_idx + 1) # +1 из-за "№ п/п"
                value = item.text() if item else ''
                row_data_dict[col_key] = value
            edited_data_list.append(row_data_dict)

        # Возвращаем обновленные данные, сохраняя остальные поля из оригинала
        updated_project_data = self._project_data_ref.copy()
        updated_project_data['data'] = edited_data_list
        return updated_project_data


    def get_current_table_id(self) -> str | None:
        """Возвращает ID таблицы, которая сейчас редактируется."""
        return self._current_table_id

    def clear_editor(self):
        """Очищает таблицу и скрывает виджет."""
        self._current_table_id = None
        self._column_keys = []
        self._project_data_ref = None
        self.table_id_label.setText("Таблица: Не выбрана")
        self.table_widget.blockSignals(True)
        self.table_widget.clearContents()
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)
        self.table_widget.blockSignals(False)
        self.setVisible(False)

    def _renumber_rows(self):
        """Обновляет нумерацию в первом столбце."""
        self.table_widget.blockSignals(True) # Блокируем сигнал itemChanged
        for row_idx in range(self.table_widget.rowCount()):
            item = self.table_widget.item(row_idx, 0)
            if item:
                item.setText(str(row_idx + 1))
            else: # Если ячейки нет, создаем
                num_item = QTableWidgetItem(str(row_idx + 1))
                num_item.setFlags(num_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table_widget.setItem(row_idx, 0, num_item)
        self.table_widget.blockSignals(False)

    @Slot()
    def _add_row(self):
        """Добавляет пустую строку в конец таблицы."""
        current_row = self.table_widget.currentRow() # Получаем текущую строку для вставки после нее
        if current_row < 0: # Если ничего не выбрано, добавляем в конец
            current_row = self.table_widget.rowCount()
        else:
            current_row += 1 # Вставляем после выбранной

        self.table_widget.insertRow(current_row)
        self._renumber_rows() # Обновляем нумерацию
        self.data_changed.emit() # Сигналим об изменении

    @Slot()
    def _delete_row(self):
        """Удаляет выбранную строку."""
        current_row = self.table_widget.currentRow()
        if current_row >= 0:
            confirm = QMessageBox.question(self, "Удаление строки",
                                           f"Вы уверены, что хотите удалить строку {current_row + 1}?",
                                           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                           QMessageBox.StandardButton.No)
            if confirm == QMessageBox.StandardButton.Yes:
                self.table_widget.removeRow(current_row)
                self._renumber_rows() # Обновляем нумерацию
                self.data_changed.emit() # Сигналим об изменении

    @Slot()
    def _move_row_up(self):
        """Перемещает выбранную строку на одну позицию вверх."""
        current_row = self.table_widget.currentRow()
        if current_row > 0: # Нельзя переместить самую верхнюю строку
            self.table_widget.blockSignals(True) # Блокируем сигналы
            # Сохраняем данные перемещаемой строки
            row_data = []
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.takeItem(current_row, col)
                row_data.append(item)

            # Удаляем строку
            self.table_widget.removeRow(current_row)
            # Вставляем строку на новую позицию
            new_row_index = current_row - 1
            self.table_widget.insertRow(new_row_index)
            # Вставляем данные обратно
            for col, item in enumerate(row_data):
                self.table_widget.setItem(new_row_index, col, item)

            self.table_widget.selectRow(new_row_index) # Выбираем перемещенную строку
            self.table_widget.blockSignals(False) # Разблокируем сигналы
            self._renumber_rows()
            self.data_changed.emit()

    @Slot()
    def _move_row_down(self):
        """Перемещает выбранную строку на одну позицию вниз."""
        current_row = self.table_widget.currentRow()
        row_count = self.table_widget.rowCount()
        if current_row >= 0 and current_row < row_count - 1: # Нельзя переместить самую нижнюю
            self.table_widget.blockSignals(True)
            row_data = []
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.takeItem(current_row, col)
                row_data.append(item)

            self.table_widget.removeRow(current_row)
            new_row_index = current_row + 1
            self.table_widget.insertRow(new_row_index)
            for col, item in enumerate(row_data):
                self.table_widget.setItem(new_row_index, col, item)

            self.table_widget.selectRow(new_row_index)
            self.table_widget.blockSignals(False)
            self._renumber_rows()
            self.data_changed.emit()


    @Slot(QTableWidgetItem)
    def _on_item_changed(self, item: QTableWidgetItem):
        """Слот, вызываемый при изменении содержимого ячейки."""
        # Игнорируем изменения в первом столбце (нумерация)
        if item.column() == 0:
            return
        print(f"Ячейка [{item.row()}, {item.column()}] изменена: {item.text()}") # Отладка
        self.data_changed.emit() # Сигналим об изменении данных