from PySide6.QtWidgets import (
        QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QCheckBox, QSizePolicy,
        QFrame, QSpacerItem # Добавлены QFrame и QSpacerItem
    )
from PySide6.QtCore import Qt, Signal, Slot

class SimpleKeyEditorWidget(QWidget):
    """
    Виджет для редактирования значения простого ключа {{...}}.
    """
    # Сигнал, испускаемый при изменении данных пользователем
    data_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)

        self._current_key_id: str | None = None # Храним ID текущего редактируемого ключа

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0) # Убираем отступы у основного layout

        # --- Виджеты редактора ---
        self.key_label = QLabel("Ключ: Не выбран")
        self.key_label.setStyleSheet("font-weight: bold;") # Жирный шрифт для имени ключа

        self.value_label = QLabel("Значение:")
        # Используем QTextEdit для возможности многострочного ввода
        self.value_edit = QTextEdit()
        self.value_edit.setAcceptRichText(False) # Принимаем только простой текст
        self.value_edit.setMinimumHeight(60) # Минимальная высота
        # Устанавливаем политику размера, чтобы поле могло растягиваться
        self.value_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.freeze_checkbox = QCheckBox("Заморозить (запретить авто-обновление/редактирование)")

        # --- Статус валидации (пока заглушка) ---
        self.status_label = QLabel("Статус: -")
        self.status_label.setStyleSheet("color: gray;")

        # --- Сборка Layout ---
        main_layout.addWidget(self.key_label)

        # Горизонтальный layout для метки "Значение" и статуса
        value_header_layout = QHBoxLayout()
        value_header_layout.addWidget(self.value_label)
        value_header_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)) # Растягивающийся промежуток
        value_header_layout.addWidget(self.status_label)
        main_layout.addLayout(value_header_layout)

        main_layout.addWidget(self.value_edit)
        main_layout.addWidget(self.freeze_checkbox)

        # Добавляем разделитель для визуального отделения (опционально)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(line)

        # Добавляем растягивающийся элемент в конец, чтобы прижать все вверх
        main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))


        # --- Подключение сигналов ---
        self.value_edit.textChanged.connect(self._on_data_edited)
        self.freeze_checkbox.stateChanged.connect(self._on_data_edited)

        # Изначально виджет скрыт
        self.setVisible(False)

    @Slot()
    def _on_data_edited(self):
        """Слот, вызываемый при изменении текста или состояния чекбокса."""
        # Испускаем сигнал, что данные изменились
        self.data_changed.emit()

    def set_key_data(self, key_id: str, data: dict | None):
        """
        Загружает данные выбранного ключа в виджеты редактора.

        Args:
            key_id: Имя ключа (например, "{{ORG_NAME}}").
            data: Словарь с данными ключа из модели Project
                  (ожидается {'value': '...', 'status': '...', 'is_frozen': ...}).
                  Может быть None, если ключ не найден (хотя этого не должно быть).
        """
        self._current_key_id = key_id
        self.key_label.setText(f"Ключ: {key_id}")

        if data:
            # Блокируем сигналы на время установки данных, чтобы не вызвать data_changed
            self.value_edit.blockSignals(True)
            self.freeze_checkbox.blockSignals(True)

            self.value_edit.setPlainText(data.get('value', ''))
            self.freeze_checkbox.setChecked(data.get('is_frozen', False))
            # Обновляем статус (пока просто текстом)
            status = data.get('status', 'unknown')
            self.status_label.setText(f"Статус: {status}")
            # Устанавливаем цвет статуса (пример)
            if status == 'filled':
                self.status_label.setStyleSheet("color: green;")
            elif status == 'empty':
                self.status_label.setStyleSheet("color: orange;")
            elif status == 'invalid':
                self.status_label.setStyleSheet("color: red;")
            else:
                self.status_label.setStyleSheet("color: gray;")

            # Разблокируем сигналы
            self.value_edit.blockSignals(False)
            self.freeze_checkbox.blockSignals(False)

            # Управляем активностью поля ввода в зависимости от "заморозки"
            self.value_edit.setReadOnly(data.get('is_frozen', False))

            self.setVisible(True) # Показываем редактор
        else:
            # Если данных нет, очищаем и скрываем
            self.clear_editor()

    def get_edited_data(self) -> dict:
        """
        Возвращает текущие данные из виджетов редактора.
        """
        if not self._current_key_id:
            return {}

        # Определяем статус (упрощенно: заполнено или пусто)
        current_value = self.value_edit.toPlainText()
        # TODO: Добавить реальную валидацию для статуса 'invalid'
        current_status = 'filled' if current_value else 'empty'

        return {
            'value': current_value,
            'status': current_status,
            'is_frozen': self.freeze_checkbox.isChecked()
        }

    def get_current_key_id(self) -> str | None:
        """Возвращает ID ключа, который сейчас редактируется."""
        return self._current_key_id

    def clear_editor(self):
        """Очищает поля редактора и скрывает его."""
        self._current_key_id = None
        self.key_label.setText("Ключ: Не выбран")
        # Блокируем сигналы перед очисткой
        self.value_edit.blockSignals(True)
        self.freeze_checkbox.blockSignals(True)
        self.value_edit.clear()
        self.freeze_checkbox.setChecked(False)
        self.value_edit.blockSignals(False)
        self.freeze_checkbox.blockSignals(False)
        self.status_label.setText("Статус: -")
        self.status_label.setStyleSheet("color: gray;")
        self.value_edit.setReadOnly(False) # Снимаем блокировку при очистке
        self.setVisible(False) # Скрываем виджет