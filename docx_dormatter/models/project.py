import json
import os
from pathlib import Path
import copy

class Project:
    """
    Класс для хранения и управления данными проекта.
    (Версия 5: add_found_table принимает template_keys)
    """
    def __init__(self):
        self.filepath: Path | None = None
        self.template_paths: list[Path] = []
        self.output_path: Path | None = None
        self.keys_data: dict = {} # { key_id: { data_dict } }
        self.is_modified: bool = False

    # ... (reset, add_template, remove_template, set_output_path, add_found_key без изменений) ...
    def reset(self):
        self.__init__()

    def add_template(self, path_str: str) -> bool:
        path = Path(path_str)
        if path.is_file() and path.suffix.lower() == '.docx':
            if path not in self.template_paths:
                self.template_paths.append(path)
                self.is_modified = True
                print(f"Шаблон добавлен: {path}")
                return True
        print(f"Ошибка: Неверный путь к шаблону или не DOCX файл: {path_str}")
        return False

    def remove_template(self, path_str: str) -> bool:
        path_to_remove = Path(path_str)
        if path_to_remove in self.template_paths:
            self.template_paths.remove(path_to_remove)
            self.is_modified = True
            print(f"Шаблон удален: {path_str}")
            return True
        return False

    def set_output_path(self, path_str: str) -> bool:
        path = Path(path_str)
        if path.is_dir():
            self.output_path = path
            self.is_modified = True
            print(f"Путь вывода установлен: {path}")
            return True
        else:
            try:
                path.mkdir(parents=True, exist_ok=True)
                self.output_path = path
                self.is_modified = True
                print(f"Путь вывода создан и установлен: {path}")
                return True
            except OSError as e:
                print(f"Ошибка установки пути вывода: {e}")
                return False


    def add_found_key(self, key_name: str):
        if key_name not in self.keys_data:
            self.keys_data[key_name] = {
                'value': '',
                'status': 'empty',
                'is_frozen': False,
            }
            self.is_modified = True
            print(f"Добавлен новый ключ: {key_name}")


    # --- Обновленный метод ---
    def add_found_table(self, table_id: str, template_keys: list[str] | None = None):
        """
        Добавляет найденный ID динамической таблицы в keys_data, если его еще нет,
        сохраняя связанные template_keys.
        """
        if table_id not in self.keys_data:
            self.keys_data[table_id] = {
                'type': 'dynamic_table',
                'columns': [], # Заголовки столбцов (можно будет заполнить позже)
                'template_keys': template_keys if template_keys else [], # Сохраняем ключи!
                'data': []
            }
            self.is_modified = True
            print(f"Добавлена новая таблица: {table_id} с template_keys: {template_keys}")
        # else: # Если таблица уже есть, может быть, обновить template_keys?
            # current_keys = self.keys_data[table_id].get('template_keys', [])
            # if template_keys and current_keys != template_keys:
            #     print(f"Предупреждение: Обновление template_keys для таблицы {table_id}")
            #     self.keys_data[table_id]['template_keys'] = template_keys
            #     self.is_modified = True
            # pass # Решаем, нужно ли обновлять ключи, если таблица уже существует

    # ... (update_key_data, get_key_data, get_all_keys_data, set_keys_data,
    #      get_project_filename, save, load без изменений) ...
    def update_key_data(self, key_id: str, data: dict):
        if key_id not in self.keys_data:
            self.keys_data[key_id] = data
            self.is_modified = True
            print(f"Элемент '{key_id}' добавлен с новыми данными.")
            return
        old_data = self.keys_data[key_id]
        data_changed = False
        if old_data.get('type') == 'dynamic_table':
            old_table_rows = old_data.get('data', [])
            new_table_rows = data.get('data', [])
            if old_table_rows != new_table_rows:
                old_data['data'] = copy.deepcopy(new_table_rows)
                data_changed = True
        else:
            if (old_data.get('value') != data.get('value') or
                    old_data.get('is_frozen') != data.get('is_frozen')):
                old_data['value'] = data.get('value', '')
                old_data['status'] = data.get('status', 'unknown')
                old_data['is_frozen'] = data.get('is_frozen', False)
                data_changed = True
        if data_changed:
            self.is_modified = True
            print(f"Данные элемента '{key_id}' обновлены в модели.")

    def get_key_data(self, key_id: str) -> dict | None:
        return self.keys_data.get(key_id)

    def get_all_keys_data(self) -> dict:
        return self.keys_data

    def set_keys_data(self, new_keys_data: dict):
        self.keys_data = new_keys_data

    def get_project_filename(self) -> str:
        if self.filepath: return self.filepath.name
        return "Безымянный"

    def save(self, path_str: str | None = None) -> bool:
        save_path = Path(path_str) if path_str else self.filepath
        if not save_path: print("Ошибка сохранения: Путь не указан."); return False
        if save_path.suffix.lower() != '.dfp': save_path = save_path.with_suffix('.dfp')
        project_data = {
            "version": "1.0",
            "template_paths": [str(p) for p in self.template_paths],
            "output_path": str(self.output_path) if self.output_path else None,
            "keys_data": self.keys_data
        }
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=4)
            self.filepath = save_path
            self.is_modified = False
            print(f"Проект успешно сохранен в: {save_path}")
            return True
        except Exception as e:
            print(f"Ошибка сохранения файла проекта '{save_path}': {e}")
            return False

    def load(self, path_str: str) -> bool:
        load_path = Path(path_str)
        if not load_path.is_file(): print(f"Ошибка загрузки: Файл не найден - {load_path}"); return False
        try:
            with open(load_path, 'r', encoding='utf-8') as f: project_data = json.load(f)
            if not all(k in project_data for k in ["template_paths", "output_path", "keys_data"]): raise ValueError("Неверный формат файла проекта.")
            self.reset()
            self.template_paths = [Path(p) for p in project_data.get("template_paths", [])]
            output_p_str = project_data.get("output_path")
            self.output_path = Path(output_p_str) if output_p_str else None
            self.keys_data = project_data.get("keys_data", {})
            self.filepath = load_path
            self.is_modified = False
            print(f"Проект успешно загружен из: {load_path}")
            return True
        except Exception as e:
            print(f"Ошибка загрузки или обработки файла проекта '{load_path}': {e}")
            self.reset(); return False