# Примерный тестовый скрипт (запускать из папки docx_formatter)
from models.project import Project
import os

# 1. Создаем проект
p = Project()
print("Проект создан.")

# 2. Добавляем шаблон (замените на реальный путь к вашему ТП++.docx)
template_file = "E:/docx_dormatter/ТП++.docx" # <-- УКАЖИТЕ ПРАВИЛЬНЫЙ ПУТЬ!
if os.path.exists(template_file):
    p.add_template(template_file)
else:
    print(f"ПРЕДУПРЕЖДЕНИЕ: Файл шаблона не найден: {template_file}")

# 3. Устанавливаем путь вывода
output_dir = "generated_docs"
p.set_output_path(output_dir) # Папка будет создана, если ее нет

# 4. Добавляем тестовые данные ключей
p.update_key_data("{{ORG_NAME}}", {"value": "Тестовая Организация", "status": "filled"})
p.update_key_data("HardwareList", {"type": "dynamic_table", "data": [{"name": "PC-01"}]})
print("Тестовые данные добавлены.")
print("Все данные ключей:", p.get_all_keys_data())

# 5. Сохраняем проект
save_file = "my_test_project.dfp"
if p.save(save_file):
    print(f"Проект сохранен в {save_file}")

    # 6. Создаем новый пустой проект и загружаем сохраненный
    p2 = Project()
    if p2.load(save_file):
        print("Проект успешно загружен.")
        print("Загруженные шаблоны:", p2.template_paths)
        print("Загруженный путь вывода:", p2.output_path)
        print("Загруженные ключи:", p2.keys_data)
        print("Имя файла проекта:", p2.get_project_filename())
else:
    print("Не удалось сохранить проект.")

# Очистка (опционально)
# if os.path.exists(save_file):
#     os.remove(save_file)
# if os.path.exists(output_dir):
#     try:
#         os.rmdir(output_dir) # Удалит, только если папка пуста
#     except OSError:
#         pass