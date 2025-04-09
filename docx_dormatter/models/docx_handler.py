import re
from pathlib import Path
import docx # type: ignore # pip install python-docx
# lxml не используется напрямую в find_keys, но понадобится позже для замены
# import lxml

class DocxHandler:
    """
    Класс для инкапсуляции операций с файлами DOCX.
    """

    # Регулярное выражение для поиска ключей вида {{KEY_NAME}}
    # Оно ищет две открывающие скобки, затем любые символы (.*?),
    # лениво (*?), до первых двух закрывающих скобок.
    KEY_PATTERN = re.compile(r"\{\{.*?\}\}")

    # Регулярное выражение для поиска маркеров динамических таблиц
    # Оно ищет {{DYNAMIC_TABLE::, затем захватывает (\w+) одно или больше
    # буквенно-цифровых символов (ID таблицы), и затем }}
    DYNAMIC_TABLE_PATTERN = re.compile(r"\{\{DYNAMIC_TABLE::(\w+)\}\}")

    def find_keys_in_template(self, docx_path: Path) -> dict[str, set]:
        """
        Находит все уникальные ключи {{...}} и маркеры {{DYNAMIC_TABLE::...}}
        в указанном DOCX файле (в параграфах и таблицах).

        Args:
            docx_path: Путь к файлу DOCX.

        Returns:
            Словарь с двумя ключами:
            'keys': set уникальных обычных ключей (включая маркеры таблиц целиком).
            'dynamic_tables': set уникальных ID динамических таблиц (из маркеров).
            Возвращает пустые множества в случае ошибки.
        """
        found_keys = set()
        dynamic_table_ids = set()

        try:
            document = docx.Document(docx_path)

            # 1. Поиск в параграфах основного текста
            for para in document.paragraphs:
                # Собираем текст из всех 'run' внутри параграфа,
                # т.к. ключ может быть разбит форматированием на несколько run.
                full_para_text = "".join(run.text for run in para.runs)
                keys_in_para = self.KEY_PATTERN.findall(full_para_text)
                found_keys.update(keys_in_para)

                # Ищем маркеры таблиц в параграфах (хотя обычно они в таблицах)
                table_markers = self.DYNAMIC_TABLE_PATTERN.findall(full_para_text)
                dynamic_table_ids.update(table_markers)


            # 2. Поиск в таблицах
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        # Обрабатываем каждый параграф внутри ячейки
                        for para in cell.paragraphs:
                            full_cell_para_text = "".join(run.text for run in para.runs)
                            keys_in_cell = self.KEY_PATTERN.findall(full_cell_para_text)
                            found_keys.update(keys_in_cell)

                            # Ищем маркеры таблиц
                            table_markers = self.DYNAMIC_TABLE_PATTERN.findall(full_cell_para_text)
                            dynamic_table_ids.update(table_markers)

            # TODO: Добавить поиск в колонтитулах (headers/footers), если необходимо
            # for section in document.sections:
            #     for header_para in section.header.paragraphs: ...
            #     for footer_para in section.footer.paragraphs: ...

        except Exception as e:
            print(f"Ошибка при чтении или обработке файла {docx_path}: {e}")
            # Возвращаем пустые множества в случае любой ошибки
            return {'keys': set(), 'dynamic_tables': set()}

        # Удаляем полные маркеры таблиц из основного списка ключей,
        # оставляя там только "обычные" ключи.
        # Это необязательно, но может быть удобнее.
        table_markers_full = {f"{{{{DYNAMIC_TABLE::{tid}}}}}" for tid in dynamic_table_ids}
        keys_only = found_keys - table_markers_full

        return {
            'keys': keys_only,
            'dynamic_tables': dynamic_table_ids
        }

# Пример использования (для теста)
if __name__ == '__main__':
    handler = DocxHandler()
    # Замените на путь к вашему файлу ТП++.docx с маркерами таблиц
    test_file = Path("ПУТЬ_К_ВАШЕМУ/ТП++.docx")
    if test_file.exists():
        results = handler.find_keys_in_template(test_file)
        print("Найденные обычные ключи:")
        print(results.get('keys', set()))
        print("\nНайденные ID динамических таблиц:")
        print(results.get('dynamic_tables', set()))
    else:
        print(f"Тестовый файл не найден: {test_file}")
