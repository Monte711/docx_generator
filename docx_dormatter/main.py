import sys
from PySide6.QtWidgets import QApplication
from views.main_window import MainWindow # Импортируем класс нашего окна

if __name__ == '__main__':
    # Создаем экземпляр приложения
    app = QApplication(sys.argv)

    # Создаем и показываем главное окно
    main_window = MainWindow()
    main_window.show()

    # Запускаем главный цикл обработки событий
    sys.exit(app.exec())