import configparser
from os import path
from PyQt5 import QtWidgets, QtGui
from ui.config_window import Ui_ConfigWindow
from static import set_text


class ConfigWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(ConfigWindow, self).__init__()
        self.setFixedSize(354, 558)
        # Инициализация окна
        self.ui_2 = Ui_ConfigWindow()
        self.ui_2.setupUi(self)
        # Путь до конфига
        self.config_dir = path.join(path.dirname(__file__), 'config')
        # Открытие файла конфига
        self.config = configparser.RawConfigParser()
        self.config.read(path.join(self.config_dir, 'config.ini'))
        # Метод на чтение конфига
        self.read_config()
        # Подключение кнопок
        self.init_handlers()
        # Текст по окну
        self.windows_text()

    # Обработка нажатия для октрытия сторонних окон
    def init_handlers(self):
        self.ui_2.pushButton.clicked.connect(self.save_config)
        self.ui_2.pushButton_2.clicked.connect(self.close_config)

    def windows_text(self):
        # Оглавление окна
        self.setWindowTitle('Окно изменения конфигураций')
        # Иконка для приложения
        self.setWindowIcon(QtGui.QIcon(path.join(self.config_dir, 'logo.ico')))
        # Текст по окну
        set_text(self.ui_2.label, 'Впишите нужные аргументы')
        set_text(self.ui_2.pushButton, 'Подтвердить')
        set_text(self.ui_2.pushButton_2, 'Отмена')

    # Метод на чтение конфига
    def read_config(self):
        with open(path.join(self.config_dir, 'config.ini')) as config:
            all_info = config.read()
            set_text(self.ui_2.textEdit, all_info)

    # Метод на сохранение конфига по нажатию кнопки
    def save_config(self):
        with open(path.join(self.config_dir, 'config.ini'), 'w') as config:
            to_config = self.ui_2.textEdit.toPlainText()
            config.write(to_config)
        # Закрывает окно после сохранения
        self.close_config()

    # Метод на закрытие конфига
    def close_config(self):
        self.ui_2.textEdit.clear()
        # Чтение конфига нужно, чтобы при последующем открытии обновлялась информация
        self.read_config()
        self.hide()
