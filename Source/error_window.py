from os import path
from PyQt5 import QtWidgets, QtGui
from ui.error_window import Ui_ErrorWindow
from static import set_text, set_title_font


class ErrorWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(ErrorWindow, self).__init__()
        # Инициализация окна
        self.ui_4 = Ui_ErrorWindow()
        self.ui_4.setupUi(self)
        # Путь до конфига
        self.config_dir = path.join(path.dirname(__file__), 'config')
        # Текст по окну
        self.windows_text()

    # Оформление окна
    def windows_text(self):
        # Оглавление окна
        self.setWindowTitle('ООО "РЭС"')
        # Логотип оглавления
        self.setWindowIcon(QtGui.QIcon(path.join(self.config_dir, 'logo.ico')))
        set_text(self.ui_4.label, '')
        set_title_font(self.ui_4.label)
