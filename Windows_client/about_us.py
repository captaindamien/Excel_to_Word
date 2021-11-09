from os import path
from PyQt5.QtCore import Qt
from PyQt5 import QtWidgets, QtGui
from ui.about_us import Ui_AboutUsWindow
from static import set_text


class AboutUsWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(AboutUsWindow, self).__init__()
        self.setFixedSize(541, 149)
        # Инициализация окна
        self.ui_3 = Ui_AboutUsWindow()
        self.ui_3.setupUi(self)
        # Путь до конфига
        self.config_dir = path.join(path.dirname(__file__), 'config')
        # Текст по окну
        self.windows_text()

    # Оформление окна
    def windows_text(self):
        # Оглавление окна
        self.setWindowTitle('О проекте')
        # Логотип оглавления
        self.setWindowIcon(QtGui.QIcon(path.join(self.config_dir, 'logo.ico')))
        # Текст по окну
        set_text(self.ui_3.label, 'ООО "Региональные энергетические системы"')
        self.ui_3.label.setAlignment(Qt.AlignCenter)
        self.ui_3.label.setStyleSheet("""
                                      font-weight: 900;
                                      """)
        set_text(self.ui_3.label_2, 'Проект создан для автоматизации труда и облегчения рутинных задач связанных '
                                    'с MsOffice.\nДля предложений по улучшению проекта можете написать на почту - '
                                    'karpushin@gkres.ru')
        self.ui_3.label_2.setAlignment(Qt.AlignCenter)
        set_text(self.ui_3.label_3, 'Разработчик - Карпушин Егор Юрьевич')
        self.ui_3.label_3.setAlignment(Qt.AlignCenter)