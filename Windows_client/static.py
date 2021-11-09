import os
from os import path
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtGui import QRegExpValidator


# Открытие файла README
def open_readme():
    os.system(f'start {path.join(path.dirname(__file__), "README.docx")}')


# Блокирование редактирования формы + изменение цвета
def lock_form(form):
    form.setReadOnly(True)
    form.setStyleSheet("""
                       background-color: #DCDCDC;
                       """)


# Вставка текста по приложению
def set_text(form, text):
    form.setText(text)


# Валидатор для excel_form (не используется)
def validator(form, form_range):
    serial_regex = QRegExp("^" + form_range)
    serial_validator = QRegExpValidator(serial_regex)
    form.setValidator(serial_validator)


# Стиль для заголовка
def set_title_font(form):
    form.setAlignment(Qt.AlignCenter)
    form.setStyleSheet("""
                       font-weight: 900;
                       """
                       )
