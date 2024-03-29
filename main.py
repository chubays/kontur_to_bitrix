# This is a sample Python script.
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import traceback
from PyQt6 import QtGui
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QMainWindow, QLabel,
                             QFileDialog, QApplication, QPushButton, QWidget, QVBoxLayout, QComboBox)
import sys

df = pd.DataFrame()
correct_df = pd.DataFrame()
FILE_NAME = ""
VERSION = 'Конвертация файла v12102023'

regions = {
    '01': '01|Республика Адыгея (Адыгея)',
    '02': '02|Республика Башкортостан',
    '03': '03|Республика Бурятия',
    '04': '04|Республика Алтай',
    '05': '05|Республика Дагестан',
    '06': '06|Республика Ингушетия',
    '07': '07|Кабардино-Балкарская Республика',
    '08': '08|Республика Калмыкия',
    '09': '09|Карачаево-Черкесская Республика',
    '10': '10|Республика Карелия',
    '11': '11|Республика Коми',
    '12': '12|Республика Марий Эл',
    '13': '13|Республика Мордовия',
    '14': '14|Республика Саха (Якутия)',
    '15': '15|Республика Северная Осетия - Алания',
    '16': '16|Республика Татарстан (Татарстан)',
    '17': '17|Республика Тыва',
    '18': '18|Удмуртская Республика',
    '19': '19|Республика Хакасия',
    '20': '20|Чеченская Республика',
    '21': '21|Чувашская Республика - Чувашия',
    '22': '22|Алтайский край',
    '23': '23|Краснодарский край',
    '24': '24|Красноярский край',
    '25': '25|Приморский край',
    '26': '26|Ставропольский край',
    '27': '27|Хабаровский край',
    '28': '28|Амурская область',
    '29': '29|Архангельская область',
    '30': '30|Астраханская область',
    '31': '31|Белгородская область',
    '32': '32|Брянская область',
    '33': '33|Владимирская область',
    '34': '34|Волгоградская область',
    '35': '35|Вологодская область',
    '36': '36|Воронежская область',
    '37': '37|Ивановская область',
    '38': '38|Иркутская область',
    '39': '39|Калининградская область',
    '40': '40|Калужская область',
    '41': '41|Камчатский край',
    '42': '42|Кемеровская область - Кузбасс',
    '43': '43|Кировская область',
    '44': '44|Костромская область',
    '45': '45|Курганская область',
    '46': '46|Курская область',
    '47': '47|Ленинградская область',
    '48': '48|Липецкая область',
    '49': '49|Магаданская область',
    '50': '50|Московская область',
    '51': '51|Мурманская область',
    '52': '52|Нижегородская область',
    '53': '53|Новгородская область',
    '54': '54|Новосибирская область',
    '55': '55|Омская область',
    '56': '56|Оренбургская область',
    '57': '57|Орловская область',
    '58': '58|Пензенская область',
    '59': '59|Пермский край',
    '60': '60|Псковская область',
    '61': '61|Ростовская область',
    '62': '62|Рязанская область',
    '63': '63|Самарская область',
    '64': '64|Саратовская область',
    '65': '65|Сахалинская область',
    '66': '66|Свердловская область',
    '67': '67|Смоленская область',
    '68': '68|Тамбовская область',
    '69': '69|Тверская область',
    '70': '70|Томская область',
    '71': '71|Тульская область',
    '72': '72|Тюменская область',
    '73': '73|Ульяновская область',
    '74': '74|Челябинская область',
    '75': '75|Забайкальский край',
    '76': '76|Ярославская область',
    '77': '77|город федерального значения Москва',
    '78': '78|город федерального значения Санкт-Петербург',
    '79': '79|Еврейская автономная область',
    '83': '83|Ненецкий автономный округ',
    '86': '86|Ханты-Мансийский автономный округ - Югра',
    '87': '87|Чукотский автономный округ',
    '89': '89|Ямало-Ненецкий автономный округ',
    '90': '90|Запорожская область',
    '91': '91|Республика Крым',
    '92': '92|город федерального значения Севастополь',
    '93': '93|Донецкая Народная Республика',
    '94': '94|Луганская Народная Республика',
    '95': '95|Херсонская область',
    '99': '99|Иные территории,включая город и космодром Байконур'}


def fill_region(row):
    if pd.notna(row['Реквизит (Россия): КПП']):
        kpp = str(row['Реквизит (Россия): КПП'])[:2]
        return regions[kpp]
    else:
        inn = str(row['Реквизит (Россия): ИНН'])[:2]
        return regions[inn]


def convert(sfera_):
    try:
        correct_df['Название компании'] = df['Реквизит (Россия): Сокращенное наименование организации']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): ИНН'] = df['Реквизит (Россия): ИНН']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): КПП'] = df['Реквизит (Россия): КПП']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): ОГРН'] = df['Реквизит (Россия): ОГРН']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): ОГРНИП'] = df['Реквизит (Россия): ОГРНИП']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Должность руководителя'] = df['Должность']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Рабочий телефон'] = df['Рабочий телефон']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Рабочий e-mail'] = df['Рабочий e-mail']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): Адрес'] = df['Реквизит (Россия): Адрес']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): Адрес - улица, номер дома'] = df['Реквизит (Россия): Адрес - улица, номер дома']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): Адрес - населенный пункт'] = df['Реквизит (Россия): Адрес - населенный пункт']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): Адрес - район'] = df['Реквизит (Россия): Адрес - район']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): Адрес - регион'] = df['Реквизит (Россия): Адрес - регион']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): Адрес - почтовый индекс'] = df['Реквизит (Россия): Адрес - почтовый индекс']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит (Россия): Адрес - страна'] = df['Реквизит (Россия): Адрес - страна']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Корпоративный сайт'] = df['Корпоративный сайт']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Источник'] = df['Источник']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Комментарий'] = df['Источник']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    # Надо проверять - если ИНН длинный - то это ИП, иначе Организация
    correct_df['Реквизит: Шаблон'] = 'Организация'
    try:
        correct_df['Реквизит (Россия): Адрес - тип'] = df['Реквизит (Россия): Адрес - тип']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Реквизит: Название'] = df['Реквизит (Россия): Сокращенное наименование организации']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Статус ЮЛ'] = df['Статус организации']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Годовой оборот'] = df['Выручка']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Основной ОКВЭД'] = df['Основной вид деятельности']
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Дополнительные ОКВЭДы'] = df['Другие виды деятельности'].str.replace(';', ',')
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    # Надо запросить сферу деятельности!!!
    try:
        correct_df['Сфера деятельности'] = sfera_
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        df['Количество сотрудников'] = df['Количество сотрудников'].fillna(0)
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Численность работников'] = df['Количество сотрудников'].astype(int)
        correct_df['Кол-во сотрудников'] = df['Количество сотрудников'].astype(int)
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    try:
        correct_df['Лицензия'] = df['Полученные лицензии'].str.replace(';', ',')
    except Exception as e:
        print('Ошибка:\n', e)
        print('Ошибка:\n', traceback.format_exc())
    correct_df['Дата регистрации компании'] = df['Реквизит (Россия): Дата государственной регистрации']
    correct_df['Регион клиента'] = correct_df.apply(lambda row: fill_region(row), axis=1)
    #    correct_df['Полученные лицензии'].replace(';', ',', inplace=True)


#    if create_deal:
#        set_create_deal()


# def set_create_deal():
#    correct_df['Создать сделку по новой воронке'] = 'Да'


class Example(QMainWindow):

    def __init__(self, spheres):
        super().__init__()

        self.generalLayout = QVBoxLayout()
        central_widget = QWidget(self)
        central_widget.setLayout(self.generalLayout)
        self.setCentralWidget(central_widget)

        self.btn = QPushButton(self)
        self.btn.setText("Выбрать файл")
        # self.setCentralWidget(btn)
        self.btn.clicked.connect(lambda: self.open_dialog())
        # self.btn.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.generalLayout.addWidget(self.btn)

        self.text_label = QLabel()
        self.text_label.setText('Укажите сферу деятельности')
        self.text_label.setFixedHeight(30)
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.generalLayout.addWidget(self.text_label)

        self.sf = QComboBox()
        self.sf.addItem('')
        self.sf.addItems([sphere for sphere in spheres])
        self.generalLayout.addWidget(self.sf)

        # self.make_deal = QCheckBox()
        # self.make_deal.setText("Создать сделку в новой воронке?")
        # self.make_deal.setAlignment(Qt.AlignmentFlag.AlignLeft)
        # self.generalLayout.addWidget(self.make_deal)

        self.btn2 = QPushButton(self)
        self.btn2.setText("Обработать файл")
        self.btn2.setDisabled(True)
        # self.setCentralWidget(btn)
        self.btn2.clicked.connect(lambda: self.convert_file())
        # self.btn.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.generalLayout.addWidget(self.btn2)

        self.status_label = QLabel()
        self.status_label.setFixedHeight(50)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.status_label.setTextFormat(Qt.TextFormat.RichText)
        self.status_label.setFont(QtGui.QFont("Times", 11))
        self.generalLayout.addWidget(self.status_label)

        self.setGeometry(200, 200, 320, 250)
        self.setWindowTitle(VERSION)
        self.show()

    def convert_file(self):
        # if self.sfera.text():
        if self.sf.currentIndex():
            convert(self.sf.currentText())  # self.make_deal.isChecked())
            self.status_label.setText("Файл обработан")
            s_fname = QFileDialog.getSaveFileName(
                self,
                "Сохранить файл",
                self.sf.currentText() + '.csv',
                "CSV (*.csv)"
            )
            if s_fname[0]:
                try:
                    correct_df.to_csv(s_fname[0], sep=';', index=False)
                    self.status_label.setText("Файл обработан и сохранён")
                except Exception as e:
                    print('Ошибка:\n', e)
                    print('Ошибка:\n', traceback.format_exc())
                    self.status_label.setText("Ошибка при сохранении файла!")
            else:
                self.status_label.setText("Файл не сохранён!")
        else:
            self.status_label.setText("Заполните сферу деятельности компаний!")

    def open_dialog(self):
        fname = QFileDialog.getOpenFileName(
            self,
            "Open File",
            "",
            "CSV (*.csv);; All Files (*)"
        )
        if fname[0]:
            global df
            try:
                df = pd.read_csv(fname[0],
                                 sep=';',
                                 encoding='utf-8',
                                 dtype={'Реквизит (Россия): ИНН': str, 'Реквизит (Россия): КПП': str})
                self.btn2.setDisabled(False)
            except Exception as ex:
                # self.status_label.setText("Ошибка при считывании файла")
                print('Ошибка:\n', ex)
                print('Ошибка:\n', traceback.format_exc())
                try:
                    df = pd.read_csv(fname[0], sep=';', encoding='windows-1251')
                    self.btn2.setDisabled(False)
                except Exception as ex1:
                    self.status_label.setText("Ошибка при считывании файла")
                    print('Ошибка:\n', ex1)
                    print('Ошибка:\n', traceback.format_exc())


def main():
    app = QApplication(sys.argv)
    with open('sphere.txt', 'r', encoding='UTF-8') as f:
        nums = f.read().splitlines()
    nums.sort()
    Example(nums)
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
