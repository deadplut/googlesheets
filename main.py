import binascii
import json
import gspread
import pyasn1


class ProgrammConnect:
    def __init__(self):
        self.path = ""
        self.table_name = ""
        self.sheet_name = ""
        self.sa = None
        self.work_sheet = None
        self.work_table = None

    def connect_account(self):

        """
            :return:
            Ф-ция осуществляет соединение с аккаунтом.
        """

        self.get_path()

        try:
            self.sa = gspread.service_account(filename=self.path)
            print('Подключение к аккаунту.\n')

        except FileNotFoundError:
            print("Путь или название файла некорректны.")
            self.connect_account()
        except (ValueError, json.decoder.JSONDecodeError, pyasn1.error.PyAsn1Error, binascii.Error):
            print("Некорректные данные, проверьте файл с ключом.")
            self.connect_account()

    def connect_table(self):

        """
            :return:
            Ф-ция осуществляет соединение с таблицей.
        """

        self.get_table_name()
        try:
            self.work_table = self.sa.open(self.table_name)
            print("Подключение к таблице.\n")
        except gspread.exceptions.SpreadsheetNotFound:
            print("Название таблицы некорректно.")
            self.connect_table()

    def connect_sheet(self):

        """
            :return:
            Ф-ция осуществляет соединение с листом таблицы.
        """

        self.get_sheet_name()
        try:
            self.work_sheet = self.work_table.worksheet(self.sheet_name)
            print("Подключение к листу.\n")
            print("Соединение установлено.")
        except gspread.exceptions.WorksheetNotFound:
            print("Название листа таблицы некорректно")
            self.connect_sheet()

    def get_path(self):
        """
            :return:
            Ф-ция осуществляет получение названия файли или получает путь.
        """
        self.path = str(input("Укажите название файла с ключом или путь к файлу:"))

    def get_table_name(self):

        """
            :return:
            Ф-ция осуществляет получение названия таблицы.
        """

        self.table_name = str(input("Укажите название таблицы:"))

    def get_sheet_name(self):

        """
            :return:
            Ф-ция осуществляет получение названия таблицы.
         """

        self.sheet_name = str(input("Укажите название листа таблицы:"))


class Excel:

    def __init__(self, work_sheet):
        self.work_sheet = work_sheet
        self.method = None

    def operation(self):

        """
            :return:
            Ф-ция предоставляет выбор метода: вставить, удалить, изменить.
            Также получает номер ячейки и значение.
        """


        work_status = True

        while work_status:
            print("Какое действие вы хотите совершить?")
            print("1 - вставить")
            print("2 - удалить")
            print("3 - изменить")

            self.method = str(input("Введите цифру действия: "))

            if self.method not in ("1", "2", "3"):
                print("Некорректный ввод.")
                print("----------------------------------------")
                continue
            cell = str(input("Введите номер ячейки: "))
            if self.method == "2":
                self.insert(self.work_sheet, cell, value="")
                print("Значение удалено.")
            else:
                value = str(input("Введите значение ячейки: "))
                self.insert(self.work_sheet, cell, value)
                print("Значение обновлено.")

            active_task = None
            while active_task not in ("y", "n"):
                active_task = input("Есть еще изменения ячеек? y/n: ")
                if active_task == "y":
                    break
                elif active_task == "n":
                    work_status = False
                    print("Завершение работы.")
                    break
                print("Некорректное значение!")

    def insert(self, ws, cell, value):

        """
            :return:
            Ф-ция осуществляет обновление значения ячейки.
        """

        return ws.update(cell, value)


"""
аккаунт = service_account.json
таблица = table_11
лист = Лист1
"""

if __name__ == '__main__':
    app = ProgrammConnect()
    app.connect_account()
    app.connect_table()
    app.connect_sheet()
    sheet = app.work_sheet
    excel = Excel(sheet)
    excel.operation()

