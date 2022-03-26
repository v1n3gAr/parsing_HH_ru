import requests
from openpyxl import Workbook
from datetime import datetime
import time


class Parser:
    def __init__(self, text: str):
        self.text = text  # Запрос, который мы будем передавать для поиска на hh.ru
        self._area = 66  # Код города Нижний Новгород, все id городов находятся в документации к API
        self._per_page = 100  # Количество отображаемых результатов
        self._url = 'https://api.hh.ru/vacancies'

    def parsing(self):
        req = requests.get(self._url, params={
            'text': self.text,
            'area': self._area,
            'per_page': self._per_page
        }
                           )
        data = req.json()['items']
        req.close()
        return data


class Excel:
    def __init__(self, date_and_time):
        self.date_and_time = date_and_time  # Берем дату и время для их добавления к наименованию файла,
        # чтобы каждый новый файл был уникален
        self.wb = Workbook()
        self.ws = self.wb.active
        headings = {
            'A1': 'Наименование вакансии',
            'B1': 'Зарплата',
            'C1': 'Требования',
            'D1': 'Обязанности',
            'E1': 'Ссылка',
        }

        for i, j in headings.items():  # Заполняем шапку Excel
            self.ws[i].value = j

    def filling_in_data(self, cell, value):
        self.ws[cell].value = value

    def save(self):
        self.wb.save(f"Парсинг за {self.date_and_time}.xlsx")


    def close(self):
        self.wb.close()


class Robot:
    def __init__(self, text):
        now = datetime.now()
        self.text = text
        self.date_and_time = now.strftime("%d.%m.%y %H_%M")

    def start(self):

        parser = Parser(self.text)
        ps = parser.parsing()
        excel = Excel(self.date_and_time)
        for i in range(len(ps)):
            try:
                excel.filling_in_data(f'A{i + 2}', ps[i]['name'])
                if ps[i]["salary"] is not None:
                    if ps[i]["salary"]["to"] is not None:
                        excel.filling_in_data(f'B{i + 2}', f'{ps[i]["salary"]["from"]} - {ps[i]["salary"]["to"]} {ps[i]["salary"]["currency"]}')
                    else:
                        excel.filling_in_data(f'B{i + 2}', f'{ps[i]["salary"]["from"]} {ps[i]["salary"]["currency"]}')
                elif ps[i]["salary"] is None:
                    excel.filling_in_data(f'B{i + 2}', 'Не указано')
                excel.filling_in_data(f'C{i + 2}', ps[i]['snippet']['requirement'])
                excel.filling_in_data(f'D{i + 2}', ps[i]['snippet']['responsibility'])
                excel.filling_in_data(f'E{i + 2}', ps[i]['alternate_url'])
                time.sleep(0.5)
            except:
                pass
        excel.save()


if __name__ == '__main__':
    andrew = Robot('Python Developer')
    andrew.start()
