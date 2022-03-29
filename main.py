import requests
from openpyxl import Workbook
from datetime import datetime
import time
import telegram_send



class Parser():
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
        self.file_name_xlsx = f"Парсинг за {self.date_and_time}.xlsx"
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

    def fill_in_requirement(self, data_from_parsing):
        try:
            for row_number in range(len(data_from_parsing)):
                    self.filling_in_data(f'A{row_number + 2}', data_from_parsing[row_number]['name'])
                    if data_from_parsing[row_number]["salary"] is not None:
                        if data_from_parsing[row_number]["salary"]["to"] is not None:
                            self.filling_in_data(f'B{row_number + 2}', f'{data_from_parsing[row_number]["salary"]["from"]} - {data_from_parsing[row_number]["salary"]["to"]} {data_from_parsing[row_number]["salary"]["currency"]}')
                        else:
                            self.filling_in_data(f'B{row_number + 2}', f'{data_from_parsing[row_number]["salary"]["from"]} {data_from_parsing[row_number]["salary"]["currency"]}')
                    elif data_from_parsing[row_number]["salary"] is None:
                        self.filling_in_data(f'B{row_number + 2}', 'з/п не указана')
                    self.filling_in_data(f'C{row_number + 2}', data_from_parsing[row_number]['snippet']['requirement'])
                    self.filling_in_data(f'D{row_number + 2}', data_from_parsing[row_number]['snippet']['responsibility'])
                    self.filling_in_data(f'E{row_number + 2}', data_from_parsing[row_number]['alternate_url'])
        except:
            pass

    def save(self):
        self.wb.save(self.file_name_xlsx)
        return self.file_name_xlsx



    def close(self):
        self.wb.close()


class Telegram():
    def send_file(self, filename):
        self.data = {
            'caption': 'Результат парсинга',
            'chat_id': '946919713',
            'url': 'https://api.telegram.org/bot5239048826:AAHYeNnl1b5NaGlbSkqw84w75xUCVafkr_M/sendDocument'
        }
        with open(filename, 'rb') as f:
            files = {'document': f}
            requests.post(self.data['url'], data=self.data, files=files)




class Robot:
    def __init__(self, text):
        now = datetime.now()
        self.text = text
        self.date_and_time = now.strftime("%d.%m.%y %H_%M")

    def start(self):
        parser = Parser(self.text)
        data_from_parsing = parser.parsing()
        excel = Excel(self.date_and_time)
        excel.fill_in_requirement(data_from_parsing)
        file_name = excel.save()
        excel.close()
        telegram = Telegram()
        telegram.send_file(filename=file_name)


if __name__ == '__main__':
    andrew = Robot('Python Developer')
    andrew.start()


