import datetime
import re
import psycopg2, xlrd

datas = ['/Users/jklyuev/Desktop/BDA/2020/01-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/02-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/03-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/04-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/05-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/06-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/07-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/08-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/09-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/10-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/11-2020.xls',
         '/Users/jklyuev/Desktop/BDA/2020/12-2020.xls', ]

t_date = datetime.date.today()  # дата V
t_number = 0  # номер V
t_age = 0  # возраст V
t_who = ''  # кто вызвал V | больной - 0; родственник - 1; МЧС - 2; очевидец - 3; соседи - 4; ССМП - 5
t_from = ''  # адрес
t_reason = ''  # повод
t_type = ''  # вызов (первичный/попутный)
t_state = ''  # вид
t_diag = ''  # диагноз
t_result = ''  # результат
t_to = ''  # доставлен
t_station = ''  # подстанция  V
t_time1 = 0  # время принятия V
t_time2 = 0  # время приезда V


def clearRow():
    t_date = 0  # дата
    t_number = 0  # номер
    t_age = 0  # возраст
    t_who = ''  # кто вызвал
    t_from = ''  # адрес
    t_reason = ''  # повод
    t_type = ''  # вызов (первичный/вторичный)
    t_state = ''  # вид
    t_diag = ''  # диагноз
    t_result = ''  # результат
    t_to = ''  # доставлен
    t_station = ''  # подстанция
    t_time1 = 0  # время принятия
    t_time2 = 0  # время приезда


def printRow():
    print(t_date, t_number, t_age, t_who, t_from, t_reason, t_type, t_state, t_diag, t_result, t_to, t_station, t_time1,
          t_time2)


conn = psycopg2.connect(
    host="localhost",
    database="amb",
    user="postgres",
    password="root")

if conn is not None:
    conn.close()
    print('Database connection closed.')

workbook = xlrd.open_workbook('/Users/jklyuev/Desktop/BDA/2020/01-2020.xls')
worksheet = workbook.sheet_by_index(0)

for i in range(worksheet.nrows):
    for j in range(20):
        if worksheet.cell(i, j).ctype == 1 and "Испол." in worksheet.cell(i, j).value:
            printRow()
            clearRow()
        elif worksheet.cell(i, j).ctype == 3: #дата, время выезда и прибытия
            ms_date_number = worksheet.cell(i, j).value
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number, workbook.datemode)
            py_date = datetime.datetime(year, month, day, hour, minute, second)
            if py_date.year == 1970:
                if t_time1 == 0:
                    t_time1 = py_date
                    t_time2 = 0
                else:
                    t_time2 = py_date
            elif py_date.year == 2020 or py_date.year == 2021 or py_date.year == 2022:
                t_date = py_date
                t_time1 = 0
                t_time2 = 0
        elif worksheet.cell(i, j).ctype == 1 and "Возраст:" in worksheet.cell(i, j).value: #возраст и кто вызвал
            t_age = worksheet.cell(i, j + 2).value

            if 'лет' in worksheet.cell(i, j + 2).value:
                t_age = int(re.findall('[0-9]+', worksheet.cell(i, j + 2).value)[0])
            else:
                t_age = int(re.findall('[0-9]+', worksheet.cell(i, j + 2).value)[0])/100

            if 'родственник' in worksheet.cell(i, j + 5).value:
                t_who = 1
            elif 'МЧС' in worksheet.cell(i, j + 5).value:
                t_who = 2
            elif 'очевидец' in worksheet.cell(i, j + 5).value:
                t_who = 3
            elif 'сосед' in worksheet.cell(i, j + 5).value:
                t_who = 4
            elif 'ССМП' in worksheet.cell(i, j + 5).value:
                t_who = 5
            else:
                t_who = 0
        elif worksheet.cell(i, j).ctype == 1 and "Номер:" in worksheet.cell(i, j).value: # номер вызова
            t_number = worksheet.cell(i, j + 1).value
        elif worksheet.cell(i, j).ctype == 1 and "Подстанция:" in worksheet.cell(i, j).value:  # подстанция
            t_station = worksheet.cell(i, j + 5).value
        elif worksheet.cell(i, j).ctype == 1 and "Вызов:" in worksheet.cell(i, j).value:  # какой вызов
            t_type = worksheet.cell(i, j + 2).value