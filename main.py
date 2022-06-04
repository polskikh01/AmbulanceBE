import datetime
import re
import psycopg2, xlrd

dataFolders = ['/Users/jklyuev/Desktop/BDA/2020/01-2020.xls',
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
t_type = ''  # вызов (первичный - 0 /попутный - 1/ повторный - 2) V
t_state = ''  # вид (несчастный - 0/внезапное - 1 /внезапное - 2) V
t_diag = ''  # диагноз
t_result = ''  # результат (помощь на месте - 0 ; отказ - 1 ; смерть до приезда - 2 ; в присутствии - 3 ; здоров - 4 ; в больницу - 5 ; до прибытия - 6 ; травмпункт - 7)
t_to = ''  # доставлен
t_station = ''  # подстанция  V
t_time1 = 0  # время принятия V
t_time2 = 0  # время приезда V


"""
Городская клиническая больница - ГКБ
Городская поликлиника - ГП
Городская больница - ГБ
Детская городская больница - ДГБ
Детская городская поликлиника - ДГП
Детская городская клиническая больница - ДГКБ
"""

def clearRow():
    t_date = None  # дата
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
    t_time1 = None  # время принятия
    t_time2 = None  # время приезда


def printRow():
    print(t_date, t_number, t_age, t_who, t_from, t_reason, t_type, t_state, t_diag, t_result, t_to, t_station, t_time1,
          t_time2)

conn = psycopg2.connect(
    host="localhost",
    database="amb",
    user="postgres",
    password="root")

""" clearing db """
cursor = conn.cursor()
cursor.execute("TRUNCATE TABLE datas RESTART IDENTITY;")
conn.commit()
cursor.close()
""" ----------------- """

for file in dataFolders:
    workbook = xlrd.open_workbook(file)
    worksheet = workbook.sheet_by_index(0)

    for i in range(worksheet.nrows):
        for j in range(20):
            if worksheet.cell(i, j).ctype == 1 and "Испол." in worksheet.cell(i, j).value:
                #printRow()

                """ inserting into db """
                cursor = conn.cursor()
                cursor.execute("INSERT INTO datas(t_date, t_number, t_age, t_who, t_from, t_reason, t_type, t_state, t_diag, t_result, t_to, t_station, t_time1, t_time2) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);",
                               (t_date, t_number, t_age, t_who, t_from, t_reason, t_type, t_state, t_diag, t_result, t_to, t_station, t_time1, t_time2))
                conn.commit()
                cursor.close()
                """ ----------------- """

                clearRow()
            elif worksheet.cell(i, j).ctype == 3: #дата, время выезда и прибытия
                ms_date_number = worksheet.cell(i, j).value
                year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number, workbook.datemode)
                py_date = datetime.datetime(year, month, day, hour, minute, second)
                if py_date.year == 1970:
                    if t_time1 == None:
                        t_time1 = py_date
                        t_time2 = None
                    else:
                        t_time2 = py_date
                elif py_date.year == 2020 or py_date.year == 2021 or py_date.year == 2022:
                    t_date = py_date
                    t_time1 = None
                    t_time2 = None
            elif worksheet.cell(i, j).ctype == 1 and "Возраст:" in worksheet.cell(i, j).value: #возраст и кто вызвал

                if 'лет' in worksheet.cell(i, j + 2).value:
                    t_age = int(re.findall('[0-9]+', worksheet.cell(i, j + 2).value)[0])
                elif 'мес' in worksheet.cell(i, j + 2).value:
                    t_age = int(re.findall('[0-9]+', worksheet.cell(i, j + 2).value)[0])/100
                else:
                    t_age = 0

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
                if 'ПСМП' in worksheet.cell(i, j + 5).value:
                    t_station = int(re.findall('[0-9]+', worksheet.cell(i, j + 5).value)[0])
                else:
                    t_station = 0
            elif worksheet.cell(i, j).ctype == 1 and "Вызов:" in worksheet.cell(i, j).value:  # тип вызова
                if 'Первичный' in worksheet.cell(i, j + 2).value:
                    t_type = 0
                elif 'Повторный' in worksheet.cell(i, j + 2).value:
                    t_type = 1
                elif 'Попутный' in worksheet.cell(i, j + 2).value:
                    t_type = 2
                else:
                    t_type = 0
            elif worksheet.cell(i, j).ctype == 1 and "Вид:" in worksheet.cell(i, j).value:  # состояние
                if 'несчастный' in worksheet.cell(i, j + 3).value:
                    t_state = 0
                elif 'неотложное' in worksheet.cell(i, j + 3).value:
                    t_state = 1
                else :
                    t_state = 2
            elif worksheet.cell(i, j).ctype == 1 and "Результат:" in worksheet.cell(i, j).value:  # Результат
                if 'на месте' in worksheet.cell(i, j + 2).value:
                    t_result = 0
                elif 'отказ' in worksheet.cell(i, j + 2).value:
                    t_result = 1
                elif 'до приезда' in worksheet.cell(i, j + 2).value:
                    t_result = 2
                elif 'в присутствии' in worksheet.cell(i, j + 2).value:
                    t_result = 3
                elif 'здоров' in worksheet.cell(i, j + 2).value:
                    t_result = 4
                elif 'в больницу' in worksheet.cell(i, j + 2).value:
                    t_result = 5
                elif 'до прибытия' in worksheet.cell(i, j + 2).value:
                    t_result = 6
                elif 'травмпункт' in worksheet.cell(i, j + 2).value:
                    t_result = 7
                else:
                    t_result = 0
            elif worksheet.cell(i, j).ctype == 1 and "Адрес:" in worksheet.cell(i, j).value:  # адрес
                if worksheet.cell(i, j + 1).value == '':
                    t_from = '-'
                else:
                    if '***' in worksheet.cell(i, j + 1).value and 'кв' in worksheet.cell(i, j + 1).value or 'к***' in worksheet.cell(i, j + 1).value:
                        t_from = ','.join(worksheet.cell(i, j + 1).value.split(',')[:-1])
                    else:
                        t_from = worksheet.cell(i, j + 1).value
            elif worksheet.cell(i, j).ctype == 1 and "Повод:" in worksheet.cell(i, j).value:  # повод
                if worksheet.cell(i, j + 1).value == '':
                    t_reason = '-'
                else:
                    if 'Температура' in worksheet.cell(i, j + 1).value:
                        t_reason = worksheet.cell(i, j + 1).value.replace("Температура", "Темп")
                    else:
                        t_reason = worksheet.cell(i, j + 1).value
            elif worksheet.cell(i, j).ctype == 1 and "Диагноз:" in worksheet.cell(i, j).value:  # диагноз
                if worksheet.cell(i, j + 1).value == '':
                    t_diag = '-'
                else:
                    t_diag = worksheet.cell(i, j + 1).value
            elif worksheet.cell(i, j).ctype == 1 and "Доставлен:" in worksheet.cell(i, j).value:  # доставлен в
                if worksheet.cell(i, j + 1).value == '':
                    t_to = '-'
                else:
                    if 'Городская клиническая больница' in worksheet.cell(i, j + 1).value:
                        t_to = worksheet.cell(i, j + 1).value.replace("Городская клиническая больница", "ГКБ")
                    elif 'Городская поликлиника' in worksheet.cell(i, j + 1).value:
                        t_to = worksheet.cell(i, j + 1).value.replace("Городская поликлиника", "ГП")
                    elif 'Городская больница' in worksheet.cell(i, j + 1).value:
                        t_to = worksheet.cell(i, j + 1).value.replace("Городская больница", "ГБ")
                    elif 'Детская городская больница' in worksheet.cell(i, j + 1).value:
                        t_to = worksheet.cell(i, j + 1).value.replace("Детская городская больница", "ДГБ")
                    elif 'Детская городская поликлиника' in worksheet.cell(i, j + 1).value:
                        t_to = worksheet.cell(i, j + 1).value.replace("Детская городская поликлиника", "ДГП")
                    elif 'Детская городская клиническая больница' in worksheet.cell(i, j + 1).value:
                        t_to = worksheet.cell(i, j + 1).value.replace("Детская городская клиническая больница", "ДГКБ")
                    else:
                        t_to = worksheet.cell(i, j + 1).value

if conn is not None:
    conn.close()
    print('Database connection closed.')