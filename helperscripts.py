import datetime
import shutil
import time
from PIL import Image
import os
import requests
import win32com.client
import gspread
from dates import *
import openpyxl
import json

# Функция записи логов фотографий
def createnewarrowinlogs(lenphotos):
    try:
        # Подключаемся к сервисному аккаунту
        gc = gspread.service_account(CREDENTIALS_FILE)
        # Подключаемся к таблице по ключу таблицы
        table = gc.open_by_key(sheetkey)
        # Открываем нужный лист
        worksheet = table.worksheet("LogsPhotos")
        # Получаем данные с листа
        dates = worksheet.get_values()
        # Получаем номер самой последней строки
        newstr = len(worksheet.col_values(1)) + 1
        # Вычисляем номер строки
        newnumber = newstr - 1
        # Определяем время выполения операции
        today = datetime.datetime.today().strftime("%d.%m.%Y | %H:%M:%S")
        # Добавляем строку в конец фаила логгирования
        worksheet.update_cell(newstr, 1, newnumber)
        worksheet.update_cell(newstr, 2, today)
        worksheet.update_cell(newstr, 3, lenphotos)
    except:
        print("Логгирование фотографий сломалось(")

# Функция записи обновления файла Call Cener
def createnewarrowincallcenter2():
    try:
        # Подключаемся к сервисному аккаунту
        gc = gspread.service_account(CREDENTIALS_FILE)
        # Подключаемся к таблице по ключу таблицы
        table = gc.open_by_key(sheetkey)
        # Открываем нужный лист
        worksheet = table.worksheet("LogsCallCenter")
        # Получаем номер самой последней строки
        newstr = len(worksheet.col_values(1)) + 1
        # Вычисляем номер строки
        newnumber = newstr - 1
        # Определяем время выполения операции
        today = datetime.datetime.today().strftime("%d.%m.%Y | %H:%M:%S")
        # Определяем диапазон для обьединения ячеек
        mergerange = "C" + str(newstr) + ":F" + str(newstr)
        #print(mergerange)
        # Обьединяем ячейки да записи
        worksheet.merge_cells(mergerange)
        # Добавляем запись в таблицу логгирования
        worksheet.update_cell(newstr, 1, newnumber)
        worksheet.update_cell(newstr, 2, today)
        worksheet.update_cell(newstr, 3, "Фаил [График 2023 ТЕСТ.xlsx] обновлён")
        # Окрашивание ячейки
        color = {"backgroundColor": {"red": 0.94, "green": 0.9, "blue": 0.15}, "horizontalAlignment": "CENTER"}
        worksheet.format("C" + str(newstr), color)
        # Делаем центрирование ячейки
        worksheet.format(mergerange, {"horizontalAlignment": "CENTER"})
    except Exception as e:
        print(f"Логгирование call-центра сломалось: {e}")

# Вспомогательная функция для importatesfromftp с преобразованием в номальный формат даты
def get_datetime_format(date_time):
    # Преобразовать в объект даты и времени
    date_time = datetime.strptime(date_time, '%Y%m%d%H%M%S')
    # Преобразовать в удобочитаемую строку даты и времени
    return date_time.strftime('%Y/%m/%d %H:%M:%S')

# Функция импорта данных
def importatesfromftp(ftp, listdirectors, element):
    # Добавляем элемент в лист
    listdirectors.append(int(element))
    # Определяем путь для папки
    path = "/" + str(element) + "/"
    # Изменение каталог работы
    ftp.cwd(path)
    # Получаем лист всех фаилов из папки
    list = ftp.nlst()
    # Удалем первые 2 элемента (так как на сервере система Linux)
    list.pop(0)
    list.pop(0)
    # Сортируем фаилы по возрастанию
    list.sort()
    # Добавляем данные в массив
    returnmass = list
    # Возвращаем полученный список
    return returnmass

# Функция получения размера изображения
def get_size_format(b, factor=1024, suffix="B"):
    """
    Scale bytes to its proper byte format
    e.g:
        1253656 => '1.20MB'
        1253656678 => '1.17GB'
    """
    for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
        if b < factor:
            return f"{b:.2f}{unit}{suffix}"
        b /= factor
    return f"{b:.2f}Y{suffix}"

# Функция конвертации изображения (уменьшения веса и подгонка под заданные параметры)
def convertimage(path):
    # Размеры изображения на выходе
    width = 1920
    height = 1440
    # Загружаем фотографию в память
    img = Image.open(path)
    # Первоначальный размер картинки
    olddimensions = img.size
    # Получаем размер изображения до компрессии
    image_size = os.path.getsize(path)
    oldsize = get_size_format(image_size)
    # Преобразуем изображение приводя его к нужным высоте и ширине и уменьшая размер
    img.thumbnail(size=(width, height))
    if img.height > 1080:
        difference_height = (height - 1080) / 2
        img = img.crop((0, 0 + difference_height, 1920, height - difference_height))
    # Сохраняем изображение
    img.save(path, optimize=True, quality=95)
    # Получаем новые размеры картинки
    newdimesions = img.size
    # Получаем размер изображение после компрессии
    image_size = os.path.getsize(path)
    newsize = get_size_format(image_size)
    # Печатаем в кносоль результат
    print(path, "с шириной, высотой: ", olddimensions, " и размером: ", oldsize, "была преобразована в: ", newdimesions,
          " и ", newsize)

# Функциия импорта и систематизация статистики по звонкам
def collectionofinformation():
    print("Время иморта статистики по звонкам")

    # Класс звонка
    class phoneCall:
        def __init__(self, name_manager, incoming_call_time, incoming_call_number, call_duration, direction, status):
            # ФИО менеджера
            self.name_manager = name_manager
            # Время входящего звонка
            self.incoming_call_time = incoming_call_time
            # Телефон звонка
            self.incoming_call_number = incoming_call_number
            # Продолжительность звонка
            self.call_duration = call_duration
            # Тип вызова (INBOUND-Входящий вызов, OUTBOUND-Исходящий вызов))
            self.direction = direction
            # Статус звонка (RECIEVED-принятый, MISSED-пропущенный, PLACED-исходящий)
            self.status = status

        # Функция печати данных
        def printdates(self):
            print(f"\t{self.name_manager}"
                  f"\t{self.incoming_call_time}"
                  f"\t\t{self.incoming_call_number}"
                  f"\t{self.call_duration}"
                  f"\t\t{self.direction}"
                  f"\t\t{self.status}")

    try:
        # Дата начала отчёта (вчерашний день начало для)
        dateAndTimeStart = (datetime.datetime.today() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")
        dateAndTimeStart += "T00:00:00.000Z"
        # Дата окончания отчёта (сегодняшний день начало дня)
        dateAndTimeEnd = datetime.datetime.today().strftime("%Y-%m-%d")
        dateAndTimeEnd += "T00:00:00.000Z"
        calls = []

        for element in numbermanagers:
            paramsinfo['userId'] = element
            paramsinfo['dateTo'] = dateAndTimeEnd
            paramsinfo['dateFrom'] = dateAndTimeStart
            statusrequest = requests.get(urlforstatistics, params=paramsinfo, headers=headers)
            # statusrequests.append(statusrequest.text)
            jsonData = json.loads(statusrequest.text)
            for elem in jsonData:
                #print(elem)
                #print(elem['direction'])
                # Вычисляем время звонка
                dateandtime = datetime.datetime.fromtimestamp(elem['startDate'] / 1000)
                # Вычисляем продолжительность разговора
                dateandtime2 = datetime.timedelta(milliseconds=elem['duration'])
                # Если вызов входящий
                if elem['direction'] == 'INBOUND':
                    # Вычисляем телефон абонента
                    phone = elem['phone_from']
                    # Добавляем в массив звонков экземпляр класса phoneCall
                    calls.append(phoneCall(name_manager=elem['abonent']['firstName'],
                                           incoming_call_time=dateandtime,
                                           incoming_call_number=phone,
                                           call_duration=dateandtime2,
                                           direction=elem['direction'],
                                           status=elem['status']))
                # Иначе вызов исходящий:
                else:
                    # Вычисляем телефон абонента
                    phone = elem['phone_to']
                    # Добавляем в массив звонков экземпляр класса phoneCall
                    calls.append(phoneCall(name_manager=elem['abonent']['firstName'],
                                           incoming_call_time=dateandtime,
                                           incoming_call_number=phone,
                                           call_duration=dateandtime2,
                                           direction=elem['direction'],
                                           status=elem['status']))
        dates = []
        # Подключаемся к сервисному аккаунту
        gc = gspread.service_account(CREDENTIALS_FILE)
        # Подключаемся к таблице по ключу таблицы
        table = gc.open_by_key(sheetkey)
        # Открываем нужный лист
        worksheet = table.worksheet("StatisticOfCalls")
        # Получаем номер самой последней строки
        newstr = len(worksheet.col_values(4)) + 1
        # Вычисляем номер строки
        newnumber = newstr - 2
        dates.append(newnumber)
        # Определяем время выполения операции
        today = datetime.datetime.today().strftime("%d.%m.%Y | %H:%M:%S")
        dates.append(today)
        # Выводим дату за которую приводим статистику
        statdate = (datetime.datetime.today() + datetime.timedelta(days=-1)).strftime("%d.%m.%Y")
        dates.append(statdate)
        # Обявлем массивы для подсчёта
        massmissescals = [0, 0, 0, 0]
        massinboundcalls = [0, 0, 0, 0]
        masssumtimes = [datetime.timedelta(milliseconds=0), datetime.timedelta(milliseconds=0), datetime.timedelta(milliseconds=0), datetime.timedelta(milliseconds=0)]
        # Пробегаемся по всем звонкам и сортируем звонки
        for element in calls:
            # Считаем статистику для первого менеджера
            if element.name_manager == fullmassmanagers[0]:
                addinfoinmass(massmissescals, massinboundcalls, masssumtimes, 0, element)
            # Считаем статистику для второго менеджера
            elif element.name_manager == fullmassmanagers[1]:
                addinfoinmass(massmissescals, massinboundcalls, masssumtimes, 1, element)
            # Считаем статистику для третьего менеджера
            elif element.name_manager == fullmassmanagers[2]:
                addinfoinmass(massmissescals, massinboundcalls, masssumtimes, 2, element)
            # Считаем статистику для четвёртого менеджера
            elif element.name_manager == fullmassmanagers[3]:
                addinfoinmass(massmissescals, massinboundcalls, masssumtimes, 3, element)
            else:
                print("Cтатистика для Неизвестного лица(")
        # Добавляем данные с разбора в результурующий массив
        for element in range(4):
            dates.append(massmissescals[element])
            dates.append(massinboundcalls[element])
            dates.append(converttoseconds(masssumtimes[element].total_seconds()))

        # Проверяем были ли записаны данные ранее
        datesfromtabel = worksheet.row_values(newnumber+1)
        if datesfromtabel[2] == dates[2]:
            print("\t\tДанные уже были записаны")
        else:
            # Записываем получившееся результаты в таблицу
            i = 0
            for element in dates:
                worksheet.update_cell(newstr, i+1, dates[i])
                i += 1

            # Выясняем кто работал в это день
            workedmanagers = [0, 0, 0]
            masscolumns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]
            for element in numbermanagers:
                urlforapi = urlapi + element + '/agent'
                status = requests.get(urlforapi, headers=headers).text
                for elem in range(0, 3):
                    workedmanagers.append(status)

            colorwork = {"backgroundColor": {"red": 0.67, "green": 1.0, "blue": 0.74}, "horizontalAlignment": "CENTER", "borders": {"top": {"style": "SOLID"}, "bottom": {"style": "SOLID"}, "left": {"style": "SOLID"}, "right": {"style": "SOLID"}}}
            coloroutput = {"backgroundColor": {"red": 1.0, "green": 0.78, "blue": 0.77}, "horizontalAlignment": "CENTER", "borders": {"top": {"style": "SOLID"}, "bottom": {"style": "SOLID"}, "left": {"style": "SOLID"}, "right": {"style": "SOLID"}}}
            colornone = {"borders": {"top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},"left": {"style": "SOLID"}, "right": {"style": "SOLID"}}}

            # Записываем получившееся результаты в таблицу
            i = 0
            for element in dates:
                match workedmanagers[i]:
                    case '"ONLINE"':
                        worksheet.update_cell(newstr, i + 1, dates[i])
                        worksheet.format(masscolumns[i] + str(newstr), colorwork)
                    case '"OFFLINE"':
                        worksheet.update_cell(newstr, i + 1, dates[i])
                        worksheet.format(masscolumns[i] + str(newstr), coloroutput)
                    case _:
                        worksheet.update_cell(newstr, i + 1, dates[i])
                        worksheet.format(masscolumns[i] + str(newstr), colornone)
                i += 1
    except Exception as e:
        print(f"Логгирование статистики по звонкам сломалось: {e}")

# Функция разбора данных по звонкам
def addinfoinmass(massmissescals, massinboundcalls, masssumtimes, numbermanager, elemclass):
    # Если вызов входящий пропущенный
    if elemclass.direction == "INBOUND" and elemclass.status == "MISSED":
        massmissescals[numbermanager] += 1
    # Если вызов входящий принятый
    elif elemclass.direction == "INBOUND" and elemclass.status == "RECIEVED":
        massinboundcalls[numbermanager] += 1
        masssumtimes[numbermanager] += elemclass.call_duration
    return [massmissescals, massinboundcalls, masssumtimes]

# Функция удобно представления времени разговора из миллисекунд в нормальное представление
def converttoseconds(totseconds):
    hours, remainder = divmod(int(totseconds), 3600)
    minutes, seconds = divmod(remainder, 60)
    result = str(hours) + ":" + str(minutes) + ":" + str(seconds)
    return result

# Функция сохранения статистики по загруженным фотографиям
def generationstatuploadphotos():
    try:
        # Проверяем дату сегодняшнюю
        today = datetime.datetime.today()
        todaytime = today.strftime("%d")
        #print(f"Сегодняшнее числo: {todaytime}")
        # Проверяем если начало месяца (01 число)
        if todaytime == "01":
            #print(f"Сегодняшнее числo: {todaytime}")
            # Вычисляем месяц за который сохраняем статистику
            statmonth = today.replace(day=15).strftime("%B")
            statyear = today.replace(day=15).strftime("%Y")
            statmonthandyear = statmonth + " " + statyear

            # Вычисляем последнюю строку для записи статистики
            # Подключаемся к сервисному аккаунту
            gc = gspread.service_account(CREDENTIALS_FILE)
            # Подключаемся к таблице по ключу таблицы
            table = gc.open_by_key(sheetkey)
            # Открываем нужный лист
            worksheet = table.worksheet("LogsPhotos")
            # Получаем номер строки для записи в стоблце L
            newstr = len(worksheet.col_values(12)) + 1

            # Получаем данные из столбца H
            massvalues = worksheet.get_values('H2:H6')
            massvalues2 = []
            sumphotos = 0
            # Преобразовываем массив
            for element in massvalues:
                massvalues2.append(int(element[0]))
                sumphotos += int(element[0])

            colornone = {"borders": {"top": {"style": "SOLID"}, "bottom": {"style": "SOLID"}, "left": {"style": "SOLID"},"right": {"style": "SOLID"}}}
            masscolumns = ["L", "M", "N", "O", "P", "Q", "R"]

            # Запись данных в табличку
            for element in range (0, 7):
                column = element + 12
                if column == 12:
                    worksheet.update_cell(newstr, column, statmonthandyear)
                    worksheet.format(masscolumns[element] + str(newstr), colornone)
                elif column == 18:
                    worksheet.update_cell(newstr, column, sumphotos)
                    worksheet.format(masscolumns[element] + str(newstr), colornone)
                else:
                    worksheet.update_cell(newstr, column, massvalues2[element-1])
                    worksheet.format(masscolumns[element] + str(newstr), colornone)

            # Обнуляем значения, которые подсчитываются онлайн
            for element in range(2,7):
                worksheet.update_cell(element, 8, 0)

            # Записываем дату с коротой считаются фотографии
            nulldate = today.strftime("%d %B %Y")
            worksheet.update_cell(2, 9, nulldate)

    except Exception as e:
        print(f"Логгирование статистики фотографий сломалось: {e}")

# Функция вывода данных о запущенной функции
def printer(time, name):
    print(f"\t{time}\t\t{name}.")
