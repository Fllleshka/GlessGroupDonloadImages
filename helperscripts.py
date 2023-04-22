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

# Функция компирования данных в фаил для работы
def checkupdatedatesexcel():
    #print("Проверка фаила на последнее изменение")
    # Вычисляем время последнего изменения основного документа
    file1 = openpyxl.load_workbook(mainfile).properties.modified
    #print(f"Дата изменения основного фаила: {file1}")
    # Вычисляем дату последнего изменения рабочего документа
    file2 = openpyxl.load_workbook(pathfile).properties.modified
    #print(f"Дата изменения фаила для работы: {file2}")
    # Вычисляем разницу времён
    diff_times = file1 - file2
    #print(f"Разница в датах изменений: {diff_times}")
    # Устанавливаем количество часов для синхронизации фаилов
    deltatime = datetime.timedelta(days=0, hours=0, minutes=5)
    #print(f"Таймер синхронизации: {deltatime}")
    # Если deltatime меньше разницы во времени изменения файлов
    if deltatime < diff_times:
        print("Синхронизация требуется")
        # Выполняем копирование
        shutil.copy2(mainfile, pathfile, follow_symlinks=True)
        print("Статус операции изменения фаила: [Фаил обновлён]")
        # Необходимо сделать измение пароля от файла

        #wb = openpyxl.load_workbook(pathfile, password)
        #print("\t\t\t", wb)
        #ws = wb.active
        #print("\t\t\t", ws)
        #wb.save(pathfile, password)
        #wb.security.workbookPassword = password
        #Экземпляр COM обьекта
        #xlApp = win32com.client.Dispatch("Excel.Application")
        #print("Экзмепляр COM: ", xlApp)
        #print ("Excel library version:", xlApp.Version)
        #xlwb = xlApp.Workbooks.Open(pathfile, Password=password)
        #print ("Фаил Ecxel: ", xlwb)
        # Закрываем COM обьект
        #xlApp.Quit()

        # Логгирование обновления фаила
        createnewarrowincallcenter2()
    # Иначе синхронизацию не выполняем.
    else:
        print("Синхронизация не требуется")

# Функция импорта данных из Excel
def importdatesformexcel(path, password):
    # Экземпляр COM обьекта
    xlApp = win32com.client.Dispatch("Excel.Application")
    # Открываем фаил
    xlwb = xlApp.Workbooks.Open(path, False, True, None, password)
    # Выбираем лист(таблицу)
    sheet = xlwb.ActiveSheet
    # Выбираем данные из range
    alldates = sheet.Range("B1:B451")
    # print(alldates)

    # Выясняем текущий мясяц
    today = datetime.datetime.today()
    todayyear = int(today.strftime("%Y"))
    listmontheng = [datetime.date(todayyear, 1, 1).strftime("%B"), datetime.date(todayyear, 2, 1).strftime("%B"),
                    datetime.date(todayyear, 3, 1).strftime("%B"), datetime.date(todayyear, 4, 1).strftime("%B"),
                    datetime.date(todayyear, 5, 1).strftime("%B"), datetime.date(todayyear, 6, 1).strftime("%B"),
                    datetime.date(todayyear, 7, 1).strftime("%B"), datetime.date(todayyear, 8, 1).strftime("%B"),
                    datetime.date(todayyear, 9, 1).strftime("%B"), datetime.date(todayyear, 10, 1).strftime("%B"),
                    datetime.date(todayyear, 11, 1).strftime("%B"), datetime.date(todayyear, 12, 1).strftime("%B")]
    # print(listmontheng)
    listmonthrus = ["ЯНВАРЬ", "ФЕВРАЛЬ", "МАРТ", "АПРЕЛЬ", "МАЙ", "ИЮНЬ", "ИЮЛЬ", "АВГУСТ", "СЕНТЯБРЬ", "ОКТЯБРЬ",
                    "НОЯБРЬ", "ДЕКАБРЬ"]
    # print(listmonthrus)
    todaymontheng = today.strftime("%B")
    # print(todaymontheng)
    todaymonthrus = listmonthrus[listmontheng.index(todaymontheng)]
    # print("Текущий месяц: ", todaymonthrus)

    # Ищем стартовую ячейку, для определения графика на этот месяц
    index = 0
    indexmonth = 0
    # Перебираем все элементы и находим нужную ячейку с текущим месяцем
    for element in alldates:
        index = index + 1
        # print("[" + str(element) + "]\t", index)
        if str(element) == todaymonthrus:
            indexmonth = index
    # print("Ячейка для старта данных: ", indexmonth)
    # Формируем название ячейки начала импорта
    firstcell = "B" + str(indexmonth)
    # Формируем название ячейки конца импорта
    lastcell = "AG" + str(indexmonth + 31)
    # print("\t\t", firstcell, "\t\t", lastcell)
    # Формируем строку для импорта
    cellsrange = firstcell + ":" + lastcell
    # print("\t\t", cellsrange)
    # Импортируем данные за нужный нам месяц
    datesforsolution = sheet.Range(cellsrange)
    # print("\t", datesforsolution)
    # Формируем данные в список
    listdatesforsolution = []
    for element in datesforsolution:
        listdatesforsolution.append(str(element))
    # Закрываем фаил
    xlwb.Close()
    # Закрываем COM обьект
    xlApp.Quit()

    # Возвращаем данные
    return listdatesforsolution

# Функция для выбоки по данным
def chosedates(dates):
    # Удаляем первый элемент
    del dates[0]
    #print("Изначальный массив:")
    #print("\t", dates)

    # Считаем дни в месяце
    index = 0
    for element in dates:
        index = index + 1
        if element == "None" or element == "Торговля":
            # Определяем количество дней в месяце
            countdaysinmonth = index - 1
            break
    #print("Дней в месяце: ", countdaysinmonth)

    # Удаляем ненужные данные
    for element in dates:
        if element == massmanagers[0]:
            delelements = dates.index(element)
    del dates[0:delelements]
    #print("\t", dates)

    # Разбиваем массив для конкретизации графика каждого менеджера
    managerlists = []
    lenmanagers = len(massmanagers) - 1
    #print("Количество менеджеров: ", lenmanagers)
    for i in range(0, lenmanagers):
        managerlist = []
        index = 0
        for element in dates:
            managerlist.append(element)
            index += 1
            if index == 32:
                break
        #print(managerlist)
        managerlists.append(managerlist)
        del dates[0:32]

    # Выясняем график работы ПП
    deldates = 32 * 8
    # Удаляем ненужные данные
    del dates[0:deldates]
    #print(dates)
    # Добавляем в массив работников данные
    managerlist = []
    index = 0
    for element in dates:
        managerlist.append(element)
        index = index + 1
        if index == countdaysinmonth + 1:
            break
    managerlists.append(managerlist)
    #print(managerlist)
    return managerlists

# Функция активирования менеджеров
def selectmenegers(managerlists):
    # Выясняем текущй день
    today = datetime.datetime.today()
    todayday = int(today.strftime("%d"))
    print("Сегодня:", todayday, today.strftime("%B"), int(today.strftime("%Y")))
    flag = True
    massworkmanagers = []
    try:
        # Изменяем статусы менеджеров call центра
        for element in managerlists:
            if element[todayday] == "В" or element[todayday] == "O" or element[todayday] == "О" or element[todayday] == "Х":
                numbermanager = numbermanagers[massmanagers.index(element[0])]
                print("Необходимо деактивировать телефон: ", element[0], "\t[", element[todayday], "]", "'",
                      numbermanager,
                      "'")
                urlforapi = urlapi + str(numbermanager) + '/agent'
                statusrequest = requests.put(urlforapi, params=paramoffline, headers=headers)
                if statusrequest == "<Response [403]>":
                    flag = False
                    print("\tЧто-то пошло не так... Нет ответа по запросу изменения статуса")
                else:
                    statusget = requests.get(urlforapi, headers=headers).text
                    print("\tСтатус менеджера: ", element[0], " = ", statusget)
            else:
                numbermanager = numbermanagers[massmanagers.index(element[0])]
                print("Необходимо активировать телефон: ", element[0], "\t[", element[todayday], "]", "'",
                      numbermanager,
                      "'")
                urlforapi = urlapi + str(numbermanager) + '/agent'
                statusrequest = requests.put(urlforapi, params=paramsonline, headers=headers)
                if statusrequest == "<Response [403]>":
                    flag = False
                    print("\tЧто-то пошло не так... Нет ответа по запросу изменения статуса")
                else:
                    # Дополнительное условие для последнего менеджера
                    massworkmanagers.append(element[todayday])
                    if len(massworkmanagers) == 4:
                        # Если 3 других менеджера работают, то 4 должен быть отключён
                        if (massworkmanagers[0] == '9.0' or massworkmanagers[0] == '10.0') and (massworkmanagers[1] == '9.0' or massworkmanagers[1] == '10.0') and (massworkmanagers[2] == '9.0' or massworkmanagers[2] == '10.0'):
                            requests.put(urlforapi, params=paramoffline, headers=headers)
                    statusget = requests.get(urlforapi, headers=headers).text
                    print("\tСтатус менеджера: ", element[0], " = ", statusget)

        if flag == True:
            return "\tCall центр успешно настроен."
        else:
            return "\tВ работе функции произошла ошибка"
    except Exception as e:
        print(f"В работе call-центра произошла ошибка: {e}")
        time.sleep(10)

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

# Функция записи логов Call Center
def createnewarrowincallcenter():
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
        # Выясняем данные кто работает
        managerslist = []
        # Выясняем статусы менеджеров
        for element in numbermanagers:
            urlforapi = urlapi + element + '/agent'
            status = requests.get(urlforapi, headers=headers).text
            managerslist.append(status)
        # Проверяем изменится ли call центр
        dates = worksheet.row_values(newnumber)
        # Если данные уже сегодня записывались, то не дублируем их
        if dates[2] == managerslist[0] and dates[3] == managerslist[1] and dates[4] == managerslist[2] and dates[5] == managerslist[3] and str(dates[1])[:10] == str(today)[:10]:
            print("\t\tДанные уже были записаны")
        # Если же эти данные не были записаны, записываем
        else:
            greencolor = {"backgroundColor": {"red": 0.63, "green": 1.0, "blue": 0.65}, "horizontalAlignment": "CENTER"}
            redcolor = {"backgroundColor": {"red": 1.0, "green": 0.65, "blue": 0.63}, "horizontalAlignment": "CENTER"}
            # Добавляем строку в конец фаила логгирования
            worksheet.update_cell(newstr, 1, newnumber)
            worksheet.update_cell(newstr, 2, today)
            worksheet.update_cell(newstr, 3, managerslist[0])
            if managerslist[0] == '"ONLINE"':
                worksheet.format("C" + str(newstr), greencolor)
            else:
                worksheet.format("C" + str(newstr), redcolor)
            worksheet.update_cell(newstr, 4, managerslist[1])
            if managerslist[1] == '"ONLINE"':
                worksheet.format("D" + str(newstr), greencolor)
            else:
                worksheet.format("D" + str(newstr), redcolor)
            worksheet.update_cell(newstr, 5, managerslist[2])
            if managerslist[2] == '"ONLINE"':
                worksheet.format("E" + str(newstr), greencolor)
            else:
                worksheet.format("E" + str(newstr), redcolor)
            worksheet.update_cell(newstr, 6, managerslist[3])
            if managerslist[3] == '"ONLINE"':
                worksheet.format("F" + str(newstr), greencolor)
            else:
                worksheet.format("F" + str(newstr), redcolor)
            # Чтобы программа не крашилась из-за лимита количества запросов ставим sleep
            time.sleep(60)
    except Exception as e:
        print(f"Логгирование call-центра сломалось: {e}")

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
        print(f"Сегодняшнее числo: {todaytime}")
        # Проверяем если начало месяца (01 число)
        if todaytime == "01":
            print(f"Сегодняшнее числo: {todaytime}")
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
    except Exception as e:
        print(f"Логгирование статистики фотографий сломалось: {e}")