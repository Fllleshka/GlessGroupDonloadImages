import datetime
import requests
import win32com.client
# Импорт баблиотеки для работы в APIGoogle
import gspread
from dates import *

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
    #print(alldates)
    # Выясняем текущий мясяц
    today = datetime.datetime.today()
    todayyear = int(today.strftime("%Y"))
    listmontheng = [datetime.date(todayyear, 1, 1).strftime("%B"), datetime.date(todayyear, 2, 1).strftime("%B"), datetime.date(todayyear, 3, 1).strftime("%B"), datetime.date(todayyear, 4, 1).strftime("%B"), datetime.date(todayyear, 5, 1).strftime("%B"), datetime.date(todayyear, 6, 1).strftime("%B"), datetime.date(todayyear, 7, 1).strftime("%B"), datetime.date(todayyear, 8, 1).strftime("%B"), datetime.date(todayyear, 9, 1).strftime("%B"), datetime.date(todayyear, 10, 1).strftime("%B"), datetime.date(todayyear, 11, 1).strftime("%B"), datetime.date(todayyear, 12, 1).strftime("%B")]
    #print(listmontheng)
    listmonthrus = ["ЯНВАРЬ", "ФЕВРАЛЬ", "МАРТ", "АПРЕЛЬ", "МАЙ", "ИЮНЬ", "ИЮЛЬ", "АВГУСТ", "СЕНТЯБРЬ", "ОКТЯБРЬ", "НОЯБРЬ", "ДЕКАБРЬ"]
    #print(listmonthrus)
    todaymontheng = today.strftime("%B")
    #print(todaymontheng)
    todaymonthrus = listmonthrus[listmontheng.index(todaymontheng)]
    #print("Текущий месяц: ", todaymonthrus)

    # Ищем стартовую ячейку, для определения графика на этот месяц
    index = 0
    indexmonth = 0
    # Перебираем все элементы и находим нужную ячейку с текущим месяцем
    for element in alldates:
        index = index + 1
        #print("[" + str(element) + "]\t", index)
        if str(element) == todaymonthrus:
            indexmonth = index
    #print("Ячейка для старта данных: ", indexmonth)
    # Формируем название ячейки начала импорта
    firstcell = "B" + str(indexmonth)
    # Формируем название ячейки конца импорта
    lastcell = "AG" + str(indexmonth + 31)
    #print("\t\t", firstcell, "\t\t", lastcell)
    # Формируем строку для импорта
    cellsrange = firstcell + ":" + lastcell
    #print("\t\t", cellsrange)
    # Импортируем данные за нужный нам месяц
    datesforsolution = sheet.Range(cellsrange)
    #print("\t", datesforsolution)
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
    lenmanagers = len(massmanagers)-1
    #print("Количество менеджеров: ", lenmanagers)
    for i in range(0, lenmanagers):
        managerlist = []
        index = 0
        for element in dates:
            managerlist.append(element)
            index = index + 1
            if index == countdaysinmonth + 1:
                break
        #print(managerlist)
        managerlists.append(managerlist)
        del dates[0:countdaysinmonth + 1]

    # Выясняем график работы ПП
    deldates = 32*8
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

    return managerlists

# Функция активирования менеджеров
def selectmenegers(managerlists):
    # Выясняем текущй день
    today = datetime.datetime.today()
    todayday = int(today.strftime("%d"))
    print("Сегодня:", todayday, today.strftime("%B") ,int(today.strftime("%Y")))
    flag = True
    # Изменяем статусы менеджеров call центра
    for element in managerlists:
        if element[todayday] != "В":
            numbermanager = numbermanagers[massmanagers.index(element[0])]
            print("Необходимо активировать телефон: ", element[0], "\t[", element[todayday], "]", "'", numbermanager, "'")
            urlforapi = urlapi + str(numbermanager) + '/agent'
            statusrequest = requests.put(urlforapi, params = paramsonline, headers = headers)
            if statusrequest == "<Response [403]>":
                flag = False
                print("\tЧто-то пошло не так... Нет ответа по запросу изменения статуса")
            else:
                statusget = requests.get(urlforapi, headers=headers).text
                print("\tСтатус менеджера: ", element[0], " = ", statusget)
        else:
            numbermanager = numbermanagers[massmanagers.index(element[0])]
            print("Необходимо деактивировать телефон: ", element[0], "\t[", element[todayday], "]", "'", numbermanager, "'")
            urlforapi = urlapi + str(numbermanager) + '/agent'
            statusrequest = requests.put(urlforapi, params = paramoffline, headers = headers)
            if statusrequest == "<Response [403]>":
                flag = False
                print("\tЧто-то пошло не так... Нет ответа по запросу изменения статуса")
            else:
                statusget = requests.get(urlforapi, headers=headers).text
                print("\tСтатус менеджера: ", element[0], " = ", statusget)
    if flag == True:
        return "\tCall центр успешно настроен."
    else:
        return "\tВ работе функции произошла ошибка"

# Функция записи логов фотографий
def createnewarrowinlogs(lenphotos):
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
    today = datetime.datetime.today().strftime("%m.%d.%Y | %H:%M:%S")
    # Добавляем строку в конец фаила логгирования
    worksheet.update_cell(newstr, 1, newnumber)
    worksheet.update_cell(newstr, 2, today)
    worksheet.update_cell(newstr, 3, lenphotos)

# Функция записи логов Call Cener
def createnewarrowincallcenter():
    # Подключаемся к сервисному аккаунту
    gc = gspread.service_account(CREDENTIALS_FILE)
    # Подключаемся к таблице по ключу таблицы
    table = gc.open_by_key(sheetkey)
    # Открываем нужный лист
    worksheet = table.worksheet("LogsCallCenter")
    # Получаем данные с листа
    dates = worksheet.get_values()
    # Получаем номер самой последней строки
    newstr = len(worksheet.col_values(1)) + 1
    # Вычисляем номер строки
    newnumber = newstr - 1
    # Определяем время выполения операции
    today = datetime.datetime.today().strftime("%m.%d.%Y | %H:%M:%S")
    # Выясняем данные кто работает
    managerslist = []
    # Выясняем статусы менеджеров
    for element in numbermanagers:
        urlforapi = urlapi + element + '/agent'
        status = requests.get(urlforapi, headers=headers).text
        managerslist.append(status)
    # Добавляем строку в конец фаила логгирования
    worksheet.update_cell(newstr, 1, newnumber)
    worksheet.update_cell(newstr, 2, today)
    worksheet.update_cell(newstr, 3, managerslist[0])
    worksheet.update_cell(newstr, 4, managerslist[1])
    worksheet.update_cell(newstr, 5, managerslist[2])
    worksheet.update_cell(newstr, 6, managerslist[3])