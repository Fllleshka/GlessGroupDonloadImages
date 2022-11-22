import datetime
import win32com.client

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
    # Перебираем все элекменты и находим нужную ячейку с текущим месяцем
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
    print(dates)

    index = 0
    for element in dates:
        index = index + 1
        if element == "None":
            # Определяем количество дней в месяце
            countdaysinmonth = index - 1
            break
    print (countdaysinmonth)

    # Считаем ненужные данные
    deldates = 1 + countdaysinmonth + 1 + countdaysinmonth + 1
    print(deldates)
    # Удаляем ненужные данные
    index = 0
    for element in dates:
        del dates[0]
        index = index + 1
        if index == deldates:
            break
    print(dates)

    managers_list = []
    index = 0
    managerlist = []
    for element in dates:
        index = index + 1
        managerlist.append(element)
        if index == 32:
            break
    managers_list.append(managerlist)

    # Выясняем текущй день
    today = datetime.datetime.today()
    todatday = today.strftime("%d")
    print(todatday)