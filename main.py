import os
from ftplib import FTP
from helperscripts import *
from threading import Thread
import pythoncom
import win32security

masslocal = [[], [], [], [], []]
massremote = [[], [], [], [], []]

# Функция формирования путей до папок на сервере картинок
def scanfilesinlocalserver():

    global masslocal
    i = 0
    # Главный путь к папкам
    mainpath = '//192.168.20.215/фото товара'
    # Пробегаемся по массиву
    for element in os.listdir(mainpath):
        # Формируем путь к папкам
        pathfolder = mainpath + "/" + element
        match element:
            case "1":
                masslocal[0] = os.listdir(pathfolder)
                masslocal[0].remove('Thumbs.db')
            case "2":
                masslocal[1] = os.listdir(pathfolder)
                masslocal[1].remove('Thumbs.db')
            case "3":
                masslocal[2] = os.listdir(pathfolder)
                masslocal[2].remove('Thumbs.db')
            case "4":
                masslocal[3] = os.listdir(pathfolder)
                masslocal[3].remove('Thumbs.db')
            case "5":
                masslocal[4] = os.listdir(pathfolder)
                masslocal[4].remove('Thumbs.db')
            case _:
                continue

# Функция сканируемых данных на удалённом сервере
def scanfilesinremoteserver():
    global massremote
    try:
        # Данные для связи с удалённым сервером
        ftp = FTP(nameSite)
        ftp.login(ftpLogin, ftpPass)
        # Получаем данные о том какие данные есть на удалённом сервере
        listalldirectors = ftp.nlst()
        listdirectors = []
        for element in listalldirectors:
            match element:
                # Первая папка для синхронизации
                case "1":
                    massremote[0] = importatesfromftp(ftp, listdirectors, element)
                # Вторая папка для синхронизации
                case "2":
                    massremote[1] = importatesfromftp(ftp, listdirectors, element)
                # Третья папка для синхронизации
                case "3":
                    massremote[2] = importatesfromftp(ftp, listdirectors, element)
                # Четвёртая папка для синхронизации
                case "4":
                    massremote[3] = importatesfromftp(ftp, listdirectors, element)
                # Пятая папка для синхронизации
                case "5":
                    massremote[4] = importatesfromftp(ftp, listdirectors, element)
                case _:
                    continue
        listdirectors.sort()
    except:
        print("Синхронизация папок не удалась. Попробуем в следующий раз")
        return

# Проверка на разность данных в локальных папках и на удалённом сервере
def comparisonlists():
    for element in range(0, 5):
        result = list(set(masslocal[element]) - set(massremote[element]))
        if result == []:
            match element:
                case 0:
                   print("Первые\t\tпапки синхронизированны!")
                case 1:
                    print("Вторые\t\tпапки синхронизированны!")
                case 2:
                    print("Третьи\t\tпапки синхронизированны!")
                case 3:
                    print("Четвёртые\tпапки синхронизированны!")
                case 4:
                    print("Пятые\t\tпапки синхронизированны!")
        else:
            print("Разность первых папкок: ", result)
            uploadfiles(element + 1, result)

# Переименование и перемещение картинки по необходимому локальному пути
def renameanduploadimage(pathimage, folder):
    # Начинаем с переименования картинки
    numberfolderfirst = str(pathimage)[53:]
    # Если папка четырёхзначная
    if numberfolderfirst[4] == "/":
        numberfoldersecond = str(numberfolderfirst)[:4]
    elif numberfolderfirst[3] == "/":
        numberfoldersecond = str(numberfolderfirst)[:3]
    else:
        numberfoldersecond = str(numberfolderfirst)[:5]
    # Название картинки
    namepic = numberfoldersecond + str(pathimage)[-4:]
    # Новый путь к картинке
    convertname = str(pathimage)[:53] + numberfoldersecond + "/" + namepic
    # Переименование картинки
    os.rename(pathimage, convertname)

    # Начинаем загрузку фотографии по необходимому местоположению
    newpathfile = str(pathimage)[:29] + str(folder) + "/" + namepic
    shutil.move(convertname, newpathfile)

# Функция загрузки фотографий на сервер
def uploadfiles(numberfolder, result):
    # Главный путь к папкам
    mainpath = '//192.168.20.215/фото товара/'
    pathtofolder = mainpath + str(numberfolder) + "/"

    # Подключение к удалённому серверу по FTP
    ftp = FTP(nameSite)
    ftp.login(ftpLogin, ftpPass)
    ftppath = "/" + str(numberfolder) + "/"
    ftp.cwd(ftppath)

    # Перебираем элементы
    for element in result:
        if element == "Thumbs.db":
            continue
        else:
            path = pathtofolder + element
            # file = open(element, "rb")
            with open(path, "rb") as file:
                ftp.storbinary("STOR " + element, file)
            file.close()
    print("Синхронизация папки ", numberfolder, " завершена.")
    ftp.quit()

# Функция папки с фотографиями для разбора и сортировка их по необходимым папкам с нужными номерами
def scanfolderforimages():
    # Массив загруженных фотографий
    massnewphotos = [0, 0, 0, 0, 0]
    # Получаем лист фаилов находящихся по адресу
    list = os.listdir(mainpath)
    # Проверка наличия фотографий
    # Если папок для разбора нет
    if list == []:
        print("\tДанные по импорту фотографий из папки для разбора отсутствуют")
    # Если есть папки для разбора
    else:
        # Выясняем количество папок
        lenphotos = len(list)
        # Логгирование фотографий
        createnewarrowinlogs(lenphotos)
        # Пробегаемся по элеметам
        for element in list:
            # Обыгрывание Thumbs.db
            if element == "Thumbs.db":
                # Выясняем путь к этому фаилу
                path = mainpath + "/" + element
                # Пытаемся удалить данный фаил
                try:
                    if os.access(path, os.R_OK and os.X_OK):
                        os.remove(path)
                except PermissionError:
                    # Оператор заглушка равноценная отсутствию операции
                    pass
            # Если же это не фаил Thumbs.db
            else:
                # Выясняем путь к этому фаилу
                pathfolder = mainpath + "/" + element
                # Получаем данные о фаилах по этому пути
                nextlist = os.listdir(pathfolder)
                # Если папка пуста то пишем о пустой папке
                if nextlist == []:
                    print("\t Папка ", element, " пуста")
                # Если папка не пуста
                else:
                    numberfolder = 1
                    for elem in nextlist:
                        # Обыгрывание Thumbs.db
                        if elem == "Thumbs.db":
                            continue
                        else:
                            # Условие перебора количества фотографий, так как 6 папки нету
                            if numberfolder >= 6:
                                break
                            else:
                                # Выясняем путь к фаилу
                                pathimage = pathfolder + "/" + elem
                                # Функция сбора статистики по загруженным фотографиям
                                massnewphotos = statisticsphotos(pathimage, massnewphotos)
                                # Уменьшение веса и подгонка фотографии
                                convertimage(pathimage)
                                # Переименоваие и загрузка фотографии
                                renameanduploadimage(pathimage, numberfolder)
                                # Увеличиваем счётчик
                                numberfolder = numberfolder + 1
        # После окончания загрузки фотографий по папкам удаляем папку
        for elem in list:
            # Обыгрывание Thumbs.db
            if element == "Thumbs.db":
                path = mainpath + "/" + element
                try:
                    if os.access(path, os.R_OK and os.X_OK):
                        os.remove(path)
                except PermissionError:
                    pass
            else:
                # Выясняем путь к папке
                path = mainpath + "/" + elem
                # Удаляем полностью папку
                try:
                    shutil.rmtree(path)
                except Exception as e:
                    print("Удаление невозможно. По причине ", e)
        t1 = Thread(target=updatedatesuploadphotos, args=(massnewphotos,))
        t1.start()
        print("Удаление папок завершено")

# Функция сбора статистики по загруженным фотографиям
def statisticsphotos(pathimage, massnewphotos):
    sd = win32security.GetFileSecurity(pathimage, win32security.OWNER_SECURITY_INFORMATION)
    owner_sid = sd.GetSecurityDescriptorOwner()
    match (str(owner_sid)):
        case masssotr.PySID_fleysner:
            massnewphotos[0] += 1
        case masssotr.PySID_kireev:
            massnewphotos[1] += 1
        case masssotr.PySID_pushkar:
            massnewphotos[2] += 1
        case masssotr.PySID_ivanov:
            massnewphotos[3] += 1
        case _:
            massnewphotos[4] += 1
    return massnewphotos

# Функция записи статистики по загруженным фотографиям в GoogleDocs
def updatedatesuploadphotos(massnewphotos):
    try:
        # Подключаемся к сервисному аккаунту
        gc = gspread.service_account(CREDENTIALS_FILE)
        # Подключаемся к таблице по ключу таблицы
        table = gc.open_by_key(sheetkey)
        # Открываем нужный лист
        worksheet = table.worksheet("LogsPhotos")
        # Получаем данные из ячеек
        massolddates = [int(worksheet.get_values(masssotr.CellTable_fleysner)[0][0]), int(worksheet.get_values(masssotr.CellTable_kireev)[0][0]), int(worksheet.get_values(masssotr.CellTable_pushkar)[0][0]), int(worksheet.get_values(masssotr.CellTable_ivanov)[0][0]), int(worksheet.get_values(masssotr.CellTable_none)[0][0])]
        # Новый массив для результата сложения
        newmass = []
        # Прибавляем новые значения к старым
        for elementfirst, elementsecond in zip(massolddates, massnewphotos):
            newmass += [elementfirst + elementsecond]
        # Обновляем значения в таблице
        worksheet.update(masssotr.CellTable_fleysner, newmass[0])
        worksheet.update(masssotr.CellTable_kireev, newmass[1])
        worksheet.update(masssotr.CellTable_pushkar, newmass[2])
        worksheet.update(masssotr.CellTable_ivanov, newmass[3])
        worksheet.update(masssotr.CellTable_none, newmass[4])

    except Exception as e:
        print(f"Логгирование статистики фотографий сломалось: {e}")

# Функция изменения Call центра
def changecallcenter():
    pythoncom.CoInitialize()
    print("Я функция изменения call-центра")
    checkupdatedatesexcel()
    datesnowmonth = importdatesformexcel(pathfile, password)
    massive = chosedates(datesnowmonth)
    result = selectmenegers(massive)
    createnewarrowincallcenter()
    print(result)

# Функция отключения всех в Call центре
def offcallcenter():
    print("Я функция полного отключения call-центра")
    # Выясняем текущий день
    today = datetime.datetime.today()
    todayday = int(today.strftime("%w"))
    if todayday == 3:
        for element in numbermanagers:
            urlforapi = urlapi + str(element) + '/agent'
            requests.put(urlforapi, params = paramoffline, headers=headers)
    else:
        return

# Класс времён
class times:
    # Время сейчас
    today = datetime.datetime.today()
    todaytime = today.strftime("%H:%M:%S")
    # Перовначальное время сканирования
    timetoScan = today.time().strftime("%H:%M")
    # Время для работы изменения Call-центра
    timetoChangeCallCenter = (today + datetime.timedelta(minutes=10)).strftime("%H:%M")
    # Время для сбора статистики по звонкам (23:30)
    timetoCollectionOfInformation = datetime.time(0, 5).strftime("%H:%M")
    # Время собрания (пока не используется)
    #timetoOffCallCenterOnMeeting = datetime.time(16, 0).strftime("%H:%M")

# Функция выбора действия от времени
def switcher(argument):
    match argument:
        # Время сканирования папки
        case times.timetoScan:
            # Сканируем папку на наличии фотографий для загрузки
            scanfolderforimages()
            # Сканируем локальные папки с фотографиями
            scanfilesinlocalserver()
            # Сканируем удалённые папки с фотографиями
            scanfilesinremoteserver()
            # Проверка различия локальной у удалённой папки
            comparisonlists()
            # Вычисление следующего времени сканирования
            nexthour = datetime.datetime.today().hour + 1
            if nexthour == 24:
                nexthour = 0
            times.timetoScan = datetime.time(nexthour, 0).strftime("%H:%M")
            print("Следующее время синхронизации фотографий:\t", times.timetoScan)
        # Время для сбора статистики call-центра
        case times.timetoCollectionOfInformation:
            # Запускаем поток с функцией сбора статистики call-центра
            t1 = Thread(target=collectionofinformation)
            t1.start()
            time.sleep(60)
        case times.timetoChangeCallCenter:
            # Запускаем поток с функцией изменения call-центра
            t2 = Thread(target=changecallcenter)
            t2.start()
            # Вычисление следующего времени сканирования
            nexthour = datetime.datetime.today().hour + 1
            if nexthour == 24:
                nexthour = 0
            times.timetoChangeCallCenter = datetime.time(nexthour, 10).strftime("%H:%M")
        # Время которое не выбрано для события
        case default:
            return print("Время сейчас:\t",argument)

# Вечный цикл с таймером 30 секунд
while True:
    # Время сейчас
    today = datetime.datetime.today()
    todaytime = today.strftime("%H:%M")
    # Запускаем функцию обработки времени
    switcher(todaytime)
    # Засыпаем функцию
    time.sleep(30)