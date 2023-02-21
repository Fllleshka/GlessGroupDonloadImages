from ftplib import FTP
import shutil
from helperscripts import *
from threading import Thread
import pythoncom

#locallist = []

masslocalfirst = []
masslocalsecond = []
masslocalthird = []
masslocalfourth = []
masslocalfifth = []
massremotefirst = []
massremotesecond = []
massremotethird = []
massremotefourth = []
massremotefifth = []

# Функция формирования путей до папок на сервере картинок
def scanfilesinlocalserver():

    #global locallist

    global masslocalfirst
    global masslocalsecond
    global masslocalthird
    global masslocalfourth
    global masslocalfifth
    i = 0
    # Главный путь к папкам
    mainpath = '//192.168.20.215/фото товара'
    # Пробегаемся по массиву
    for element in os.listdir(mainpath):
        # Формируем путь к папкам
        pathfolder = mainpath + "/" + element
        match element:
            case "1":
                #locallist.index(0) = os.listdir(pathfolder)
                #locallist.append(os.listdir(pathfolder))
                #print(locallist[0])
                #locallist[0].remove('Thumbs.db')

                masslocalfirst = os.listdir(pathfolder)
                masslocalfirst.remove('Thumbs.db')
            case "2":
                masslocalsecond = os.listdir(pathfolder)
                masslocalsecond.remove('Thumbs.db')
            case "3":
                masslocalthird = os.listdir(pathfolder)
                masslocalthird.remove('Thumbs.db')
            case "4":
                masslocalfourth = os.listdir(pathfolder)
                masslocalfourth.remove('Thumbs.db')
            case "5":
                masslocalfifth = os.listdir(pathfolder)
                masslocalfifth.remove('Thumbs.db')
            case _:
                continue

# Функция сканируемых данных на удалённом сервере
def scanfilesinremoteserver():
    global massremotefirst
    global massremotesecond
    global massremotethird
    global massremotefourth
    global massremotefifth
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
                    massremotefirst = importatesfromftp(ftp, listdirectors, element)
                # Вторая папка для синхронизации
                case "2":
                    massremotesecond = importatesfromftp(ftp, listdirectors, element)
                # Третья папка для синхронизации
                case "3":
                    massremotethird = importatesfromftp(ftp, listdirectors, element)
                # Четвёртая папка для синхронизации
                case "4":
                    massremotefourth = importatesfromftp(ftp, listdirectors, element)
                # Пятая папка для синхронизации
                case "5":
                    massremotefifth = importatesfromftp(ftp, listdirectors, element)
                case _:
                    continue
        listdirectors.sort()
    except:
        print("Синхронизация папок не удалась. Попробуем в следующий раз")
        return

# Проверка на разность данных в локальных папках и на удалённом сервере
def comparisonlists():

    # Первая папка
    #result = list(set(locallist[1]) - set(massremotefirst))
    result = list(set(masslocalfirst) - set(massremotefirst))
    if result == []:
        print("Первые\t\tпапки синхронизированны!")
    else:
        print("Разность первых папкок: ", result)
        uploadfiles(1, result)

    # Вторая папка
    result = list(set(masslocalsecond) - set(massremotesecond))
    if result == []:
        print("Вторые\t\tпапки синхронизированны!")
    else:
        print("Разность вторых папкок: ", result)
        uploadfiles(2, result)

    # Третья папка
    result = list(set(masslocalthird) - set(massremotethird))
    if result == []:
        print("Третьи\t\tпапки синхронизированны!")
    else:
        print("Разность третьих папкок: ", result)
        uploadfiles(3, result)

    # Четвёртая папка
    result = list(set(masslocalfourth) - set(massremotefourth))
    if result == []:
        print("Четвёртые\tпапки синхронизированны!")
    else:
        print("Разность четвёртых папкок: ", result)
        uploadfiles(4, result)

    # Пятая папка
    result = list(set(masslocalfifth) - set(massremotefifth))
    if result == []:
        print("Пятые\t\tпапки синхронизированны!")
    else:
        print("Разность пятых папкок: ", result)
        uploadfiles(5, result)

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
                # Выясняем путь к фаилу
                path = mainpath + "/" + elem
                # Удаляем полностью папку
                try:
                    shutil.rmtree(path)
                except Exception as e:
                    print("Удаление невозможно. По причине ", e)

        print("Удаление папок завершено")

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
    timetoChangeCallCenter = datetime.time(19, 5).strftime("%H:%M")
    # Время для сбора статистики по звонкам (01:00)
    #timetoCollectionOfInformation = datetime.time(1, 0).strftime("%H:%M")
    timetoCollectionOfInformation = (datetime.datetime.today() + datetime.timedelta(seconds=10)).strftime("%H:%M")
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
            # Запускаем поток с функцией изменения call-центра
            t1 = Thread(target=changecallcenter)
            t1.start()
            print("Следующее время проверки:\t", times.timetoScan)
            collectionofinformation()
        case times.timetoCollectionOfInformation:
            print()
            #collectionofinformation()

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