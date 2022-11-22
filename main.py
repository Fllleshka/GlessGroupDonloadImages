import time
import os
from PIL import Image
from ftplib import FTP
import shutil
from dates import *
from helperscripts import *

masslocalfirst = []
masslocalsecond = []
masslocalthird = []
masslocalfourth = []
masslocalfifth = []

# Функция формирования путей до папок на сервере картинок
def scanfilesinlocalserver():
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

massremotefirst = []
massremotesecond = []
massremotethird = []
massremotefourth = []
massremotefifth = []

# Фукнкция сканируемых данных на удалённом сервере
def scanfilesinremoteserver():
    global massremotefirst
    global massremotesecond
    global massremotethird
    global massremotefourth
    global massremotefifth
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
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotefirst = list
            # Вторая папка для синхронизации
            case "2":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotesecond = list
            # Третья папка для синхронизации
            case "3":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotethird = list
            # Четвёртая папка для синхронизации
            case "4":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotefourth = list
            # Пятая папка для синхронизации
            case "5":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotefifth = list
            case _:
                continue
    listdirectors.sort()

# Проверка на разность данных в локальных папках и на удалённом сервере
def comparisonlists():
    # Первая папка
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
    img.thumbnail(size = (width, height))
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
    print(path, "с шириной, высотой: ", olddimensions, " и размером: ", oldsize, "была преобразована в: ", newdimesions , " и ", newsize)

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
    # Путь к главной папке
    mainpath = '//192.168.20.215/фото товара/фото товара для разбора'
    list = os.listdir(mainpath)
    if list == []:
        print("\tДанные по импорту фотографий из папки для разбора отсутствуют")
    else:
        for element in list:
            # Обыгрывание Thumbs.db решение удалить пока не найдено(
            if element == "Thumbs.db":
                path = mainpath + "/" + element
                try:
                    if os.access(path, os.R_OK and os.X_OK):
                        os.remove(path)
                except PermissionError:
                    pass
            else:
                pathfolder = mainpath + "/" + element
                nextlist = os.listdir(pathfolder)
                # Если папка пуста то пишем о пустой папке
                if nextlist == []:
                    print("\t Папка ", element, " пуста")
                else:
                    numberfolder = 1
                    for elem in nextlist:
                        if elem == "Thumbs.db":
                            continue
                        else:
                            # Условие перебора количества фотографий, так как 6 папки нету
                            if numberfolder >= 6:
                                break
                            else:
                                pathimage = pathfolder + "/" + elem
                                convertimage(pathimage)
                                renameanduploadimage(pathimage, numberfolder)
                                numberfolder = numberfolder + 1
    # После окончания загрузки фотографий по папкам удаляем папку
    for elem in list:
        # Обыгрывание Thumbs.db решение удалить пока не найдено(
        if element == "Thumbs.db":
            path = mainpath + "/" + element
            try:
                if os.access(path, os.R_OK and os.X_OK):
                    os.remove(path)
            except PermissionError:
                pass
        else:
            path = mainpath + "/" + elem
            shutil.rmtree(path)
    print("Удаление папок завершено")

# Функция изменения Call центра
def changecallcenter():
    print("Я функция изменения call-центра")
    pathfile = u"//192.168.20.215/сетевой диск2/График взаимодействия/График 2022 ТЕСТ.xlsx"
    password = "888"
    datesnowmonth = importdatesformexcel(pathfile, password)
    chosedates(datesnowmonth)

# Класс времён
class times:
    today = datetime.datetime.today()
    todaytime = today.strftime("%H:%M:%S")
    timetoScan = today.time().strftime("%H:%M")
    timetoChangeCallCenter = datetime.time(19, 5).strftime("%H:%M")
    #importdatesfromexcel = datetime.datetime.today().time().strftime("%H:%M")

# Функция выбора действия от времени
def switcher(argument):
    match argument:
        case times.timetoScan:
            scanfolderforimages()
            scanfilesinlocalserver()
            scanfilesinremoteserver()
            comparisonlists()
            nexthour = datetime.datetime.today().hour + 1
            if nexthour == 24:
                nexthour = 0
            times.timetoScan = datetime.time(nexthour, 0).strftime("%H:%M")
            print("Следующее вермя проверки:\t", times.timetoScan)
            changecallcenter()
        case default:
            return print("Время сейчас:\t",argument)

# Вечный цикл с таймером 60 секунд
while True:
    # Время сейчас
    today = datetime.datetime.today()
    todaytime = today.strftime("%H:%M")
    switcher(todaytime)
    time.sleep(60)
