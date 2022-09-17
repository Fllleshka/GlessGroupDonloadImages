import datetime
import time
import os
from PIL import Image
from ftplib import FTP
import shutil

masslocalfirst = []
masslocalsecond = []
masslocalthird = []
masslocalfourth = []
masslocalfifth = []

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
        # Формируем поуть к папке
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

def scanfilesinremoteserver():
    global massremotefirst
    global massremotesecond
    global massremotethird
    global massremotefourth
    global massremotefifth
    ftp = FTP('gless.group')
    result = ftp.login('mazda_import', '{(U9Av74{Dn?')
    listalldirectors = ftp.nlst()
    listdirectors = []
    for element in listalldirectors:
        match element:
            case "1":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotefirst = list
            case "2":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotesecond = list
            case "3":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotethird = list
            case "4":
                listdirectors.append(int(element))
                path = "/" + str(element) + "/"
                ftp.cwd(path)
                list = ftp.nlst()
                list.pop(0)
                list.pop(0)
                list.sort()
                massremotefourth = list
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

def renameanduploadimage(pathimage, folder):
    # Начинаем с переименования картинки
    numberfolderfirst = str(pathimage)[53:]
    #print("numberfolderfirst:", numberfolderfirst)
    # Если папка четырёхзначная
    if numberfolderfirst[4] == "/":
        numberfoldersecond = str(numberfolderfirst)[:4]
    elif numberfolderfirst[3] == "/":
        numberfoldersecond = str(numberfolderfirst)[:3]
    else:
        numberfoldersecond = str(numberfolderfirst)[:5]
    #print("\tСтарный путь к картинке: \t", pathimage)
    # Название картинки
    namepic = numberfoldersecond + str(pathimage)[-4:]
    #print("\tНазвание картинки: \t", namepic)
    # Новый путь к картинке
    convertname = str(pathimage)[:53] + numberfoldersecond + "/" + namepic
    #print("\tНовый путь к картинке: \t\t", convertname)
    # Переименование картинки
    os.rename(pathimage, convertname)
    #print("Переименование завершено")

    # Начинаем загрузку фотографии по необходимому местоположению
    newpathfile = str(pathimage)[:29] + str(folder) + "/" + namepic
    newlocation = shutil.move(convertname, newpathfile)
    #print("Загружено: \t", newlocation)

def uploadfiles(numberfolder, result):
    # Главный путь к папкам
    mainpath = '//192.168.20.215/фото товара/'
    pathtofolder = mainpath + str(numberfolder) + "/"
    listphotos = os.listdir(pathtofolder)

    ftp = FTP('gless.group')
    ftp.login('mazda_import', '{(U9Av74{Dn?')
    ftppath = "/" + str(numberfolder) + "/"
    ftp.cwd(ftppath)

    for element in result:
        path = pathtofolder + element
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

def scanfolderforimages():
    # Путь к главной папке
    mainpath = '//192.168.20.215/фото товара/фото товара для разбора'
    list = os.listdir(mainpath)
    if list == []:
        print("\tДанные по импорту фотографий из папки для разбора отсутствуют")
    else:
        for element in list:
            pathfolder = mainpath + "/" + element
            nextlist = os.listdir(pathfolder)
            #print(pathfolder, "\t", nextlist)
            # Если папка пуста то пишем о пустой папке
            if nextlist == []:
                print("\t Папка ", element, " пуста")
            else:
                numberfolder = 1
                for elem in nextlist:
                    if elem == "Thumbs.db":
                        continue
                    else:
                        #print("Работаю с", elem)
                        pathimage = pathfolder + "/" + elem
                        convertimage(pathimage)
                        renameanduploadimage(pathimage, numberfolder)
                        numberfolder = numberfolder + 1
    # После окончания загрузки фотографий по папкам удаляем папку
    for elem in list:
        path = mainpath + "/" + elem
        shutil.rmtree(path)
    print("Удаление папок завершено")

class times:
    today = datetime.datetime.today()
    todaytime = today.strftime("%H:%M:%S")
    timetoScan = datetime.time(16, 13).strftime("%H:%M")

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
        case default:
            return print("Время сейчас:\t",argument)

while True:
    # Время сейчас
    today = datetime.datetime.today()
    todaytime = today.strftime("%H:%M")
    switcher(todaytime)
    time.sleep(60)
