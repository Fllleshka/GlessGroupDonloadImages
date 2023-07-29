import datetime
import time
from ftplib import FTP
from helperscripts import *
from classes import *
from threading import Thread
import pythoncom
import win32security

masslocal = [[], [], [], [], []]
massremote = [[], [], [], [], []]

# Переименование и перемещение картинки по необходимому локальному пути
def renameanduploadimage(pathimage, folder):
    lenmailfolder = 62
    # Начинаем с переименования картинки
    numberfolderfirst = str(pathimage)[lenmailfolder:]
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
    convertname = str(pathimage)[:lenmailfolder] + numberfoldersecond + "/" + namepic
    # Переименование картинки
    os.rename(pathimage, convertname)

    # Начинаем загрузку фотографии по необходимому местоположению
    newpathfile = str(pathimage)[:29] + str(folder) + "/" + namepic
    shutil.move(convertname, newpathfile)

# Функция загрузки фотографий на сервер
def uploadfiles(numberfolder, result):
    # Путь к элементу
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

# Функция выбора действия от времени
def switcher(argument):
    match argument:
        # Время сканирования папки
        case times.timetoScan:
            # Инициализация класса
            x = class_photos(argument)
            # Вызов функции сканирования локальных папок
            x.scanfolderforimages()
            # Вызов функции сканирования удалённых папок
            x.scanfilesinremoteserver()
            # Вызов функции выявления различия файлов на локальном и удалённом сервере
            x.comparisonlists()

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
            times.timetoCollectionOfInformation = datetime.time(0, 5).strftime("%H:%M")
            print("Следующее время для сбора статистики по звонкам\t", times.timetoCollectionOfInformation)
        case times.timetoChangeCallCenter:
            # Запускаем поток с функцией изменения call-центра
            t2 = Thread(target=changecallcenter)
            t2.start()
            # Вычисление следующего времени сканирования
            nexthour = datetime.datetime.today().hour + 1
            if nexthour == 24:
                nexthour = 0
            times.timetoChangeCallCenter = datetime.time(nexthour, 10).strftime("%H:%M")
            print("Следующее время для работы изменения Call-центра\t", times.timetoChangeCallCenter)
        case times.timetoGenerationStatUploadPhotos:
            # Запускаем поток с функцией подсчёта статистики загруженных фотографий
            t3 = Thread(target=generationstatuploadphotos)
            t3.start()
            time.sleep(60)
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