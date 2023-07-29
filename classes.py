import time
import datetime
import os
from collections import namedtuple

from tqdm import tqdm
from ftplib import FTP
from dates import *

# Класс времён
class times:
    # Время сейчас
    today = datetime.datetime.today()
    todaytime = today.strftime("%H:%M:%S")
    # Первоначальное время сканирования
    timetoScan = today.time().strftime("%H:%M")
    #timetoScan = (today + datetime.timedelta(minutes=15)).strftime("%H:%M")
    # Время для работы изменения Call-центра
    timetoChangeCallCenter = (today + datetime.timedelta(minutes=2)).strftime("%H:%M")
    # Время для сбора статистики по звонкам
    #timetoCollectionOfInformation = datetime.time(0, 5).strftime("%H:%M")
    timetoCollectionOfInformation = (today + datetime.timedelta(minutes=5)).strftime("%H:%M")
    # Время собрания (пока не используется)
    #timetoOffCallCenterOnMeeting = datetime.time(16, 0).strftime("%H:%M")
    # Время сбора статистики по месячной работе прикрепления фотографий к карточкам товаров
    #timetoGenerationStatUploadPhotos = datetime.time(2, 30).strftime("%H:%M")
    timetoGenerationStatUploadPhotos = (today + datetime.timedelta(minutes=7)).strftime("%H:%M")
    # Время для проверки 2.0 на сканирование фотографий
    timetoScan_2_0 = datetime.time(3, 15).strftime("%H:%M")

# Класс работы с фотографиями
class class_photos(object):
    # Массив локальных файлов
    masslocal = [[], [], [], [], []]
    # Массив дат создания локальных файлов
    masslocaldates = [[], [], [], [], []]
    # Массив удалённых файлов
    massremote = [[], [], [], [], []]
    # Массив дат создания удалённых файлов
    massremotedates = [[], [], [], [], []]

    def __init__(self, argument):
        self.date = argument

    # Функция сканирования локальных папок с фотографиями
    def scanfolderforimages(self):
        print(f"Начинаем сканирование данных из локальных папок")
        # Пробегаемся по массиву и заполняем данные по называнию файлов
        for element in tqdm(os.listdir(mainpath)):
            # Формируем путь к папкам
            pathfolder = mainpath + "/" + element
            match element:
                case "1":
                    self.masslocal[0] = os.listdir(pathfolder)
                    self.masslocal[0].remove('Thumbs.db')
                case "2":
                    self.masslocal[1] = os.listdir(pathfolder)
                    self.masslocal[1].remove('Thumbs.db')
                case "3":
                    self.masslocal[2] = os.listdir(pathfolder)
                    self.masslocal[2].remove('Thumbs.db')
                case "4":
                    self.masslocal[3] = os.listdir(pathfolder)
                    self.masslocal[3].remove('Thumbs.db')
                case "5":
                    self.masslocal[4] = os.listdir(pathfolder)
                    self.masslocal[4].remove('Thumbs.db')
                case _:
                    continue
        # Если время для продвинутого сканирования, запускаем
        if self.date == times.timetoScan_2_0:
            print(f"Начинаем сканирование данных времени из локальных папок")
            # Пробегаемся по сформированному массиву, чтобы извлечь данные времени создания
            for element in tqdm(self.masslocal):
                # Вычисляем номер папки
                numberfolder = self.masslocal.index(element) + 1
                # Запускам цикл по фотографиям, которые находятся в папках
                for elem in element:
                    # Формируем путь к элементу
                    pathphoto = mainpath + str(numberfolder) + "/" + str(elem)
                    # Вычисляем дату создания фотографии
                    datephoto = os.stat(pathphoto).st_ctime
                    # Добавляем данные по результатам в массив
                    dateandtime = datetime.datetime.fromtimestamp(datephoto).strftime('%Y %m %d %H:%M:%S')
                    self.masslocaldates[numberfolder-1].append(dateandtime)

    # Функция сканирования удалённых папкок с фотографиями
    def scanfilesinremoteserver(self):
        print(f"Начинаем сканирование данных из удалённых папок")
        # Инициализируем попытку сбора данных с удалённого сервера
        try:
            # Открываем связь с удалённым сервером
            datesftp = FTP(nameSite)
            datesftp.login(ftpLogin, ftpPass)
            # Получаем данные о том какие данные есть на удалённом сервере
            listalldirectors = datesftp.nlst()
            for element in tqdm(listalldirectors):
                match element:
                    # Первая папка для синхронизации
                    case "1":
                        self.massremote[0] = self.importatesfromftp(datesftp, element)
                    # Вторая папка для синхронизации
                    case "2":
                        self.massremote[1] = self.importatesfromftp(datesftp, element)
                    # Третья папка для синхронизации
                    case "3":
                        self.massremote[2] = self.importatesfromftp(datesftp, element)
                    # Четвёртая папка для синхронизации
                    case "4":
                        self.massremote[3] = self.importatesfromftp(datesftp, element)
                    # Пятая папка для синхронизации
                    case "5":
                        self.massremote[4] = self.importatesfromftp(datesftp, element)
                    case _:
                        continue
            # Закрываем соединение с удалённым сервером
            datesftp.close()
            # Если время для продвинутого сканирования, запускаем
            if self.date == times.timetoScan_2_0:
                # Пробегаемся по сформированному массиву, чтобы извлечь данные времени создания
                for element in tqdm(self.massremote):
                    # Открываем связь с удалённым сервером
                    datesftp = FTP(nameSite)
                    datesftp.login(ftpLogin, ftpPass)
                    # Вычисляем номер папки
                    numberfolder = self.massremote.index(element) + 1
                    # Запускам цикл по фотографиям, которые находятся в папках
                    for elem in element:
                        # Формируем путь к элементу
                        remotepathphoto = "/" + str(numberfolder) + "/" + str(elem)
                        self.importremotedatesfromftp(datesftp, remotepathphoto, numberfolder)
                        # Вычисляем дату создания фотографии
                        #datephoto = os.stat(pathphoto).st_ctime
                        # Добавляем данные по результатам в массив
                        #self.masslocaldates[numberfolder - 1].append(time.ctime(datephoto))
                    # Закрываем соединение с удалённым сервером
                    datesftp.close()
        except Exception as e:
            print("Синхронизация папок не удалась. Попробуем в следующий раз.")
            print(f"\t{e}")
            return

    # Функция импорта данных
    def importatesfromftp(self, datesftp, element):
        # Определяем путь для папки
        path = "/" + str(element) + "/"
        # Изменение каталог работы
        datesftp.cwd(path)
        # Получаем лист всех фаилов из папки
        list = datesftp.nlst()
        # Удалем первые 2 элемента (так как на сервере система Linux)
        list.pop(0)
        list.pop(0)
        # Сортируем фаилы по возрастанию
        list.sort()
        # Возвращаем полученный список
        return list

    # Функция импорта данных дат файлов на удалённом сервере
    def importremotedatesfromftp(self, datesftp, remotepathphoto, numberfolder):
        cmdrequest = "MDTM " + remotepathphoto
        liastmodifited = datesftp.voidcmd(cmdrequest)[4:].strip()
        year = liastmodifited[:4]
        month = liastmodifited[4:6]
        day = liastmodifited[6:8]
        hour = liastmodifited[8:10]
        minutes = liastmodifited[10:12]
        seconds = liastmodifited[12:14]
        dateandtime = year + " " + month + " " + day + " " + hour + ":" + minutes + ":" + seconds
        self.massremotedates[numberfolder-1].append(dateandtime)
        return dateandtime

    # Функция различия локальной у удалённой папки
    def comparisonlists(self):
        for element in range(0, 5):
            result = list(set(self.masslocal[element]) - set(self.massremote[element]))
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
                match element:
                    case 0:
                        print("Разность первых папкок: ", result)
                        self.uploadfiles(element + 1, result)
                    case 1:
                        print("Разность вторых папкок: ", result)
                        self.uploadfiles(element + 1, result)
                    case 2:
                        print("Разность третьих папкок: ", result)
                        self.uploadfiles(element + 1, result)
                    case 3:
                        print("Разность четвёртых папкок: ", result)
                        self.uploadfiles(element + 1, result)
                    case 4:
                        print("Разность пятых папкок: ", result)
                        self.uploadfiles(element + 1, result)

    # Функция загрузки фотографий на сервер
    def uploadfiles(self, numberfolder, result):
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