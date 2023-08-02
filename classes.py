import datetime
import os
import gspread
import win32security
import shutil

from tqdm import tqdm
from ftplib import FTP
from dates import *
from PIL import Image
from threading import Thread

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
    timetoScan_2_0 = datetime.time(3, 0).strftime("%H:%M")

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
    # Массив загружаемых фотографий
    massnewphotos = [0, 0, 0, 0, 0]

    def __init__(self, argument):
        self.date = argument

    # Функция записи логов папок фотографий
    def createnewarrowinlogs(self, lenphotos):
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

    # Функция сбора статистики по загруженным фотографиям
    def statisticsphotos(self, pathimage, massnewphotos):
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

    # Функция получения размера изображения
    def get_size_format(self, b, factor=1024, suffix="B"):
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
    def convertimage(self, path):
        # Размеры изображения на выходе
        width = 1920
        height = 1440
        # Загружаем фотографию в память
        img = Image.open(path)
        # Первоначальный размер картинки
        olddimensions = img.size
        # Получаем размер изображения до компрессии
        image_size = os.path.getsize(path)
        oldsize = self.get_size_format(image_size)
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
        newsize = self.get_size_format(image_size)
        # Печатаем в кносоль результат
        print(f"{path} с шириной, высотой: {olddimensions} и размером: {oldsize} была преобразована в: {newdimesions} и {newsize}")

    # Переименование и перемещение картинки по необходимому локальному пути
    def renameanduploadimage(self, pathimage, folder):
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

    # Функция записи статистики по загруженным фотографиям в GoogleDocs
    def updatedatesuploadphotos(self, massnewphotos):
        try:
            # Подключаемся к сервисному аккаунту
            gc = gspread.service_account(CREDENTIALS_FILE)
            # Подключаемся к таблице по ключу таблицы
            table = gc.open_by_key(sheetkey)
            # Открываем нужный лист
            worksheet = table.worksheet("LogsPhotos")
            # Получаем данные из ячеек
            massolddates = [int(worksheet.get_values(masssotr.CellTable_fleysner)[0][0]),
                            int(worksheet.get_values(masssotr.CellTable_kireev)[0][0]),
                            int(worksheet.get_values(masssotr.CellTable_pushkar)[0][0]),
                            int(worksheet.get_values(masssotr.CellTable_ivanov)[0][0]),
                            int(worksheet.get_values(masssotr.CellTable_none)[0][0])]
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

    # Функция сканирование папки для разбора фотографий
    def scanfolderwithimages(self):
        # Массив загруженных фотографий
        massnewphotos = [0, 0, 0, 0, 0]
        # Получаем лист фаилов находящихся по адресу
        list = os.listdir(mainpathanalysis)
        # Проверка наличия фотографий
        # Если папок для разбора нет
        if list == []:
            print("\tДанные по импорту фотографий из папки для разбора отсутствуют")
        # Если есть папки для разбора
        else:
            # Выясняем количество папок
            lenphotos = len(list)
            # Логгирование фотографий
            self.createnewarrowinlogs(lenphotos)
            # Пробегаемся по элеметам
            for element in list:
                # Обыгрывание Thumbs.db
                if element == "Thumbs.db":
                    # Выясняем путь к этому фаилу
                    path = mainpathanalysis + "/" + element
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
                    pathfolder = mainpathanalysis + "/" + element
                    # Получаем данные о фаилах по этому пути
                    nextlist = os.listdir(pathfolder)
                    # Если папка пуста, то пишем о пустой папке
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
                                    massnewphotos = self.statisticsphotos(pathimage, massnewphotos)
                                    # Уменьшение веса и подгонка фотографии
                                    self.convertimage(pathimage)
                                    # Переименоваие и загрузка фотографии
                                    self.renameanduploadimage(pathimage, numberfolder)
                                    # Увеличиваем счётчик
                                    numberfolder = numberfolder + 1
            # После окончания загрузки фотографий по папкам удаляем папку
            for elem in list:
                # Обыгрывание Thumbs.db
                if element == "Thumbs.db":
                    path = mainpathanalysis + "/" + element
                    try:
                        if os.access(path, os.R_OK and os.X_OK):
                            os.remove(path)
                    except PermissionError:
                        pass
                else:
                    # Выясняем путь к папке
                    path = mainpathanalysis + "/" + elem
                    # Удаляем полностью папку
                    try:
                        shutil.rmtree(path)
                    except Exception as e:
                        print("Удаление невозможно. По причине ", e)
            t1 = Thread(target=self.updatedatesuploadphotos, args=(massnewphotos,))
            t1.start()
            print("Удаление папок завершено")

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