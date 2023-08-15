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
            printer(times.timetoScan, "Функция сканирования папки с фотографиями")

            # Инициализация класса
            x = class_photos(argument)
            t0 = Thread(target = x.startprocessing())
            t0.start()

            # Вычисление следующего времени изменения Call центра
            if argument == times.timetoScan_2_0:
                times.timetoChangeCallCenter = datetime.time(5, 10).strftime("%H:%M")

            nexthour = datetime.datetime.today().hour + 1
            if nexthour == 24:
                nexthour = 0
            times.timetoScan = datetime.time(nexthour, 0).strftime("%H:%M")
            print("Следующее время синхронизации фотографий:\t", times.timetoScan)

        # Время для сбора статистики call-центра
        case times.timetoCollectionOfInformation:
            printer(times.timetoCollectionOfInformation, "Время импорта статистики по звонкам")

            # Запускаем поток с функцией сбора статистики call-центра
            t1 = Thread(target=collectionofinformation)
            t1.start()

            times.timetoCollectionOfInformation = datetime.time(0, 5).strftime("%H:%M")
            print("Следующее время для сбора статистики по звонкам\t", times.timetoCollectionOfInformation)

        # Время для изменения call-центра
        case times.timetoChangeCallCenter:
            printer(times.timetoChangeCallCenter, "Функция изменения call центра")
            # Запускаем поток с функцией изменения call-центра
            # Инициализация класса
            x = class_call_center(argument)
            # Запускам поток который вызывает стартовую функцию
            t2 = Thread(target = x.changecallcenter())
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
            times.timetoGenerationStatUploadPhotos = datetime.time(0, 15).strftime("%H:%M")
            print("Следующее время для подведения статистики по загруженным фотографиям\t", times.timetoChangeCallCenter)

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