from helperscripts import *
from classes import *

# Функция выбора действия от времени
def switcher(argument):
    match argument:
        # Время для сканирования папки
        case times.timetoScan:
            text = "Функция сканирования папки с фотографиями"
            printer(times.timetoScan, text)

            try:
                # Инициализация класса
                x = class_photos(argument, timetowaitingfunction)
                # Запускаем поток с функцией разбора и синхронизации фотографий
                t0 = Thread(target = x.startprocessing())
                t0.start()

                # Вычисление следующего времени изменения Call центра
                if argument == times.timetoScan_2_0:
                    nexthour2 = datetime.datetime.today().hour + 3
                    times.timetoChangeCallCenter = datetime.time(nexthour2, 0).strftime("%H:%M")

                nexthour = datetime.datetime.today().hour + 1
                if nexthour == 24:
                    nexthour = 0
                times.timetoScan = datetime.time(nexthour, 0).strftime("%H:%M")
                print("Следующее время синхронизации фотографий:\t", times.timetoScan)
            except Exception as exception:
                # Инициализация класса
                error_message = class_send_erorr_message(argument, text, exception, botkey)
                # Функция отправки сообщения в чат системному администратору
                error_message.send_message()

        # Время для сбора статистики call-центра
        case times.timetoCollectionOfInformation:
            text = "Время импорта статистики по звонкам"
            printer(times.timetoCollectionOfInformation, text)

            try:
                # Инициализация класса
                x = class_collecion_of_information(argument)
                # Запускаем поток с функцией сбора статистики call-центра
                t1 = Thread(target=x.collectionofinformation)
                t1.start()
                times.timetoCollectionOfInformation = datetime.time(0, 5).strftime("%H:%M")
                print("Следующее время для сбора статистики по звонкам\t", times.timetoCollectionOfInformation)
            except Exception as exception:
                # Инициализация класса
                error_message = class_send_erorr_message(argument, text, exception, botkey)
                # Функция отправки сообщения в чат системному администратору
                error_message.send_message()

        # Время для изменения call-центра
        case times.timetoChangeCallCenter:
            text = "Функция изменения call центра"
            printer(times.timetoChangeCallCenter, text)

            try:
                # Инициализация класса
                x = class_call_center(argument)
                # Запускаем поток с функцией изменения call-центра
                t2 = Thread(target = x.changecallcenter)
                t2.start()
                # Вычисление следующего времени сканирования
                nexthour = datetime.datetime.today().hour + 1
                if nexthour == 24:
                    nexthour = 0
                times.timetoChangeCallCenter = datetime.time(nexthour, 10).strftime("%H:%M")
                print("Следующее время для работы изменения Call-центра\t", times.timetoChangeCallCenter)

            except Exception as exception:
                # Инициализация класса
                error_message = class_send_erorr_message(argument, text, exception, botkey)
                # Функция отправки сообщения в чат системному администратору
                error_message.send_message()

        # Время для сбора статистики по загруженным фотографиям
        case times.timetoGenerationStatUploadPhotos:
            text = "Функция сбора статистики по загруженным фотографиям"
            printer(times.timetoChangeCallCenter, text)

            try:
                # Инициализация класса
                x = class_generation_stat_uploadphotos(argument)
                # Запускаем поток с функцией подсчёта статистики загруженных фотографий
                t3 = Thread(target=x.generationstatuploadphotos)
                t3.start()

                times.timetoGenerationStatUploadPhotos = datetime.time(2, 30).strftime("%H:%M")
                print("\tСледующее время для подведения статистики по загруженным фотографиям\t", times.timetoChangeCallCenter)
            except Exception as exception:
                # Инициализация класса
                error_message = class_send_erorr_message(argument, text, exception, botkey)
                # Функция отправки сообщения в чат системному администратору
                error_message.send_message()

        # Время проверки прайс лист на обновление
        case times.timetoScanUpdatePrise:
            text = "Функция проверка прайс листа"
            printer(times.timetoChangeCallCenter, text)

        # Время которое не выбрано для события
        case default:
            return print("Время сейчас:\t", default)

# Вечный цикл с таймером 30 секунд
if __name__ == '__main__':
    while True:
        # Время сейчас
        today = datetime.datetime.today()
        todaytime = today.strftime("%H:%M")
        # Запускаем функцию обработки времени
        switcher(todaytime)
        # Засыпаем функцию
        time.sleep(30)