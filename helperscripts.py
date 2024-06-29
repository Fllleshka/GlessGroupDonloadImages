import datetime
import requests

from dates import *

# Функция вывода данных о запущенной функции
def printer(time, name):
    print(f"\t{time}\t\t{name}.")

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

# Функция сохраниения фаилов в папки менеджеров
def savefileinfolder(message, bot):
    match message.chat.id:
        case userid.id_konovalov:
            bot.send_message(message.chat.id, "Готов к приёму файлов", reply_markup = types.ReplyKeyboardRemove())
            bot.register_next_step_handler(message, savefileinfolder2, bot, class_pathmanagers.konovalov)
        case userid.id_zagravskiy:
            bot.send_message(message.chat.id, "Готов к приёму файлов", reply_markup=types.ReplyKeyboardRemove())
            bot.register_next_step_handler(message, savefileinfolder2, bot, class_pathmanagers.zagravskiy)
        case userid.id_beregovoy:
            bot.send_message(message.chat.id, "Готов к приёму файлов", reply_markup=types.ReplyKeyboardRemove())
            bot.register_next_step_handler(message, savefileinfolder2, bot, class_pathmanagers.beregovoy)
        case _:
            text = "Вы не имеете доступа к данной функции"
            print(text)
            bot.send_message(message.chat.id, text)