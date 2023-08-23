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