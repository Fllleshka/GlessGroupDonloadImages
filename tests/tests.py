import unittest
from classes import *

# Тест класса отправки сообщений в телеграмм бота
class Test_class_send_erorr_message(unittest.TestCase):

    # Проверяем функцию отправки сообщений
    def test_send_message(self):
        # Инициализация класса
        testmessage = class_send_erorr_message('11:16', 'Тестовая функция', 'exeption')
        # Запуск тестирования функции
        self.assertEqual(testmessage.send_message(), 'Возникла проблема с функцией: Тестовая функция [11:16]\nОшибка типа:\n{exeption}\n')

    # Проверяем функцию сохранения статистики по загруженным фотографиям
    def test_generationstatuploadphotos(self):
        # Инициализация класса
        testmessage = class_generation_stat_uploadphotos('11:16')
        # Запуск тестирования функции
        self.assertEqual(testmessage.generationstatuploadphotos(), "\tВремя для обнуления ещё не пришло.")

if __name__ == '__main__':
    unittest.main()