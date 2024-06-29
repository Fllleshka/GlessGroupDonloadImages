import unittest
from classes import *

# Тест класса отправки сообщений в телеграмм бота
class Test_class_send_erorr_message(unittest.TestCase):

    # Проверяем функцию отправки сообщений
    def test_send_message(self):
        testmessage = class_send_erorr_message('11:16', 'Тестовая функция', 'exeption')
        self.assertEqual(testmessage.send_message(), 'Возникла проблема с функцией: Тестовая функция [11:16]\nОшибка типа:\n{exeption}\n')

if __name__ == '__main__':
    unittest.main()