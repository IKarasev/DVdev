# coding: utf8
""" Testing-comaprison of utility and svod reports as excel files
By Karasev I.A.
"""

import xlrd
import xlwt
import sys


"""Константы"""
UTIL_COLS = {
	'VAGON_NUMBER': 8,
	'IN_DATE': 4,
	'TYPE': 19,
	'PART_NUMBER': 16,
	'GRADATION': 20,
	'NAME_MC': 7,
	'SCEP': 1
}

SVOD_COLS = {
	'VAGON_NUMBER': 5,
	'IN_DATE': 8,
	'PART_NUMBER': 10,
	'TYPE': 9,
	'WIDTH': 14,
	'COMMENT': 28,
	"GRADATION": 29,
	'SCEP': 30
}

PART_TYPE = {
	'КП': 'Колесная пара',
	'БР': 'Боковая рама',
	'НБ': 'Надрессорная балка',

}


def main():
	"""Функция выполняемая при запуске скрипта"""
	print("Сравнение открыто")


def get_excel_file():
	"""Запрашивает путь к файлу excel и возвращает книгу в виде xlrd"""
	filepath = input("Введите путь к файлу: ")
	wb = xlrd.open_workbook(filepath,formatting_info=True)
	print(type(wb))
	print("Файл %s загружен"%(filepath))
	return wb


def compare_row(util_row, svod_row):
	"""Производит сравнение строк и возвращает True если критерии
	удволетворены и False если нет"""
	return True


def analyse_files():
	"""Проводит сравнительный анализ двух файлов и записывает результат
	в svod"""
	print("Загрузите сводную таблицу")
	svod = get_excel_file()
	print("Загрузите утилиту")
	util = get_excel_file()


main()