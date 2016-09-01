# coding: utf8
""" Testing-comaprison of utility and svod reports as excel files
By Karasev I.A.
"""

import xlrd
import xlwt
import sys


SVOD_VAGON_N_COL = 5
SVOD_IN_DATE_COL = 8
UTIL_VAGON_N_COL = 8
UTIL_IN_DATE_COL = 4
UTIL_TYPE_COL = 19



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

	print 

main()