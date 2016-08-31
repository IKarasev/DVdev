# coding: utf8
""" Testing-comaprison of utility and svod reports as excel files
By Karasev I.A.
"""

import xlrd
import xlwt
import sys

def main():
	"""Функция выполняемая при запуске скрипта"""
	print("Сравнение открыто")


def get_excel_file():
	"""Запрашивает путь к файлу excel и возвращает книгу в виде xlrd"""
	filepath = input("Введите путь к файлу: ")
	wb = xlrd.open_workbook(filepath,formatting_info=True)
	sheet = wb.sheet_by_index(0)
	print(type(sheet))
	print("Файл %s загружен"%(filepath))
	return sheet


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