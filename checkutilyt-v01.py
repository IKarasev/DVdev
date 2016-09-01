# coding: utf8
""" Testing-comaprison of utility and svod reports as excel files
By Karasev I.A.
"""

import xlrd
import xlwt
import datetime
from xlutils.copy import copy


"""Константы"""
UTIL_DATA = {
	'VAGON_NUMBER': 7,
	'IN_DATE': 3,
	'TYPE': 18,
	'PART_NUMBER': 17,
	'GRADATION': 19,
	'NAME_MC': 6,
	'SCEP': 0,
}

UTIL_ROW_START = 2

SVOD_DATA = {
	'VAGON_NUMBER': 4,
	'IN_DATE': 7,
	'PART_NUMBER': 9,
	'TYPE': 8,
	'WIDTH': 13,
	'COMMENT': 27,
	"GRADATION": 28,
	'SCEP': 29
}

SVOD_ROW_START = 3

PART_TYPE = {
	'КП': 'Колесная пара',
	'БР': 'Боковая рама',
	'НБ': 'Надрессорная балка',

}

DATE_FRAME = 3


def main():
	"""Функция выполняемая при запуске скрипта"""
	print("Сравнение открыто")
	analyse_files()


def get_excel_file():
	"""Запрашивает путь к файлу excel и возвращает книгу в виде xlrd"""
	filepath = input("Введите путь к файлу: ")
	wb = xlrd.open_workbook(filepath, formatting_info=True, on_demand=True)
	print(type(wb))
	print("Файл %s загружен"%(filepath))
	return wb


def compare_row(util_row, svod_row, svod):
	"""Производит сравнение строк и возвращает True если критерии
	удволетворены и False если нет"""

	result = False

	"""Сравниваем номер вагона"""
	if util_row[UTIL_DATA['VAGON_NUMBER']].value == \
		svod_row[SVOD_DATA['VAGON_NUMBER']].value:

		"""Номер вагона совпал, подготавливаем даты и сравниваем (СВОД-3 <= УТИЛИТА <= СВОД+3"""
		util_date = text_to_date(util_row[UTIL_DATA["IN_DATE"]].value)
		svod_date = xl_to_date(svod_row[SVOD_DATA["IN_DATE"]].value,svod)
		margin = datetime.timedelta(days = DATE_FRAME)

		if svod_date - margin <= util_date <= svod_date + margin:

			"""Дата утилиты входит в период свода (-3,+3)
				Проверяем есть ли номер детали в утилите
			"""


	return result

"""******* РАБОТА С ДАТАМИ **********************"""

def xl_to_date(cell_value,wb):
	"""Преобразуем дату из excel в дату python"""
	pydate = datetime.datetime(*xlrd.xldate_as_tuple(cell_value,wb.datemode))
	return pydate

def text_to_date(cell_value):
	"""Преобразуем текст даты из Утилиты в дату python
	   формат даты в ячейке 01/01/2014
	"""
	pydate = datetime.datetime.strptime(cell_value,"%d/%m/%Y")
	return pydate

"""**********************************************"""


def analyse_files():
	"""Проводит сравнительный анализ двух файлов и записывает результат
	в svod"""
	print("Загрузите сводную таблицу")
	svod = get_excel_file()
	print("Загрузите утилиту")
	util = get_excel_file()

	print("Зоздается таблица результатов")
#	svod_result = copy(svod)

#	print("Введите путь для сохранения результата")
#	result_path = input(">:")

	"""Получаем первые листа свода и утилиты"""
	util_sheet = util.sheet_by_index(0)
	svod_sheet = svod.sheet_by_index(0)
#	svod_result_sheet = svod_result.get_sheet(0)

	"""Сохраняем результат"""
#	svod_result.save("%sresult.xls"%(result_path))
#	print("Результат сохранен в %s"%(result_path))

	res = compare_row(util_sheet.row(UTIL_ROW_START), \
			svod_sheet.row(SVOD_ROW_START), svod)

	print(res)

main()