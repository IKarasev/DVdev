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
	'PART_NUMBER': 15,
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
	'COMMENT': 26,
	"GRADATION": 13,
	'SCEP': 27,
	'UTIL_ROW':28
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


def get_util_gradation(util_row):
	grad = util_row[UTIL_DATA["GRADATION"]].value.split("-")
	grad[0] = int(grad[0])
	grad[1] = int(grad[1])
	return grad


def compare_row(util_row, svod_row, svod):
	"""Производит сравнение строк и возвращает True если критерии
	удволетворены и False если нет"""

	result = False

#	print("--->(0) Сравнение началось")

	"""(1) Сравниваем номер вагона"""
	if util_row[UTIL_DATA['VAGON_NUMBER']].value == svod_row[SVOD_DATA['VAGON_NUMBER']].value:

#		print("--->(1) Номер вагона совпал")

		"""(2) Номер вагона совпал, подготавливаем даты и сравниваем (СВОД-3 <= УТИЛИТА <= СВОД+3"""
		util_date = text_to_date(util_row[UTIL_DATA["IN_DATE"]].value)
		svod_date = xl_to_date(svod_row[SVOD_DATA["IN_DATE"]].value,svod)
		margin = datetime.timedelta(days = DATE_FRAME)

		if (svod_date - margin).date() < util_date.date() < (svod_date + margin).date():
			"""(2) Дата утилиты входит в период свода (-3,+3)
				Проверяем есть ли номер детали в утилите
			"""

#			print("--->(2) Даты совпали")

			if util_row[UTIL_DATA["PART_NUMBER"]].value:
				"""(3) У утилиты есть номер детали"""

#				print("--->(3) Номер детали есть")

				"""(4) Сравниваем номера деталей в утилите и своде"""
				if int(util_row[UTIL_DATA["PART_NUMBER"]].value) == int(svod_row[SVOD_DATA["PART_NUMBER"]].value):
					"""Номер детали совпал - записи совпали"""
					
#					print("--->(4) Номер детали совпадает")
					
					result = True


			elif PART_TYPE[util_row[UTIL_DATA["TYPE"]].value] == svod_row[SVOD_DATA["TYPE"]].value:
				"""(5) У утилиты нет номера детали. Сравниваем тип детали"""
#				print("--->(5) Тип детали совпал")

				if util_row[UTIL_DATA["TYPE"]].value == "КП":
					"""(6) Тип детали совподает. Проверяем, является ли деталь КП"""
#					print("--->(6) Деталь является КП")

					util_gradation = get_util_gradation(util_row)

					"""(7)Деталь является КП, сравниваем градацию"""
					if util_gradation[1] <= int(svod_row[SVOD_DATA["GRADATION"]].value) <= util_gradation[0]:
						"""Если градация свода входит в градацию утилиты, то записываем результат"""
						result = True
#						print("--->(7) Градация совподает")

					"""(8) Деталь не является КП, но тип детали сходится, проверяем, было ли схождение раньше"""

				elif not svod_row[SVOD_DATA["COMMENT"]].value:
					"""Ранних схождений не найденно - записываем текущее схождение"""
					result = True

#					print("--->(7) НЕ КП")
#					print("--->(8) Записи нет")


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

	print("Cоздается таблица результатов")
	svod_result = copy(svod)

	print("Введите путь для сохранения результата")
	result_path = input(">:")

	"""Получаем первые листы свода и утилиты"""
	util_sheet = util.sheet_by_index(0)
	svod_sheet = svod.sheet_by_index(0)
	svod_result_sheet = svod_result.get_sheet(0)

	"""Проводим проверку по всем строкам"""
	for util_row_num in range(UTIL_ROW_START, util_sheet.nrows):
		print("Строка утилиты: %s"%(util_row_num+1))

		for svod_row_num in range(SVOD_ROW_START, svod_sheet.nrows):
			"""Проверка если запись совполает"""
			comp_result = compare_row(util_sheet.row(util_row_num), svod_sheet.row(svod_row_num), svod)

			"""Если схождение - записываем результат в итоговую таблицу"""
			if comp_result:
				"""Записываем наименование МЦ"""
				svod_result_sheet.write(svod_row_num, SVOD_DATA["COMMENT"], util_sheet.cell(util_row_num,UTIL_DATA["NAME_MC"]).value)
				"""Записываем сцеп"""
				svod_result_sheet.write(svod_row_num, SVOD_DATA["SCEP"], util_sheet.cell(util_row_num,UTIL_DATA["SCEP"]).value)
				"""Записываем соотвествующую строку утилиты"""
				svod_result_sheet.write(svod_row_num, SVOD_DATA["UTIL_ROW"], str(util_row_num))

	"""Сохраняем результат"""
	svod_result.save("%s\\result.xls"%(result_path))
	print("Результат сохранен в %s"%(result_path))

	""" ТЕСТЫ при разработке """
#	res = compare_row(util_sheet.row(UTIL_ROW_START), \
#			svod_sheet.row(SVOD_ROW_START), svod)

main()