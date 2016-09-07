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
	'COMMENT': 25,
	"GRADATION": 13,
	'SCEP': 26,
	'UTIL_ROW':27
}

SVOD_ROW_START = 3

PART_TYPE = {
	'КП': 'Колесная пара',
	'БР': 'Боковая рама',
	'НБ': 'Надрессорная балка',

}

DATE_FRAME = 3

SAVE_RANGE = 500


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


def compare_row(util_row, svod_row, svod_row_num, svod, util_cache):
	"""Производит сравнение строк и возвращает True если критерии
	удволетворены и False если нет"""

	result = False

	if not svod_row[SVOD_DATA["COMMENT"]].value:
		"""(1) Проверяем, есть ли уже записанная деталь"""
#		print("-->(1) Записи нет")

		if util_row[UTIL_DATA['VAGON_NUMBER']].value == svod_row[SVOD_DATA['VAGON_NUMBER']].value:
			"""(2) Записи нет - проверяем номер вагона"""
#			print("-->(2) Номер вагона совпал")
			
			util_date = text_to_date(util_row[UTIL_DATA["IN_DATE"]].value)
			svod_date = xl_to_date(svod_row[SVOD_DATA["IN_DATE"]].value,svod)
			margin = datetime.timedelta(days = DATE_FRAME)

			if (svod_date - margin).date() < util_date.date() < (svod_date + margin).date():
				"""(3) Номер вагона совпал - проверяем дату"""
#				print("-->(3) Дата совпала")

				if util_row[UTIL_DATA["PART_NUMBER"]].value:
					"""(4) Дата совпала - проверяем наличие номера детали"""
#					print("-->(4) Номер есть")

					if int(util_row[UTIL_DATA["PART_NUMBER"]].value) == int(svod_row[SVOD_DATA["PART_NUMBER"]].value):
						"""(5) Номер детали есть - сравниваем"""
#						print("-->(5) Номер совпал")

						"""Номер детали совпал - записываем данные строку свода
						и наличие номер детали в util_cache"""
						util_cache.append({"svod_row_num": svod_row_num, "type": 1})
						result = True

				elif PART_TYPE[util_row[UTIL_DATA["TYPE"]].value] == svod_row[SVOD_DATA["TYPE"]].value:
					"""(6) У утилиты нет номера детали. Сравниваем тип детали"""
#					print("-->(6) Тип детали совпал")

					if util_row[UTIL_DATA["TYPE"]].value == "КП":
						"""(7) Тип детали совпа. Проверяем это колесная пара?"""
#						print("-->(6) Тип КП")
						util_gradation = get_util_gradation(util_row)

						if util_gradation[1] <= int(svod_row[SVOD_DATA["GRADATION"]].value) <= util_gradation[0]:
							"""(8) Тип детали является КП. Сравниваем Градацию"""
#							print("-->(8) Градация совпала")
							""" Градация совпала - записываем в КЭШ номер строки свода и градацию"""
							util_cache.append({"svod_row_num": svod_row_num, "type": 2})
							result = True
						else:
							"""Градация не совпала, записываем в КЭШ без типа"""
							util_cache.append({"svod_row_num": svod_row_num, "type": 0})
							result = True
					else:
						"""Тип детали не КП, но совпала - записываем в КЭШ без типа"""
						util_cache.append({"svod_row_num": svod_row_num, "type": 0})
						result = True

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
"""******* РАБОТА С КЭШЕМ УТИЛИТЫ****************"""
def check_util_cashe(util_cache):
	if util_cache:
		for item in util_cache:
			if item["type"] == 1:
				return item["svod_row_num"]
		for item in util_cache:
			if item["type"] == 2:
				return item["svod_row_num"]
		return util_cache[0]["svod_row_num"]
	else:
		return 0
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

	"""Подсчитываем количество строк в утилите и своде"""
	util_rows = util_sheet.nrows
	svod_rows = svod_sheet.nrows

	print("Строк в утилите: %s"%(util_rows - UTIL_ROW_START))
	print("Строк в своде:   %s"%(svod_rows - SVOD_ROW_START))

	"""Проводим проверку по всем строкам"""
	for util_row_num in range(UTIL_ROW_START, util_rows):
		print("Строка утилиты: %s"%(util_row_num+1), end=" --> ")
		util_cache = []
		comp_result = False

		for svod_row_num in range(SVOD_ROW_START, svod_rows):
			try:
				"""Проверка если запись совполает"""
				comp_row_result = compare_row(util_sheet.row(util_row_num), svod_sheet.row(svod_row_num), svod_row_num, svod, util_cache)
				if comp_row_result:
					comp_result = True
			except Exception:
				print("warnign: util-%s  svod-%s"%(util_row_num+1, svod_row_num+1))

		print(str(comp_result))

		if comp_result:
			"""Если схождение - записываем результат в итоговую таблицу"""
			write_row_number = check_util_cashe(util_cache)
#			print("-->row_write: "+str(write_row_number))

			"""Записываем наименование МЦ"""
			svod_result_sheet.write(write_row_number, SVOD_DATA["COMMENT"], util_sheet.cell(util_row_num,UTIL_DATA["NAME_MC"]).value)
			"""Записываем сцеп"""
			svod_result_sheet.write(write_row_number, SVOD_DATA["SCEP"], util_sheet.cell(util_row_num,UTIL_DATA["SCEP"]).value)
			"""Записываем соотвествующую строку утилиты"""
			svod_result_sheet.write(write_row_number, SVOD_DATA["UTIL_ROW"], str(util_row_num))

		if not util_row_num%SAVE_RANGE:
			"""Сохраняем каждые SAVE_RANGE строк"""
			svod_result.save("%s\\result.xls"%(result_path))
			print("Progress saved")

	"""Сохраняем результат"""
	svod_result.save("%s\\result.xls"%(result_path))
	print("Результат сохранен в %s"%(result_path))

	""" ТЕСТЫ при разработке """
#	res = compare_row(util_sheet.row(UTIL_ROW_START), \
#			svod_sheet.row(SVOD_ROW_START), svod)


main()