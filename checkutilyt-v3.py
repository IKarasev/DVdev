# coding: utf8
""" Testing-comaprison of utility and svod reports as excel files
By Karasev I.A.
"""

import xlrd
import xlwt
import datetime
import getopt
import sys
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
	'IN_VAGON_NUMBER': 2,
	'OUT_VAGON_NUMBER': 19,
	'IN_DATE': 4,
	'OUT_DATE': 13,
	'IN_PART_NUMBER': 7,
	'OUT_PART_NUMBER': 14,
	'TYPE': 6,
	'IN_RADATION': 10,
	'OUT_RADATION': 17,
	'IN_COMMENT': 20,
	'OUT_COMMENT': 21,
	'UTIL_ROW':22
}

SVOD_ROW_START = 1

PART_TYPE = {
	'КП': 'Колесная пара',
	'БР': 'Боковая рама',
	'НБ': 'Надрессорная балка',

}

DATE_FRAME = 3

SAVE_RANGE = 500

SAVE_NAME = "result.xlx"

UTIL_FILE = ""
SVOD_FILE = ""

MONTHS = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь']

def main(argv):
	"""Функция выполняемая при запуске скрипта"""

	"""Обработка аргументов"""
	try:
		opts, args = getopt.getopt(argv,"hd:s:c:n:",["DATE_FRAME=","SAVE_RANGE=","SVOD_COMMENT=","SAVE_NAME"])
	except getopt.GetoptError:
		print("ERROR: wrong arguments")
		sys.exit(2)

	global SAVE_RANGE
	global DATE_FRAME
	global SVOD_DATA
	global SAVE_NAME

	try:
		for opt, arg in opts:
			if opt == '-h':
				print('\n-d [--DATE_FRAME] <num>   - опрределяет рамки времени от даты свода в днях, в который должна попасть дата утилиты' + \
					'(по умолчанию - %s дней)\n'%(DATE_FRAME))
				print('-s [--SAVE_RANGE] <num>   - число, период строк утилыты, через который будет производится сохранение результатов ' + \
					'(по умолчанию - %s строк)\n'%(SAVE_RANGE))
				print('-c [--SVOD_COMMENT] <num> - номер столбца таблицы свода (по умолчанию - %s)\n'%(SVOD_DATA["IN_COMMENT"]))
				print('-n [--SAVE_NAME] <str>    - имя сохраняемого файла (по умолчанию - ' + SAVE_NAME + ')\n')
				sys.exit(2)
			elif opt in("-d","--DATE_FRAME"):
				DATE_FRAME = int(arg)
			elif opt in ("-s","--SAVE_RANGE"):
				SAVE_RANGE = int(arg)
			elif opt in ("-c","--SVOD_COMMENT"):
				SVOD_DATA["IN_COMMENT"] = int(arg) - 1
				SVOD_DATA["OUT_COMMENT"] = int(arg)
				SVOD_DATA["UTIL_ROW"] = int(arg) + 1
			elif opt in ("-n","--SAVE_NAME"):
				SAVE_NAME = str(arg)
	except Exception:
		print("ERROR: arguments couldn't be parsed")
		sys.exit(2)

	print('--> SAVE RANGE:          %s rows'%(SAVE_RANGE))
	print('--> DATE FRAME:          %s days'%(DATE_FRAME))
	print('--> SVOD COMMENT COLUMN: %s'%(SVOD_DATA["IN_COMMENT"]))

	"""Начало обработки"""
	print("Сравнение открыто")
	analyse_files()


"""Function to load an excel file"""
def get_excel_file():
	"""Запрашивает путь к файлу excel и возвращает книгу в виде xlrd"""
	try:
		filepath = input("Введите путь к файлу: ")
		wb = xlrd.open_workbook(filepath, formatting_info=False, on_demand=True)
		print(type(wb))
		print("Файл %s загружен"%(filepath))
		return wb
	except Exception:
		print("ОШИБКА: Не возможно загрузить файл")
		sys.exit(2)

"""Function to get gradation of the part from util"""
def get_util_gradation(util_row):
	grad_str = str(util_row[UTIL_DATA["GRADATION"]].value)
	if grad_str:
		grad = grad_str.split("-")
		grad[0] = int(grad[0]+"0")
		grad[1] = int(grad[1]+"0")
	else:
		grad = [-5,-5]
	return grad

"""Function to get int from string"""
def int_from_str(str):
	if not str:
		return int(str)
	else:
		return -1

"""******* Процедура сравнения строк утилиты и 140 (свода) ******************************"""
def compare_row(util_row, svod_row, svod_row_num, svod, util_cache, comments_cache, mode):
	result = False
	"""Проверяем режим сравнения - IN-приход, OUT-расход"""
	if mode = "IN":
		cache_mode = 0
	elif mode == "OUT":
		cache_mode = 1
	else:
		print("Нет такого [mode] для проверки")
		sys.exit(2)

	if not svod_row[SVOD_DATA[mode+"_COMMENT"]].value and not comments_cache[cache_mode][svod_row_num]:
		"""(1) Проверяем, есть ли уже записанная деталь"""
#		print("-->(1) Записи нет")

		if util_row[UTIL_DATA['VAGON_NUMBER']].value == svod_row[SVOD_DATA['IN_VAGON_NUMBER']].value:
			"""(2) Записи нет - проверяем номер вагона"""
#			print("-->(2) Номер вагона совпал")

			util_date = text_to_date(util_row[UTIL_DATA["IN_DATE"]].value)
			svod_date = xl_to_date(svod_row[SVOD_DATA[mode+"_DATE"]].value,svod)
			margin = datetime.timedelta(days = DATE_FRAME)

			if (svod_date - margin).date() < util_date.date() < (svod_date + margin).date():
				"""(3) Номер вагона совпал - проверяем дату"""
#				print("-->(3) Дата совпала")
				
				if str(util_row[UTIL_DATA["PART_NUMBER"]].value):
					"""(4) Дата совпала - проверяем наличие номера детали"""
#					print("-->(4) Номер есть")

					if str(util_row[UTIL_DATA["PART_NUMBER"]].value) == str(svod_row[SVOD_DATA[mode+"_PART_NUMBER"]].value):
						"""(5) Номер детали есть - сравниваем"""
#						print("-->(5) Номер совпал")

						"""Номер детали совпал - записываем данные строку свода
						и наличие номер детали в util_cache. Сразу возвращаем данные на запись"""
						util_cache.append({"svod_row_num": svod_row_num, "type": 1})
						return True

				elif PART_TYPE[util_row[UTIL_DATA["TYPE"]].value] == svod_row[SVOD_DATA["TYPE"]].value:
					"""(6) У утилиты нет номера детали. Сравниваем тип детали"""
#					print("-->(6) Тип детали совпал")
					
					if util_row[UTIL_DATA["TYPE"]].value == "КП":
						"""(7) Тип детали совпа. Проверяем это колесная пара?"""
#						print("-->(7) Тип КП")
						util_gradation = get_util_gradation(util_row)

						if util_gradation[1] <= int_from_str(svod_row[SVOD_DATA[mode+"_GRADATION"]].value) <= util_gradation[0]:
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
	
"""**********************************************"""
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

	print("Записей в утилите: %s"%(util_rows - UTIL_ROW_START))
	print("Записей в своде:   %s"%(svod_rows - SVOD_ROW_START))

	"""В данном кэше сохраняется список совпадений детали для текущей итерации в утилите"""
	in_util_cache = []
	out_util_cash = []

	"""В данном кэше храниться история совпадений (наличие комментариев) в своде: 0-нет, 1-есть, [0] - приход, [1] - расход"""
	comments_cache = [[0]*svod_rows, [0]*svod_rows]

	"""Проводим проверку по всем строкам"""
	for util_row_num in range(UTIL_ROW_START, util_rows):
		print("Строка утилиты: %s"%(util_row_num+1), end=" --> ")
		#Очистка кэша совпадений предыдущей итерации

		del in_util_cache[:]
		del out_util_cash[:]

		comp_result = False

		for svod_row_num in range(SVOD_ROW_START, svod_rows):
			try:
				"""Проверка если запись совполает"""
				comp_row_result = compare_row(util_sheet.row(util_row_num), svod_sheet.row(svod_row_num), \
					svod_row_num, svod, in_util_cache, comments_cache, "IN")
				
				if comp_row_result:
					comp_result = True
			except Exception:
				print("warnign: util-%s  svod-%s"%(util_row_num+1, svod_row_num+1))
#		print("Util cache: "+str(util_cache))
		print(str(comp_result))

		if comp_result:
			"""Если схождение - записываем результат в итоговую таблицу"""
			write_row_number = check_util_cashe(util_cache)
			#Ставим пометку в кэш совпадений
			comments_cache[write_row_number] = 1
#			print("-->row_write: "+str(write_row_number))

			"""Записываем наименование МЦ"""
			svod_result_sheet.write(write_row_number, SVOD_DATA["COMMENT"], util_sheet.cell(util_row_num,UTIL_DATA["NAME_MC"]).value)
			"""Записываем сцеп"""
			svod_result_sheet.write(write_row_number, SVOD_DATA["SCEP"], util_sheet.cell(util_row_num,UTIL_DATA["SCEP"]).value)
			"""Записываем соотвествующую строку утилиты"""
			svod_result_sheet.write(write_row_number, SVOD_DATA["UTIL_ROW"], str(util_row_num+1))

		if not util_row_num%SAVE_RANGE:
			"""Сохраняем каждые SAVE_RANGE строк"""
			svod_result.save("%s\\result.xls"%(result_path))
			print("Progress saved")

	"""Сохраняем результат"""
	svod_result.save("%s\\%s"%(result_path, SAVE_NAME))
	print("Результат сохранен в %s\\%s"%(result_path, SAVE_NAME))
#	print("Карта совподения"+str(comments_cache))

if __name__ == "__main__":
	main(sys.argv[1:])