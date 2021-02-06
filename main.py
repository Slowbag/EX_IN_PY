import os
import doge
import pandas as pd
import xlrd, xlwt
import subprocess
import time

startTime = time.time()

rb = xlrd.open_workbook('../EX_IN_PY/ex_in_py.xls',formatting_info=True)

sheet = rb.sheet_by_index(0)
#b4:9
b3 = sheet.row_values(2)[1]
b4 = sheet.row_values(3)[1]
b5 = sheet.row_values(4)[1]
b6 = sheet.row_values(5)[1]
b7 = sheet.row_values(6)[1]
b8 = sheet.row_values(7)[1]
b9 = sheet.row_values(8)[1]

#c4:9
c3 = sheet.row_values(2)[2]
c4 = sheet.row_values(3)[2]
c5 = sheet.row_values(4)[2]
c6 = sheet.row_values(5)[2]
c7 = sheet.row_values(6)[2]
c8 = sheet.row_values(7)[2]
c9 = sheet.row_values(8)[2]

#D4:9
d3 = sheet.row_values(2)[3]
d4 = sheet.row_values(3)[3]
d5 = sheet.row_values(4)[3]
d6 = sheet.row_values(5)[3]
d7 = sheet.row_values(6)[3]
d8 = sheet.row_values(7)[3]
d9 = sheet.row_values(8)[3]

#E4:9
e3 = sheet.row_values(2)[4]
e4 = sheet.row_values(3)[4]
e5 = sheet.row_values(4)[4]
e6 = sheet.row_values(5)[4]
e7 = sheet.row_values(6)[4]
e8 = sheet.row_values(7)[4]
e9 = sheet.row_values(8)[4]

#1-3 уровень продукций
vp0=b4*b5
vp0_5=c4*b5
vp1=c4*c5

lv2=vp0_5-vp0
lv3=vp1-vp0_5
#Итог
itog=lv2+lv3

#4 уровень
lv4_0=b4*b6*b8*b9 // 1000
lv4_1=c4*b6*b8*b9 // 1000
lv4_2=c4*c6*b8*b9 // 1000
lv4_3=c4*c6*c8*b9 // 1000
lv4_4=c4*c6*c8*c9 // 1000
lv4=((c4*c6*c8*c9)-(b4*b6*b8*b9)) // 1000

#Численость рабочих
vphr=lv4_1-lv4_0 // 1000
#Кол-дней отработаных
vpd=lv4_2-lv4_1 // 1000
#Сред. прод.рабочих дней
vpp=lv4_3-lv4_2 // 1000
#Сред. час. выроботка
vpch=lv4_4-lv4_3 // 1000
#Итог
sumir=(lv4_1-lv4_0)+(lv4_2-lv4_1)+(lv4_3-lv4_2)+(lv4_4-lv4_3)

def view_table():
	print("Показать таблицу")
	df = pd.DataFrame({'Yslovnoe oboznach': ['VP', 'CHR', 'GV', 'D',
							'DV', 'P', 'CHV'],
					'Yroven_pokazatel\n bazoviy' : [b3, b4, b5,
							b6, b7, b8, b9],
					'Yroven_pokazatel\n tekyshiy' : [c3, c4, c5,
							c6, c7, c8, c9],
					'Izmineniya\n absolut' : [d3,d4, d5,
							d6, d7, d8, d9],
					'Izmineniya\n otnositelnoe' : [e3, e4, e5,
					e6, e7, e8, e9]})
	print(df)

def view_lv1_3():
	print("1-3 уровень продукций")
	#1-3 уровень продукций
	vp0=b4*b5
	vp0_5=c4*b5
	vp1=c4*c5

	lv2=vp0_5-vp0
	lv3=vp1-vp0_5
	#Итог
	itog=lv2+lv3
	print(f'{vp0},{vp0_5},{vp1}')
	print(f'Рост числености рабочих = {lv2}\nПовышение уровня производительности труда = {lv3}')

def view_lv4():
	print("4 уровень продукций")
	#4 уровень
	lv4_0=b4*b6*b8*b9 // 1000
	lv4_1=c4*b6*b8*b9 // 1000
	lv4_2=c4*c6*b8*b9 // 1000
	lv4_3=c4*c6*c8*b9 // 1000
	lv4_4=c4*c6*c8*c9 // 1000
	lv4=((c4*c6*c8*c9)-(b4*b6*b8*b9)) // 1000
	print(f'{lv4_0},{lv4_1},{lv4_2},{lv4_3},{lv4_4}\nОбъем продукций вырос {lv4}')

def otchet():
	print("Остальное")
	#Численость рабочих
	vphr=lv4_1-lv4_0 // 1000
	#Кол-дней отработаных
	vpd=lv4_2-lv4_1 // 1000
	#Сред. прод.рабочих дней
	vpp=lv4_3-lv4_2 // 1000
	#Сред. час. выроботка
	vpch=lv4_4-lv4_3 // 1000
	#Итог
	sumir=(lv4_1-lv4_0)+(lv4_2-lv4_1)+(lv4_3-lv4_2)+(lv4_4-lv4_3)
	print(f'Численость рабочих {vphr}\nКол-дней отработаных {vpd}')
	print(f'Сред. прод.рабочих дней {vpp}\nСред. час. выроботка {vpch}')
	print(f'Итог {sumir}')

def table_otchet():
	print("Таблица отчета")
	df = pd.DataFrame({'Yslovnoe oboznach': ['VP', 'CHR', 'GV', 'D',
							'DV', 'P', 'CHV'],
					'Yroven_pokazatel\n bazoviy' : [b3, b4, b5,
							b6, b7, b8, b9],
					'Yroven_pokazatel\n tekyshiy' : [c3, c4, c5,
							c6, c7, c8, c9],
					'Izmineniya\n absolut' : [d3,d4, d5,
							d6, d7, d8, d9],
					'Izmineniya\n otnositelnoe' : [e3, e4, e5,
					e6, e7, e8, e9],
					'lv1-3' : [vp0, vp0_5, vp1, lv2, lv3, itog, ' '],
					'lv4' : [lv4_0, lv4_1, lv4_2, lv4_3, lv4_4, lv4, ' '],
					'CHR_VPD_VPP_VPCH' : [vphr, vpd, vpp, vpch, sumir, ' ', ' ']})
	print(df)

def table_save():
	print("Таблица сохранина")
	df = pd.DataFrame({'Yslovnoe oboznach': ['VP', 'CHR', 'GV', 'D',
							'DV', 'P', 'CHV'],
					'Yroven_pokazatel\n bazoviy' : [b3, b4, b5,
							b6, b7, b8, b9],
					'Yroven_pokazatel\n tekyshiy' : [c3, c4, c5,
							c6, c7, c8, c9],
					'Izmineniya\n absolut' : [d3,d4, d5,
							d6, d7, d8, d9],
					'Izmineniya\n otnositelnoe' : [e3, e4, e5,
					e6, e7, e8, e9],
					'lv1-3' : [vp0, vp0_5, vp1, lv2, lv3, itog, ' '],
					'lv4' : [lv4_0, lv4_1, lv4_2, lv4_3, lv4_4, lv4, ' '],
					'CHR_VPD_VPP_VPCH' : [vphr, vpd, vpp, vpch, sumir, ' ', ' ']})
	df.to_excel('../EX_IN_PY/ex_in_py_otchet.xlsx')
	subprocess.Popen(('start', '../EX_IN_PY/ex_in_py_otchet.xlsx'), shell = True)

def ex():
	print('Выход')
	exit()

while True:

	os.system("cls")
	doge.doge()
	print(f'Меню\n1.Показать таблицу\n2.Показать 1-3 уровень продукций')
	print(f'3.Показать 4 уровень продукций\n4.Остальное\n5.Показать таблице отчета')
	print(f'6.Сохранить отчет и открыть\n7.Выход')
	endTime = time.time() #время конца замера
	totalTime = endTime - startTime #вычисляем затраченное время
	print("Время, затраченное на выполнение данного кода = ", totalTime)
	x  = int(input('(→_→)'))
	if x == 0: break
	elif x == 1: view_table()
	elif x == 2: view_lv1_3()
	elif x == 3: view_lv4()
	elif x == 4: otchet()
	elif x == 5: table_otchet()
	elif x == 6: table_save()
	elif x == 7: ex()
	else: print('Некорректный ввод')

	y = input('Вы хотите продолжить Y или N?')
	if y == "0" or y == 'no' or y == 'N' or y == 'n' or y == 'нет': exit()
	if y == "1" or y == 'yes' or y == 'Y' or y == 'y' or y == 'да': continue
