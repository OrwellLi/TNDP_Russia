"""
Created on Thu 22 MAY 16:22:22 2025

@author OrwellLi

Project based on solving Transit Network Design Problem (TNDP) and then will be optimized

"""
#---------------------------------
#Импорт библиотек и настроек

from openpyxl import load_workbook
import numpy as np
import copy as copy
import matplotlib.pyplot as plt
import pandas as pd
import time 
starttime = time.time
import openrouteservice
from math import radians, cos, sin, asin, sqrt
import datetime
import math
from random import seed
from random import randint
seed(1)
from tqdm import tqdm
import random

#--------------------------------
#Загрузка параметров

average_car_speed = 180. #[km/h]
max_catchment_area_perimeter = 5. #[часы]
maximum_potential_access_egress_dist = average_car_speed * max_catchment_area_perimeter
max_accsess_egress_time = 0.5

VoTime_access    = 67.5 #ценность времени до прибытия в ап(из дома в ап)
VoTime_waiting   = 75   #ценность времени ожидания
VoTime_invehicle = 50   #нормальное время в дороге
VoTime_transfer  = 50   #нормальное время во время трансфера
VoTime_egress    = 67.5 #ценность времени после выхода из ап(из ап до дома)

weight_acc = VoTime_access / VoTime_invehicle    #распределение весов и перевод их в нормальные единицы, так как ценность времени измеряется в Руб/ч
weight_wait = VoTime_waiting / VoTime_invehicle
weight_inv = 1.0
weight_tran = VoTime_transfer / VoTime_invehicle
weight_egr = VoTime_egress / VoTime_invehicle

time_access = 0.60 #time from leaving home to seat in plane
time_dec = 0.60    #time from leaving plane to reach home
fac_dt = 1         #множитель времени, для поправки момента искревления пути
distance_acc = 50  #расстояние доступа
distance_dec = 50  #расстояние выхода
speed_kreys = 850  #крейсерская скорость

#матрица времени для разных режимов[hours]
time_acc_m = [0.5, 0.25, 0.0]
time_wait_m = [1.8333, 0.5, 0.0]
time_transfer_m = [1.5, 1.0, 0.0]
time_egres_m = [0.5, 0.25, 0.0]

time_wait = time_wait_m[0]

lowerboundary_distance = 200. #граничное условие для дистанции(мин дистанция для учета маршрута)
lowerboundary_duration = 2.   #граничное условие для времени(минимальная длительность)


fc_detour = [1., 1.09, 1.20] #plane, HSR, car коэффициент удлинения пути для различного транспорта
vehicle_speed = [850., 220., 90.] #plane, HSR, car

daily_operational_hours = 18.0 #часы работы сети в день

#веса в функции полезности(определить xxx)
#a1 = -1 * xxx #вес общего времени поездки (отрицательный – чем больше время, тем хуже). 
#a2 = +1 * xxx  #вес частоты отправлений (положительный – чем чаще рейсы, тем лучше).
#a3 = -1 * xxx #вес пересечения границы (отрицательный – усложняет поездку).

'''--------------------------------------------------------
-----------------------   ШAГ1    -------------------------
---------    Загрузка данных o пассажиропотоках -----------'''

print('')
print('log. ШAГ 1  ------  Загрузка данных o пассажиропотоках')

# location = ([r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_vvo.xlsx', 
#              r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_vko.xlsx',
#              r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_svo.xlsx',
#              r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_ovb.xlsx',
#              r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_led.xlsx',
#              r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_kzn.xlsx',
#              r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_ikt.xlsx',
#              r'/Users/Dimka/Desktop/HSA network design/TNDP_Russia/avia_par_dme.xlsx]'])



location = ([r'TNDP_Russia/avia_par_vvo.xlsx', 
             r'TNDP_Russia/avia_par_vko.xlsx',
             r'TNDP_Russia/avia_par_svo.xlsx',
             r'TNDP_Russia/avia_par_ovb.xlsx',
             r'TNDP_Russia/avia_par_led.xlsx',
             r'TNDP_Russia/avia_par_kzn.xlsx',
             r'TNDP_Russia/avia_par_ikt.xlsx',
             r'TNDP_Russia/avia_par_dme.xlsx'])

sheet = (['avia_par_vvo',
          'avia_par_vko',
          'avia_par_svo', 
          'avia_par_ovb', 
          'avia_par_led', 
          'avia_par_kzn', 
          'avia_par_ikt', 
          'avia_par_dme'])

Airport_data = [['Unit', 'Msr', 'orig_ap', 'dest_ap', 'pax']]
Airport_data_frequency = [['Unit', 'Msr', 'orig_ap', 'dest_ap', 'flights']]

print('')
print('log. Bытаскиваем данные от АП по паксам и полетам')

with tqdm(total=len(location), desc="Processing", bar_format="{l_bar}{bar} [ time left: {remaining} ]", position=0, leave=True) as pbar:  
    for u in range(len(location)):
        wb = load_workbook(location[u])
        ws = wb[sheet[u]]
        
        #загружаем файлы для обработки
        years = np.array([[i.value for i in j] for j in ws['A1':'LL1']]) 
        XX = np.zeros(len(years[0]), dtype=object)
        for i in range(len(XX)):
            XX[i] = str(copy.copy(years[0,i]))
        years = XX #years
        
        #ищем колонку sum
        column = np.where(years=='sum')[0][0]
        alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        
        #строим таблицу
        upperleft = str('A') + str('1')
        lowerright = str(alphabet[column]) + str(ws.max_row)
        Table = np.array([[i.value for i in j] for j in ws[upperleft:lowerright]])

        #убираем ненужные колонки + проверка
        correction = 0
        for column in range(len(Table[0])):
            if Table[0][column - correction] != 'xx':
                if Table[0][column - correction] != 'sum':
                    Table = np.delete(Table, column - correction, axis=1)
                    correction = correction + 1

        #собираем нужную информацию из табллиц
        for row in range(len(Table)):
            if Table[row][0] == 'PAS':
                if Table[row][1] == 'PAS_CRD':
                    Airport_data = np.append(Airport_data, [Table[row]], axis=0)
            if Table[row][0] == 'FLIGHT':
                if Table[row][1] == 'CAF_PAS':
                    Airport_data_frequency = np.append(Airport_data_frequency, [Table[row]], axis=0)
                    
        pbar.update(1)
        
#save pax
print('')
print('log. Сохранение таблицы с паксами')
df = pd.DataFrame(Airport_data)
df.to_excel(excel_writer = "Airport_data.xlsx")

#save freq
print('')
print('log. Сохранение таблицы с полетами')
df = pd.DataFrame(Airport_data_frequency)
df.to_excel(excel_writer = "TNDP_Russia/Airport_data_frequency.xlsx")

print('')
print('Well done')


'''--------------------------------------------------------
-----------------------   ШAГ 2    -------------------------
------------    Очистка и модификация данных -----------'''

print('')
print('log. ШAГ 2  ------  Очистка и модификация данных')
print('')
print('log. Загрузка данных по паксам')
wb_Airport_data = load_workbook(r'Airport_data.xlsx')
ws_Airport_data = wb_Airport_data['Sheet1']
Airport_data = np.array([[i.value for i in j] for j in ws_Airport_data['B2':'H14174']])  


print('')
print('log. Загрузка данных по полетам')
wb_Airport_data_frequency = load_workbook(r'TNDP_Russia/Airport_data_frequency.xlsx')
ws_Airport_data_frequency = wb_Airport_data_frequency['Sheet1']

data_rows = []
for row in ws_Airport_data_frequency.iter_rows(min_row=2, values_only=True):
    # Пропускаем пустые строки
    if any(cell is not None for cell in row):
        data_rows.append(list(row))

Airport_data_frequency = np.array(data_rows, dtype=object)


print('')
print('log. Подгружаем данные об аэропортах и убираем лишние ячейки')

#Проверка
correction = 0
for row in range(len(Airport_data)):
    if Airport_data[row-correction][6] == ': ':
        Airport_data = np.delete(Airport_data, [row-correction], axis=0)
        correction = correction+1
        
correction = 0
for row in range(len(Airport_data)):
    if Airport_data[row-correction][3] == 'ZZZZ' or Airport_data[row-correction][5] == 'ZZZZ':
        Airport_data = np.delete(Airport_data, [row-correction], axis=0)
        correction = correction+1
        

print('')
print( 'log. Идентификация оставшихся аэропортов')
#Создаем список уникальных ап
set_of_airports = set() 

for row in range(len(Airport_data)-1):
    # Добавляем только не-None значения
    if Airport_data[row+1][2] is not None and str(Airport_data[row+1][2]).strip() != '':
        set_of_airports.add(str(Airport_data[row+1][2]).strip())
    if Airport_data[row+1][3] is not None and str(Airport_data[row+1][3]).strip() != '':
        set_of_airports.add(str(Airport_data[row+1][3]).strip())

set_of_airports.discard('None')
set_of_airports.discard('')
# print(set_of_airports)
    
Num_airports  =  len(set_of_airports)

list_Airport_temp = sorted(list(set_of_airports))

list_of_airports = np.zeros(Num_airports, dtype=object)

for n in range(Num_airports):
    list_of_airports[n] = list_Airport_temp[n]


print('')
print('log. Загружаем характеристики аэропортов.')
print('     по категории  : аэропорт по коду ИАТА') 
print('     топографически: страна')
print('     географически : широта, долгота')

wb_Aiport_info = load_workbook(r'TNDP_Russia/airportinformation3.xlsx')
ws_Airport_info = wb_Aiport_info['airportdata']

name = np.array([[i.value for i in j] for j in ws_Airport_info['B2':'GPG2']]) 
XX = np.zeros(len(name[0]), dtype=object)
for i in range(len(XX)):
    XX[i] = str(copy.copy(name[0,i]))
name = XX #Название
# print("")
# print(name)

country = np.array([[i.value for i in j] for j in ws_Airport_info['B4':'GPG4']]) 
XX = np.zeros(len(country[0]), dtype=object)
for i in range(len(XX)):
    XX[i] = copy.copy(country[0,i])
country = XX #Страна
# print("")
# print(country)

IATA = np.array([[i.value for i in j] for j in ws_Airport_info['B7':'GPG7']]) 
XX = np.zeros(len(IATA[0]), dtype=object)
for i in range(len(XX)):
    XX[i] = str(copy.copy(IATA[0,i]))
IATA = XX #ИАТА-код
# print("")
# print(IATA)

Airport_longitude = np.array([[i.value for i in j] for j in ws_Airport_info['B8':'GPG8']]) 
XX = np.zeros(len(Airport_longitude[0]))
for i in range(len(XX)):
    XX[i] = str(copy.copy(Airport_longitude[0,i]))
Airport_longitude = XX #широта
# print("")
# print(Airport_longitude)

Airport_latitude = np.array([[i.value for i in j] for j in ws_Airport_info['B9':'GPG9']]) 
XX = np.zeros(len(Airport_latitude[0]))
for i in range(len(XX)):
    XX[i] = str(copy.copy(Airport_latitude[0,i]))
Airport_latitude = XX #долгота
# print("")
# print(Airport_latitude)

airport_information = np.full((Num_airports, 5),'n/a', dtype=object) #[IATA, name, coun, AP_lat, AP_lon]

#заполняем таблицу
for row in range(len(list_of_airports)):
    airport_information[row][0] = list_of_airports[row]
    for column in range(len(name)):
        if IATA[column] == airport_information[row][0]:
            # airport_information[row][1] = IATA[column]
            airport_information[row][1] = name[column]
            airport_information[row][2] = country[column]
            airport_information[row][3] = Airport_latitude[column]
            airport_information[row][4] = Airport_longitude[column]
            

#обновляем последовательность ап в коде IATA
for n in range(Num_airports):
    list_of_airports[n] = airport_information[n][0]

def Airport_index_num(a): #находим индекс аэропорта
    try:
        # Нормализуем входные данные (удаляем пробелы, приводим к верхнему регистру)
        search_code = str(a).strip().upper()
        
        # Ищем совпадение
        for idx, airport_code in enumerate(list_of_airports):
            if str(airport_code).strip().upper() == search_code:
                return idx
                
        # Если не найдено, выводим отладочную информацию
        return -1
        
    except Exception as e:
        print(f"Ошибка при поиске аэропорта '{a}': {str(e)}")
        return -1


print('')
print ('log. Комбинируем характеристики аэропорта в одну матрицу')
# print ('')

#строим Origin-Destinaion матрицу
OD_matrix = np.zeros((len(list_of_airports), len(list_of_airports)),dtype=object)
OD_matrix_frequency = np.zeros((len(list_of_airports), len(list_of_airports)),dtype=object)

OD_matrix = np.zeros((len(list_of_airports), len(list_of_airports)), dtype=object)

for row in range(len(Airport_data)-1):
    orig = Airport_data[row+1][2]
    dest = Airport_data[row+1][3]
    # Проверяем, что orig и dest не None и не строка "None"
    if (orig is not None and dest is not None and 
        orig != "None" and dest != "None"):
        j = Airport_index_num(orig)
        i = Airport_index_num(dest)
        # Убедимся, что индексы валидны (не -1)
        if i != -1 and j != -1:
            pax = Airport_data[row+1][4]
            # Убедимся, что pax не None
            OD_matrix[i,j] = pax if pax is not None else 0


print('')
print('log. Матрица сделана')        
        
OD_matrix_float = np.zeros_like(OD_matrix, dtype=float)
for i in range(OD_matrix.shape[0]):
    for j in range(OD_matrix.shape[1]):
        try:
            OD_matrix_float[i,j] = float(OD_matrix[i,j])
        except (ValueError, TypeError):
            OD_matrix_float[i,j] = 0.0

#Заполняем OD матрицу частотой полетов
for row in range(len(Airport_data_frequency)-1):
    j = Airport_index_num(Airport_data_frequency[row+1][3])  # origin airport
    i = Airport_index_num(Airport_data_frequency[row+1][4])  # destination airport
    freq_y = Airport_data_frequency[row+1][5]  # frequency value

    if freq_y == ': ':
        freq_y = 0
    OD_matrix_frequency[i,j] = float(freq_y) / 365.  # convert yearly to daily
    if freq_y == 0 and OD_matrix[i,j] > 0:
        OD_matrix_frequency[i,j] = -1 * float('inf')  # маркер для отсутствующих данных
        
mirror_matrix = np.zeros((len(list_of_airports), len(list_of_airports)), dtype=float)
for i in range(len(mirror_matrix)):
    for j in range(len(mirror_matrix)):
        mirror_matrix[i,j] = np.maximum(OD_matrix_float[i,j], OD_matrix_float[j,i]) / 2.0
        mirror_matrix[j,i] = mirror_matrix[i,j]  # Не нужно copy.copy для numpy массивов           

mirror_matrix_freq = np.zeros((len(list_of_airports), len(list_of_airports)), dtype=object)
for i in range(len(mirror_matrix_freq)):
    for j in range(len(mirror_matrix_freq)):
        val_i = OD_matrix_frequency[i,j]
        val_j = OD_matrix_frequency[j,i]
        if val_i in [': ', 'n/a'] or val_j in [': ', 'n/a']:
            mirror_matrix_freq[i,j] = 'n/a'
        elif val_i == ':' or val_j == ':':
            mirror_matrix_freq[i,j] = ':'
        else:
            mirror_matrix_freq[i,j] = max(float(val_i), float(val_j)) / 2.0
        mirror_matrix_freq[j,i] = mirror_matrix_freq[i,j]
        
        
# print(OD_matrix_frequency)
        
OD_matrix_numeric = np.zeros_like(OD_matrix, dtype=float)

for i in range(OD_matrix.shape[0]):
    for j in range(OD_matrix.shape[1]):
        # Пропускаем специальные значения
        if OD_matrix[i,j] in [':', 'n/a', 'None', None]:
            OD_matrix_numeric[i,j] = 0.0
        else:
            try:
                OD_matrix_numeric[i,j] = float(OD_matrix[i,j])
            except (ValueError, TypeError):
                OD_matrix_numeric[i,j] = 0.0
                print(f"Нечисловое значение в [{i},{j}]: {OD_matrix[i,j]}")

            

sum_vertical = np.zeros((len(list_of_airports), 1))
for i in range(len(list_of_airports)):
    row_sum = 0.0
    for val in OD_matrix[i, :]:
        try:
            row_sum += float(val)
        except (ValueError, TypeError):
            pass  # Пропускаем нечисловые значения
    sum_vertical[i] = row_sum
    
   
sum_horizontal = np.zeros((len(list_of_airports), 1))
for j in range(len(list_of_airports)):
    row_sum = 0.0
    for val in OD_matrix[:, j]:
        try:
            row_sum += float(val)
        except (ValueError, TypeError):
            pass  # Пропускаем нечисловые значения
    sum_horizontal[j] = row_sum   
   

OD_matrix = OD_matrix.astype(object)
mirror_matrix = mirror_matrix.astype(object)

# Теперь можно присваивать строки
for i in range(len(OD_matrix)):
    for j in range(len(OD_matrix)):
        if OD_matrix[i,j] == 0:
            OD_matrix[i,j] = ':'
        if mirror_matrix[i,j] == 0:
            mirror_matrix[i,j] = ':'

# Сохраняем исходную версию массива
airport_information_OG = copy.copy(airport_information)

# Проверяем и корректируем размеры sum_horizontal
if sum_horizontal.shape[1] < airport_information.shape[1]:
    sum_horizontal = np.pad(sum_horizontal, ((0, 0), (0, airport_information.shape[1] - sum_horizontal.shape[1])), mode='constant')


sum_vertical = np.zeros((len(OD_matrix), 1), dtype=float)
for i in range(len(OD_matrix)):
    # Преобразуем строку OD_matrix[i, :] в массив float, заменяя нечисловые значения на 0
    row = np.array([float(x) if x not in [':', 'n/a', 'None'] else 0.0 for x in OD_matrix[i, :]])
    sum_vertical[i] = np.sum(row)

# Создаем sum_horizontal (суммы столбцов OD_matrix)
sum_horizontal = np.zeros((len(OD_matrix), 1), dtype=float)
for j in range(len(OD_matrix)):
    # Преобразуем столбец OD_matrix[:, j] в массив float, заменяя нечисловые значения на 0
    column = np.array([float(x) if x not in [':', 'n/a', 'None'] else 0.0 for x in OD_matrix[:, j]])
    sum_horizontal[j] = np.sum(column)

# Создаем основную часть таблицы (нормальную матрицу)
total_matrix = np.append(sum_vertical, OD_matrix, axis=1)
total_matrix = np.append(airport_information, total_matrix, axis=1)

# Транспонируем airport_information для верхней части
airport_information_transposed = np.transpose(airport_information)  # Форма: (5, Num_airports)

# Преобразуем sum_horizontal, чтобы его форма была (1, Num_airports)
sum_horizontal_reshaped = sum_horizontal.T  # Форма: (1, Num_airports)

# Объединяем airport_information_transposed и sum_horizontal_reshaped по оси 0
airport_information_transposed = np.append(airport_information_transposed, sum_horizontal_reshaped, axis=0)  # Форма: (6, Num_airports)

# Создаем левый верхний угол
left_hand_corner = np.full((airport_information_transposed.shape[0], airport_information_transposed.shape[0]), '', dtype=object)

# Объединяем left_hand_corner и airport_information_transposed по оси 1
upper_bar = np.append(left_hand_corner, airport_information_transposed, axis=1)


if upper_bar.shape[1] > total_matrix.shape[1]:
    # Дополняем total_matrix пустыми строками
    total_matrix = np.pad(
        total_matrix,
        ((0, 0), (0, upper_bar.shape[1] - total_matrix.shape[1])),
        mode='constant',
        constant_values=''
    )
else:
    # Дополняем upper_bar пустыми строками
    upper_bar = np.pad(
        upper_bar,
        ((0, 0), (0, total_matrix.shape[1] - upper_bar.shape[1])),
        mode='constant',
        constant_values=''
    )

# Объединяем верхнюю и нижнюю части
total_matrix = np.append(upper_bar, total_matrix, axis=0)

# Восстанавливаем исходный массив
airport_information = copy.copy(airport_information_OG)

# Экспорт total_matrix в Excel
print('')
print('log. Сохранение полной матрицы')
df = pd.DataFrame(total_matrix).T
df.to_excel(excel_writer="TNDP_Russia/DM_air_matrix.xlsx")




#-----------------------------------------------------------------------
#construct total output matrix which is mirrored
# Конструируем итоговую матрицу

sum_vertical_mirror = np.zeros((len(mirror_matrix), 1), dtype=float)
for i in range(len(mirror_matrix)):
    # Преобразуем строку mirror_matrix[i, :] в массив float, заменяя нечисловые значения на 0
    row = np.array([float(x) if x not in [':', 'n/a', 'None'] else 0.0 for x in mirror_matrix[i, :]])
    sum_vertical_mirror[i] = np.sum(row)

# Создаем sum_horizontal_mirror (суммы столбцов mirror_matrix)
sum_horizontal_mirror = np.zeros((len(mirror_matrix), 1), dtype=float)
for j in range(len(mirror_matrix)):
    # Преобразуем столбец mirror_matrix[:, j] в массив float, заменяя нечисловые значения на 0
    column = np.array([float(x) if x not in [':', 'n/a', 'None'] else 0.0 for x in mirror_matrix[:, j]])
    sum_horizontal_mirror[j] = np.sum(column)

# Создаем основную часть таблицы (зеркальную матрицу)
total_mirror_matrix = np.append(sum_vertical_mirror, mirror_matrix, axis=1)
total_mirror_matrix = np.append(airport_information, total_mirror_matrix, axis=1)

# Транспонируем airport_information для верхней части
airport_information_transposed = np.transpose(airport_information)  # Форма: (5, Num_airports)

# Преобразуем sum_horizontal_mirror, чтобы его форма была (1, Num_airports)
sum_horizontal_mirror_reshaped = sum_horizontal_mirror.T  # Форма: (1, Num_airports)

# Объединяем airport_information_transposed и sum_horizontal_mirror_reshaped по оси 0
airport_information_transposed = np.append(airport_information_transposed, sum_horizontal_mirror_reshaped, axis=0)  # Форма: (6, Num_airports)

# Создаем левый верхний угол
left_hand_corner_mirror = np.full((airport_information_transposed.shape[0], airport_information_transposed.shape[0]), '', dtype=object)

# Объединяем left_hand_corner_mirror и airport_information_transposed по оси 1
upper_bar_mirror = np.append(left_hand_corner_mirror, airport_information_transposed, axis=1)


if upper_bar_mirror.shape[1] > total_mirror_matrix.shape[1]:
    # Дополняем total_mirror_matrix нулями
    total_mirror_matrix = np.pad(
        total_mirror_matrix,
        ((0, 0), (0, upper_bar_mirror.shape[1] - total_mirror_matrix.shape[1])),
        mode='constant', constant_values=''
    )
else:
    # Дополняем upper_bar_mirror нулями
    upper_bar_mirror = np.pad(
        upper_bar_mirror,
        ((0, 0), (0, total_mirror_matrix.shape[1] - upper_bar_mirror.shape[1])),
        mode='constant', constant_values=''
    )

# Объединяем верхнюю и нижнюю части
total_mirror_matrix = np.append(upper_bar_mirror, total_mirror_matrix, axis=0)

# Восстанавливаем исходный массив
airport_information = copy.copy(airport_information_OG)

# Экспорт total_mirror_matrix в Excel
print('')
print('log. Сохранение отзеркаленой матрицы')
df = pd.DataFrame(total_mirror_matrix).T
df.to_excel(excel_writer="TNDP_Russia/DM_air_matrix_mirror.xlsx")            
            
#export total_matrix to excel
print('')
print('log. Сохранение полной матрицы полетов')

freq_matrix = np.zeros_like(OD_matrix_frequency, dtype=float)
for i in range(OD_matrix_frequency.shape[0]):
    for j in range(OD_matrix_frequency.shape[1]):
        value = OD_matrix_frequency[i, j]
        if value in [': ', 'n/a', 'None', None, ':']:
            freq_matrix[i, j] = 0.0  # или другое значение по умолчанию
        else:
            try:
                freq_matrix[i, j] = float(value)
            except (ValueError, TypeError):
                freq_matrix[i, j] = 0.0
                print(f"Нечисловое значение в [{i},{j}]: {value}")
                
                
df = pd.DataFrame(freq_matrix, index=list_of_airports, columns=list_of_airports)
df.to_excel(excel_writer="TNDP_Russia/freq_air_matrix.xlsx", sheet_name="Flight Frequencies")
            
#export matrix to excel
print('')
print ('log. Сохранение зеркальной матрицы полетов')

mirror_freq_matrix = np.zeros_like(mirror_matrix_freq, dtype=float)
for i in range(mirror_matrix_freq.shape[0]):
    for j in range(mirror_matrix_freq.shape[1]):
        value = mirror_matrix_freq[i, j]
        if value in [':', 'n/a', 'None', None]:
            mirror_freq_matrix[i, j] = 0.0  # или другое значение по умолчанию
        else:
            try:
                mirror_freq_matrix[i, j] = float(value)
            except (ValueError, TypeError):
                mirror_freq_matrix[i, j] = 0.0
                print(f"Нечисловое значение в [{i},{j}]: {value}")
                
df_mirror = pd.DataFrame(mirror_freq_matrix, index=list_of_airports, columns=list_of_airports)
df_mirror.to_excel(excel_writer="TNDP_Russia/freq_air_matrix_mirror.xlsx", sheet_name="Mirrored Flight Frequencies")

print('')
print ('Well done')



'''--------------------------------------------------------
-----------------------   ШAГ 3    -------------------------
------------    Очистка и модификация данных -----------'''


print('')
print('log. ШAГ 3  ------  Расчёт расстояний между городами и аэропортами')

print('')
print('log. Добавляем данные по аэропортам, делая исключения:')
print('     добавляем только русские АП (Европы нет)')

print('')
print(airport_information)

# correction=0
# for x in range(len(airport_information)):   
#     if airport_information[x-correction][3] != 'Russia':
#         airport_information = np.delete(airport_information, x-correction, axis = 0)
#         mirror_matrix = np.delete(mirror_matrix, x-correction, axis = 0)
#         mirror_matrix = np.delete(mirror_matrix, x-correction, axis = 1)
#         mirror_matrix_freq = np.delete(mirror_matrix_freq, x-correction, axis = 0)
#         mirror_matrix_freq = np.delete(mirror_matrix_freq, x-correction, axis = 1)
#         correction = correction + 1
#import core cities
print('')
print('log. Загружаем таблицу растояний ЖД путей российских городов ') 
wb = load_workbook(r'TNDP_Russia/Core_cities_geography.xlsx')
ws_vertices = wb['Duration_road']

length_V = 8
range_len_V = range(8)

V = np.array([[i.value for i in j] for j in ws_vertices['G2':'N2']]) 
XX = np.zeros(length_V, dtype=object)
for i in range(len(XX)):
    XX[i] = V[0,i]
V = XX #vertex latitudes

V_country = np.array([[i.value for i in j] for j in ws_vertices['G3':'N3']]) 
XX = np.zeros(length_V, dtype=object)
for i in range(len(XX)):
    XX[i] = V_country[0,i]
V_country = XX #vertex latitudes

V_pop = np.array([[i.value for i in j] for j in ws_vertices['G6':'N6']]) 
XX = np.zeros(length_V)
for i in range(len(XX)):
    XX[i] = V_pop[0,i]
V_pop = XX #vertex populations

V_lat = np.array([[i.value for i in j] for j in ws_vertices['G4':'N4']]) 
XX = np.zeros(length_V)
for i in range(len(XX)):
    XX[i] = V_lat[0,i]
V_lat = XX #vertex latitudes

V_lon = np.array([[i.value for i in j] for j in ws_vertices['G5':'N5']]) 
XX = np.zeros(length_V)
for i in range(len(XX)):
    XX[i] = V_lon[0,i]
V_lon = XX #vertex longitudes 

print(airport_information)
#sort latitudes and longitudes based on airport information list
Airport_lat = np.zeros((len(airport_information)))
Airport_lon = np.zeros((len(airport_information)))
for row in range(len(Airport_lon)):
    Airport_lat[row] = airport_information[row][3]
    Airport_lon[row] = airport_information[row][4]
    
#build OD matrices for airport to city
City_to_Airport_Distance = np.zeros((len(V), len(airport_information))) #greater circle distance
City_to_Airport_Duration = np.zeros((len(V), len(airport_information))) #duration by car

#define haversine formula
def haversine(lat_i, lon_i, lat_j, lon_j):

    # convert decimal degrees to radians 
    lat_i, lon_i, lat_j, lon_j = map(radians, [lat_i, lon_i, lat_j, lon_j])

    # haversine formula 
    lat_delta = (lat_j - lat_i) 
    lon_delta = (lon_j - lon_i) 
    sr = sqrt(sin(lat_delta/2)**2 + cos(lat_i) * cos(lat_j) * sin(lon_delta/2)**2)
    R_earth = 6371.000 #radius earth [m]
    ds_gc = R_earth * 2 * asin(sr)
    return ds_gc #[km]

#calculate greater circle distances matrix
print('')
print( 'build city-to-airport distance matrix' )
for i in range(len(City_to_Airport_Distance)):
    for j in range(len(City_to_Airport_Distance[0])):
        City_to_Airport_Distance[i,j] = haversine(V_lat[i],V_lon[i],airport_information[:,3][j],airport_information[:,4][j])

df = pd.DataFrame(City_to_Airport_Distance)
df.to_excel(excel_writer = "TNDP_Russia/City_to_Airport_Distance.xlsx")



'''--------------------------------------------------------
-----------------------   ШAГ 4    -------------------------
--- Расчет времени доступа между городами и аэропортами ---'''

print('')
print ('log. Загрузка CIty - to - Airport расстояний через API(OpenRouteService)' )

#define infeasible city-airport combinations
for i in range(len(City_to_Airport_Distance)):
    for j in range(len(City_to_Airport_Distance[0])):
        if City_to_Airport_Distance[i,j] > maximum_potential_access_egress_dist:
            City_to_Airport_Duration[i,j] = float('inf')
'''
#export duration matrix to allow for in-between saving
df = pd.DataFrame(CI_to_AP_dura)
df.to_excel(excel_writer = "C:/Users/909448/Documents/##AFSTUDEREN/modelling/Netwerk/CI_to_AP_dura.xlsx")
'''
            
wb = load_workbook(r'TNDP_Russia/City_to_Airport_Duration.xlsx')
ws = wb['Sheet1']
City_to_Airport_Duration = np.array([[i.value for i in j] for j in ws['B2':'I9']],dtype=object)


print('')
print(airport_information)

#determine access and egress times
saver = 0
counter = 0
for v in range(len(City_to_Airport_Duration)):
    for ap in range(len(City_to_Airport_Duration[0])):
        value = City_to_Airport_Duration[v, ap]
        # Преобразуем значение в float
        if isinstance(value, str):
            if value.lower() == "inf":
                value = float('inf')
            else:
                value = float(value)
        else:
            value = float(value)
        if value < 0.000001:
            print('')
            print ('от' , ap,airport_information[:,1][ap],'(города) до ',v,V[v], ' (аэропорта)',)
            randomizer = 0
            
            distance = haversine(V_lat[v], V_lon[v], float(airport_information[ap, 3]), float(airport_information[ap, 4]))
            if distance * 1000 > 5900000:  # Преобразование в метры и проверка лимита 5900 км
                City_to_Airport_Duration[v, ap] = float('inf')
                print(f"Distance {distance:.2f} km exceeds limit, setting to inf")
                continue
            
            if randomizer == 0:
                coords = ((V_lon[v],V_lat[v]), (airport_information[:,4][ap],airport_information[:,3][ap]))
                # print('')
                # print(coords)
                # print('')
                client = openrouteservice.Client(key='5b3ce3597851110001cf62481a403ebd11a94d8fbf2320c2f8eac293') # Specify your personal API key
                routes = client.directions(
                    coordinates=coords,
                    profile='driving-car',
                    radiuses=[10000,10000],
                    format='geojson',
                    validate=False,
                )
                # print(routes)
            # if randomizer == 1:
            #     coords = ((V_lon[v],V_lat[v]), (airport_information[:,7][ap],airport_information[:,6][ap]))
            #     client = openrouteservice.Client(key='5b3ce3597851110001cf62483b780fd5c5dd455da89cc2eaa17543d8') # Specify your personal API key
            #     routes = client.directions(
            #         coordinates=coords,
            #         profile='driving-car',
            #         radiuses=[4000,4000],
            #         format='geojson',
            #         validate=False,
            #     )
            # if randomizer == 2:
            #     coords = ((V_lon[v],V_lat[v]), (airport_information[:,7][ap],airport_information[:,6][ap]))
            #     client = openrouteservice.Client(key='5b3ce3597851110001cf62481e0b82020ae745c2ae7b6642eabcc7a9') # Specify your personal API key
            #     routes = client.directions(
            #         coordinates=coords,
            #         profile='driving-car',
            #         radiuses=[4000,4000],
            #         format='geojson',
            #         validate=False,
            #     )
            print(routes['features'][0]['properties']['segments'][0]['duration']/60/60)
            if routes['features'][0]['properties']['segments'][0]['duration']/60/60 < max_catchment_area_perimeter:
                City_to_Airport_Duration[v,ap] = routes['features'][0]['properties']['segments'][0]['duration']/60/60
            else:
                City_to_Airport_Duration[v,ap] = float('inf')
            counter = counter+1
            saver = saver+1
            print( counter, '(',v,',',ap,')')#,coords
            if saver == 19:
                df_dur = pd.DataFrame(City_to_Airport_Duration)
                df_dur.to_excel(excel_writer = "TNDP_Russia/City_to_Airport_Duration.xlsx")
                
                
df_dur = pd.DataFrame(City_to_Airport_Duration)
df_dur.to_excel(excel_writer = "TNDP_Russia/City_to_Airport_Duration.xlsx")
          
df_dur = pd.DataFrame(airport_information)
df_dur.to_excel(excel_writer = "TNDP_Russia/airport_information.xlsx")

"""

#""" #only activated in first run
print('')
print('log. Загрузка City - to - Airport матрицы расстояний')
wb = load_workbook(r'TNDP_Russia/City_to_Airport_Duration.xlsx')
ws  =    wb['Sheet1']
CI_to_AP_dura = np.array([[i.value for i in j] for j in ws['B2':'I9']],dtype=object)