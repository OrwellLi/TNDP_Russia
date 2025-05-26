list_of_airports = np.zeros(Num_airports, dtype=object)

for n in range(Num_airports):
    list_of_airports[n] = list_Airport_temp[n]

print('')
print('log. Загружаем характеристики аэропортов.')
print('по категории  : аэропорт по коду ИАТА') 
print('топографически: страна')
print('географически : широта, долгота')

wb_Aiport_info = load_workbook(r'airportinformation3.xlsx')
ws_Airport_info = wb_Aiport_info['airportdata']

name = np.array([[i.value for i in j] for j in ws_Airport_info['B2':'GPG2']]) 
XX = np.zeros(len(name[0]), dtype=object)
for i in range(len(XX)):
    XX[i] = str(copy.copy(name[0,i]))
name = XX #Название

country = np.array([[i.value for i in j] for j in ws_Airport_info['B4':'GPG4']]) 
XX = np.zeros(len(country[0]), dtype=object)
for i in range(len(XX)):
    XX[i] = copy.copy(country[0,i])
country = XX #Страна

IATA = np.array([[i.value for i in j] for j in ws_Airport_info['B7':'GPG7']]) 
XX = np.zeros(len(IATA[0]), dtype=object)
for i in range(len(XX)):
    XX[i] = str(copy.copy(IATA[0,i]))
IATA = XX #ИАТА-код

Aiport_longitude = np.array([[i.value for i in j] for j in ws_Airport_info['B8':'GPG8']]) 
XX = np.zeros(len(Aiport_longitude[0]))
for i in range(len(XX)):
    XX[i] = str(copy.copy(Aiport_longitude[0,i]))
Aiport_longitude = XX #широта

Aiport_latitude = np.array([[i.value for i in j] for j in ws_Airport_info['B9':'GPG9']]) 
XX = np.zeros(len(Aiport_latitude[0]))
for i in range(len(XX)):
    XX[i] = str(copy.copy(Aiport_latitude[0,i]))
Aiport_latitude = XX #долгота

airport_information = np.full((Num_airports, 5),'n/a', dtype=object) #[IATA, name, coun, AP_lat, AP_lon]

#заполняем таблицу
for row in range(len(list_of_airports)):
    airport_information[row][0] = list_of_airports[row]
    for column in range(len(name)):
        if IATA[column] == airport_information[row][0]:
            # airport_information[row][1] = IATA[column]
            airport_information[row][1] = name[column]
            airport_information[row][2] = country[column]
            airport_information[row][3] = Aiport_latitude[column]
            airport_information[row][4] = Aiport_longitude[column]

#обновляем последовательность ап в коде IATA
for n in range(Num_airports):
    list_of_airports[n] = airport_information[n][0]

#убираем строки из данных о частоте полетов, не указанными в Aipport_data, чтобы не столкнуться с ошибками
correction=0
for row in range(len(Airport_data_frequency)-1):
    if Airport_data_frequency[row+1-correction][3] not in list(list_of_airports) or Airport_data_frequency[row+1-correction][5] not in list(list_of_airports):
        Airport_data_frequency = np.delete(Airport_data_frequency, row+1-correction, axis = 0)
        correction = correction + 1

def Airport_index_num(a): #находим индекс аэропорта
    return np.where(list_of_airports==a)[0][0]

print ('')
print ('log. Комбинируем характеристики аэропорта в одну матрицу')

#строим Origin-Destinaion матрицу
OD_matrix = np.zeros((len(list_of_airports), len(list_of_airports)),dtype=object)
OD_matrix_frequency = np.zeros((len(list_of_airports), len(list_of_airports)),dtype=object)

for row in range(len(Airport_data)-1):
    j = Airport_index_num(Airport_data[row+1][3])
    i = Airport_index_num(Airport_data[row+1][5])
    pax = float(Airport_data[row+1][6])
    OD_matrix[i,j] = pax

#fill OD matrix with flight data (frequency)
for row in range(len(Airport_data_frequency)-1):
    j = Airport_index_num(Airport_data_frequency[row+1][3])
    i = Airport_index_num(Airport_data_frequency[row+1][5])
    freq_y = Airport_data_frequency[row+1][6]
    if freq_y == ': ':
        freq_y = 0
    OD_matrix_frequency[i,j] = float(freq_y) / 365. #from yearly to daily
    if freq_y == 0 and OD_matrix[i,j] > 0:
        OD_matrix_frequency[i,j] = -1 * float('inf')