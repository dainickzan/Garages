from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog #Оставь надежду, всяк сюда входяищй
import re
from pandas import read_excel, notna
def parser(file_path:str): #функция для парса таблицы в нужном формате TODO: докрутить формат, починить парс станций
    file = read_excel(file_path, header=None) #Чтение в дефолтном виде
    result = {'metadata': {}, 'stations': []} #создание динамического словаря для записи результатов
    for idx, row in file.iterrows(): #поиск шапки таблицы
        row_str = ' '.join(str(cell) for cell in row if notna(cell))
        
        if 'Станция отправления:' in row_str:
            station_from_code = file.iloc[idx, 3] if len(row) > 3 else None
            station_from_name = file.iloc[idx, 5] if len(row) > 5 else None
            result['metadata']['station_from'] = {
                'code': int(station_from_code),
                'name': station_from_name
            }
        
        elif 'Станция назначения:' in row_str:
            station_to_code = file.iloc[idx, 3] if len(row) > 3 else None
            station_to_name = file.iloc[idx, 5] if len(row) > 5 else None
            result['metadata']['station_to'] = {
                'code': int(station_to_code),
                'name': station_to_name
            }
        
        elif 'Дата расчета:' in row_str:
            date_value = file.iloc[idx, 5] if len(row) > 5 else None
            result['metadata']['calculation_date'] = date_value
        
        elif 'Расстояние' in row_str and 'км' in row_str:
            numbers = re.findall(r'\d+', row_str)
            if len(numbers) >= 2:
                result['metadata']['distance_km'] = int(numbers[0])
                result['metadata']['time_hours'] = int(numbers[1])
#    data_start_idx = None
 #   for idx, row in file.iterrows():
  #      first_cell = str(row[0]) if len(row) > 0 and notna(row[0]) else ''
   #     if first_cell.isdigit() and len(first_cell) == 6:
    #        data_start_idx = idx
     #       break
    #if data_start_idx is not None:
     #   headers = range(16)
        #for col in range(len(file.columns)):
         #   cell_value = file.iloc[data_start_idx, col]
          #  headers.append(cell_value if notna(cell_value) else f'Column_{col}')
        
      #  stations_data = file.iloc[data_start_idx:data_start_idx+2].reset_index(drop=True)
       # stations_data.columns = headers
        
        #for _, row in stations_data.iterrows():
         #   station_info = {}
          #  for col_name, value in row.items():
           #     if notna(value):
            #        station_info[col_name] = value
            
            #if headers[-1] in station_info:
             #   coords_str = str(station_info[headers[-1]])
              #  coords = re.findall(r'[-+]?\d*\.\d+|\d+', coords_str)
               # if len(coords) != 0:
                #    station_info['latitude'] = float(coords[0])
                 #   station_info['longitude'] = float(coords[1])
                  #  result['stations'].append(station_info)

    data_start_idx = None
    for idx, row in file.iterrows():
        first_cell = str(row[0]) if len(row) > 0 and notna(row[0]) else ''
        if first_cell.isdigit() and len(first_cell) == 6:
            data_start_idx = idx
            break
    
    # Парсинг всех станций с координатами (не только первых двух)
    if data_start_idx is not None:
        headers = range(16)  # Используем числовые индексы вместо заголовков
        
        # Проходим по всем строкам начиная с data_start_idx до конца файла
        for idx in range(data_start_idx, len(file)):
            row_data = file.iloc[idx]
            
            # Проверяем наличие координат в последнем столбце (предполагается что координаты там)
            coords_str = str(row_data[headers[-1]]) if len(row_data) > headers[-1] else ''
            
            # Ищем координаты в строке (может быть в любом столбце)
            coords_found_in_row = False
            
            # Создаем словарь для информации о станции
            station_info = {}
            
            # Проверяем все ячейки строки на наличие координат и собираем информацию о станции
            for col_idx in headers:
                value = row_data[col_idx] if col_idx < len(row_data) else None
                
                # Если значение не пустое - добавляем в информацию о станции
                if notna(value):
                    station_info[col_idx] = value
            
                    # Проверяем ячейку на наличие координат (формат XX.XXXXX YY.YYYYY)
                    cell_coords_match = re.findall(r'[-+]?\d*\.\d+|\d+', str(value))
                    if len(cell_coords_match) >= 2:
                        try:
                            lat_lon_found_in_cell_coords_match_positions=[]
                            lat_lon_found_in_cell_coords_match_positions.extend([float(cell_coords_match[0]), float(cell_coords_match[1])])
                            coords_found_in_row=True 
                        except ValueError: 
                            pass

    return {'From':result['metadata'].get('station_from', {}),
            'To':result['metadata'].get('station_to', {}),
            'Distance':result['metadata'].get('distance_km'),
            'H':result['metadata'].get('time_hours'),
            'Date':result['metadata'].get('calculation_date'),
            'Stations':result['stations']}


app = QApplication(list('0'))#костыль для отказа от лишних библиотек
class MainWindow(QMainWindow):#Дальше живут драконы
    def __init__(self):
        super().__init__()#Я не знаю зачем нужна эта строчка, но без неё qt не работает
        self.button = QPushButton()#Инициализация кнопки
        self.setCentralWidget(self.button)#Растяжение кнопки на всю площадь окна
        self.button.clicked.connect(self.click)#Прикручивание функции к кнопке
    def click(self):
         ##filename = str(QFileDialog.getOpenFileName())[2:-19]#Выбор файла и обрезание лишней фигни
          print(parser('xmp.xlsx')["Stations"])
          
window = MainWindow()#Инициализация окна
window.show()#Отображение окна
app.exec()#Запуск рабочего цикла
