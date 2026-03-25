from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog #Оставь надежду, всяк сюда входяищй
from re import search
from pandas import DataFrame, Timedelta, Timestamp, read_excel, notna, to_datetime
from requests import get


api_key = '' #!!ОБЯЗАТЕЛЬНО ЗАПОЛНИТЬ, ПЕРЕД ИСПОЛЬЗОВАНИЕМ!!


def parser(file_path:str): #функция для парса таблицы в нужном формате
    file = read_excel(file_path).dropna(axis=1,how='all').dropna(how='all') #читает ексельку и сразу выпиливает пустые столбцы
    result={'data':{}, 'stations':{}} #словарь-мусоровоз для всех записей
    c = 0 #счётчик для станций
    for row in range(file.shape[0]):
              for col in range(file.shape[1]):#перебор оставшихся объектов по полученным размерам таблицы
                current_obj=file.iat[row,col]
                if notna(current_obj):#проверка на наличие объекта в ячейке, т.к. полное отсеивание сломает формат
                    if 'Станция отправления:' in str(current_obj):#начало проверок на шапку, лаконично
                          result['data']['from']={'code':file.iat[row,col+2], 'name':file.iat[row,col+3]}#Запись, шаблонная
                    elif 'Станция назначения:' in str(current_obj):
                          result['data']['to']={'code':file.iat[row,col+2], 'name':file.iat[row,col+3]}
                    elif 'Дата расчета:' in str(current_obj):
                          result['data']['date']=file.iat[row,col+3]
                    elif 'Расстояние' in str(current_obj):
                          result['data']['dist']=int(file.iat[row,col][12:-3])
                    elif 'Время' in str(current_obj):
                          result['data']['time']=int(file.iat[row,col][7:-6])
                    elif notna(search(r'[-+]?\d+',str(current_obj))) and len(str(current_obj))==6 and notna(file.iat[row,file.shape[1]-1]):#чудо написанное в полночь, стоящее исключительно на костылях
                         result['stations'][str(c)]={'code':current_obj,'cords':str(file.iat[row,file.shape[1]-1]),'dist':int(file.iat[row,3])}#запись в формате номер -> код и координаты, расстояние до конечной
                         c+=1#увеличение счётчика
    return result#вывод

def weatherget(time, lat, long): #функция для запроса к openweathermap
      payload = {'lat': lat, 'lon': long, 'dt':((time-Timedelta(hours=3)) - Timestamp("1970-01-01")) // Timedelta('1s'),'appid':api_key, 'units':'metric'}  #запись аргументов для обращения
      r = get('https://api.openweathermap.org/data/3.0/onecall/timemachine', params=payload) #get запрос серверу
      return min(r.json().get('hourly',[]), key=lambda x:abs(x['dt']-payload['dt']))['temp'] #вывод ответа

def table(result, filename):#Подсчёт времени по формуле и выведение в нужном формате
     temp = 0 #счётчик для температуры
     for i in result['stations']:
          temp = temp + (weatherget(result['date'], map(float, result['stations'][str(i)]['cords'].split(sep=',')))*(int(result['stations'][str(i)]['dist'])/result['dist'])) #получение температуры и умножение её на коэффициент "значимости"
     ore = 79.5*(0.5-temp)*1300/250000 #Для руды
     konc = 79.5*(0.5-temp)*1240/250000 #Для концентрата
     izvest = 79.5*(temp)*850/250000 #Для извести
     dol = 79.5*(temp)*900/250000 #Для доломита
     coal = 79.5*(0.5-temp)*1100/250000 #Для угля
     fin = DataFrame({'':'Время разморозки', #форматирование
            'Руда':str(object=ore//60)+':'+str(ore%60),
            'Концентрат':str(konc//60)+':'+str(konc%60),
            'Известь':str(izvest//60)+':'+str(izvest%60),
            'Доломит':str(dol//60)+':'+str(dol%60),
            'Уголь':str(coal//60)+':'+str(coal%60)})
     fin.to_excel(filename) #сохранение в ексель
     return 1



app = QApplication([])#костыль для отказа от лишних библиотек
class MainWindow(QMainWindow):#Дальше живут драконы
    def __init__(self):
        super().__init__()#Я не знаю зачем нужна эта строчка, но без неё qt не работает
        self.setWindowTitle("Выберите файл")#установка названия окна
        self.button = QPushButton('Нажмите для выбора файла')#Инициализация кнопки
        self.setCentralWidget(self.button)#Растяжение кнопки на всю площадь окна
        self.button.clicked.connect(self.click)#Прикручивание функции к кнопке

    def click(self):
         filename = str(QFileDialog.getOpenFileName())[2:-19]#Выбор файла и обрезание тех информации
         table(parser(filename),filename[:-4]+'_результаты.xlsx') #использование функций для чтения файла, рассчётов и записи в файл с названием в формате оригинальное название + _результаты
         self.button.setText("Результаты сохранены в "+filename.split('/')[-1][:-5]+'_результаты.xlsx')#изменение текста кнопки
         self.setWindowTitle("Полный путь до файла: "+filename[:-5]+'_результаты.xlsx')#изменение названия окна

window = MainWindow()#Инициализация окна
window.show()#Отображение окна
app.exec()#Запуск рабочего цикла
