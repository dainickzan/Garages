from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog #Оставь надежду, всяк сюда входяищй
import re
from pandas import read_excel, notna
def parser(file_path:str): #функция для парса таблицы в нужном формате TODO: докрутить формат, починить парс станций
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
                    elif notna(re.search(r'[-+]?\d+',str(current_obj))) and len(str(current_obj))==6 and notna(file.iat[row,file.shape[1]-1]):#чудо написанное в полночь, стоящее исключительно на костылях
                         result['stations'][str(c)]={'code':current_obj,'cords':str(file.iat[row,file.shape[1]-1]).split(',')}#запись в формате номер -> код и координаты
                         c+=1#увеличение счётчика
    return result#вывод

app = QApplication(list('0'))#костыль для отказа от лишних библиотек
class MainWindow(QMainWindow):#Дальше живут драконы
    def __init__(self):
        super().__init__()#Я не знаю зачем нужна эта строчка, но без неё qt не работает
        self.button = QPushButton()#Инициализация кнопки
        self.setCentralWidget(self.button)#Растяжение кнопки на всю площадь окна
        self.button.clicked.connect(self.click)#Прикручивание функции к кнопке
    def click(self):
         ##filename = str(QFileDialog.getOpenFileName())[2:-19]#Выбор файла и обрезание лишней фигни
          print(parser('xmp.xlsx'))
          
window = MainWindow()#Инициализация окна
window.show()#Отображение окна
app.exec()#Запуск рабочего цикла
