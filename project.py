from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog #Оставь надежду, всяк сюда входяищй
from rxls.reader import read
app = QApplication(list('0'))#костыль для отказа от лишних библиотек
class MainWindow(QMainWindow):#Дальше живут драконы
    def __init__(self):
        super().__init__()#Я не знаю зачем нужна эта строчка, но без неё qt не работает
        self.button = QPushButton()#Инициализация кнопки
        self.setCentralWidget(self.button)#Растяжение кнопки на всю площадь окна
        self.button.clicked.connect(self.click)#Прикручивание функции к кнопке
    def click(self):
          filename = str(QFileDialog.getOpenFileName())[2:-19]#Выбор файла и обрезание лишней фигни
          tempdb = read(filename,header=False).to_pandas() #Получение информации из таблицы
          
window = MainWindow()#Инициализация окна
window.show()#Отображение окна
app.exec()#Запуск рабочего цикла
