
import pandas
import pyautogui as gui, time
import pyperclip
import numpy as np
import data
import euro

class Logowanie:

    def logowanie(self):

        screenWidth, screenHeight = gui.size()
        gui.moveTo(5,screenHeight-5)
        gui.click()
        time.sleep(0.5)
        gui.typewrite('eden nowszy (2)', interval=0.1)
        gui.press('enter')
        time.sleep(0.5)
        gui.typewrite('florenz', interval=0.1)
        gui.press('enter')
        gui.typewrite('florenz', interval=0.1)
        gui.press('enter')
        gui.press('enter')

class Pobieranie_Danych:

    def pobieranie_danych(self, nazwa, marka):

        if nazwa == 288:
            
            time.sleep(0.1)
            gui.click(85, 372)
            time.sleep(0.1)
            gui.click(85, 210)
            gui.click(85, 210)
            time.sleep(0.1)
            gui.click(622, 321)
            gui.typewrite('288', interval=0.01)
            gui.press('enter')
            gui.click(651, 274)
            lista = ['%', 'pl', 'pln', marka, '%', "'0','1', '2', '3','4','5','6','7','8','9','10','11','12'"]
            for i in range(len(lista)):
                if i != 1:
                    gui.press('enter')
                gui.press('enter')
                gui.press('enter')
                gui.press('f2')
                gui.typewrite(lista[i], interval=0.01)
            gui.click(1002, 623)
            time.sleep(5)
            gui.keyDown('ctrl')
            gui.press('a')
            gui.keyUp('ctrl')
            time.sleep(1)
            gui.keyDown('ctrl')
            gui.press('c')
            gui.keyUp('ctrl')

        elif nazwa == 'zrealizowane':
            
            time.sleep(0.1)
            gui.click(85, 372)
            time.sleep(0.1)
            gui.mouseDown(227, 187)
            gui.mouseUp(227, 240)
            gui.click(76, 161)
            gui.click(76, 161)
            gui.click(654, 724)
            time.sleep(0.1)
            gui.click(973, 284)
            gui.keyDown('backspace')
            gui.click(487, 284)
            gui.typewrite(self.data2.data + '..' + self.data.data, interval=0.01)
            gui.click(760, 284)
            gui.typewrite('dlewandowska', interval=0.01)
            gui.click(661, 354)
            for i in range(7):
                gui.press('down')
            gui.press('enter')
            gui.click(1207, 354)
            #time.sleep(30)
            time.sleep(15)
            gui.click(406, 417)
            gui.keyDown('ctrl')
            gui.press('c')
            gui.keyUp('ctrl')
            #time.sleep(100)
            time.sleep(5)

        elif nazwa == 'anulowane':

            time.sleep(0.1)
            gui.click(85, 372)
            time.sleep(0.1)
            gui.click(85, 210)
            gui.click(85, 210)
            time.sleep(0.1)
            gui.click(622, 321)
            gui.typewrite('388', interval=0.01)
            gui.press('enter')
            gui.click(651, 274)
            gui.press('enter')
            gui.press('enter')
            gui.press('F2')
            gui.typewrite(self.data2.data, interval=0.01)
            gui.press('enter')
            gui.press('enter')
            gui.press('enter')
            gui.typewrite(self.data.data, interval=0.01)
            gui.click(1002, 700)
            time.sleep(5)
            gui.keyDown('ctrl')
            gui.press('a')
            gui.keyUp('ctrl')
            gui.keyDown('ctrl')
            gui.press('c')
            gui.keyUp('ctrl')
            #time.sleep(200)
            time.sleep(15)

class Wklejanie_Danych:

    def wklejanie_danych(self):

        a = pyperclip.paste().split('\n')
        for i in range(len(a)):
            a[i] = a[i].split('\t')
            c = len(a[i]) - 1
            a[i][c] = ''.join(list(a[i][c])[:-1])
        self.object = pandas.DataFrame(a[1:-1], columns = a[0])
        self.object = self.object.replace(np.nan, '', regex=True)
        self.object = self.object.replace('"', '', regex=True)

class Edycja_Danych:

    def edycja_danych(self, nazwa):

        if nazwa == 'zrealizowane':

            self.object['Ilosc'] = self.object['Ilosc'].replace(',00', '', regex=True)
            self.object['Ilosc'] = self.object['Ilosc'].astype(int)
            self.object['Sprzedaż'] = self.object['Sprzedaż'].replace(',', '.', regex=True)
            self.object['Sprzedaż'] = self.object['Sprzedaż'].astype(float)
            
        elif nazwa == 'anulowane':

            self.object['gNiezrealizowanaIlosc'] = self.object['gNiezrealizowanaIlosc'].replace(',00', '', regex=True)
            self.object['gNiezrealizowanaIlosc'] = self.object['gNiezrealizowanaIlosc'].astype(int)
            self.object['gNiezrealizowanaWartosc'] = self.object['gNiezrealizowanaWartosc'].replace(',', '.', regex=True)
            self.object['gNiezrealizowanaWartosc'] = self.object['gNiezrealizowanaWartosc'].astype(float)
            self.object['lCenaNetPoRab'] = self.object['lCenaNetPoRab'].replace(',', '.', regex=True)
            self.object['lCenaNetPoRab'] = self.object['lCenaNetPoRab'].astype(float)
            
            EUR = euro.Euro()

            self.object['gNiezrealizowanaWartosc'] = round(self.object['gNiezrealizowanaWartosc'] * EUR, 2)
            self.object['lCenaNetPoRab'] = round(self.object['lCenaNetPoRab'] * EUR, 2)

class Eden(pandas.DataFrame, Logowanie, Pobieranie_Danych, Wklejanie_Danych, Edycja_Danych):

    def __init__(self, nazwa, marka = 'broger'):

        super().__init__()
        
        self.data = data.Data()
        self.data2 = data.Data(1)
        
        self.logowanie()
        self.pobieranie_danych(nazwa, marka)
        self.wklejanie_danych()
        self.edycja_danych(nazwa)

        self.append(super().__init__(self.object))

        self.zamkniecie_edena()

    def zamkniecie_edena(self):

        time.sleep(10)
        gui.click(1900, 10)
        gui.press('enter')

#pandas.set_option("display.max_columns", None, 'display.width', 1000)
