
import requests
from bs4 import BeautifulSoup

class Euro(float):

    def __new__(cls):

        r = requests.get('https://www.nbp.pl/home.aspx?f=/kursy/kursya.html')

        soup = BeautifulSoup(r.text, 'html.parser')

        soup.find('table', {'class': 'pad5'})
        b = soup.find_all('td', {'class': "bgt2 right"})
    
        slownik = {}
        for i in range(int(len(b)/2)):

            slownik[b[i*2].text] = b[2*i + 1].text
    
        slownik['1 EUR'] = slownik['1 EUR'].replace(',', '.')
        
        return float.__new__(cls, slownik['1 EUR'])