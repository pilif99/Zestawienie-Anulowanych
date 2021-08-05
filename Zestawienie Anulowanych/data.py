
import datetime as dt

class Data:

    def __init__(self, x = 0):

        self.data = dt.datetime.now()

        if x:

            a = self.data.year
            b = self.data.month
            c = self.data.day

            if b == 1:
                a -= 1
                b = 12
            else:
                b -= 1

            if c > 28:
                c = 28

            self.data = dt.datetime(a, b, c)

        self.data = self.data.strftime('%Y%m%d')

    def __str__(self):

        return self.data