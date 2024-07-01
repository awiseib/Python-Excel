
# Python to Excel Imports
import xlwings as xw
from threading import Thread
from queue import Queue

BN = 'WebAPI_MD.xlsx'
LD = 'Sheet1'

def createBook():
    bk = xw.Book()
    bk.save(BN)
    bk.activate(BN)
    

def buildHeaders():
    headers = ["Conid", "Bid", "Ask", "Last"]
    for i in range(len(headers)):
      x = letterIncr(i)
      q.put([LD,'{}1'.format(x),headers[i]])

def letterIncr(letterInt):
    incrLetter = chr(ord('@')+letterInt+1)
    return incrLetter

def write_to_workbook():
    while True:
        params = q.get()
        book = params[0]
        cell = params[1]
        content = params[2]
        xw.Book(BN).sheets[book].range(cell).value = content
        print(params)

q = Queue()

for i in range(100):
    t = Thread(
        target=write_to_workbook, 
        daemon=True
      ).start()
q.join()

try:
    xw.Book(BN)
except FileNotFoundError:
    createBook()