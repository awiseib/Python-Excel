
import xlwings as xw
from threading import Thread
from queue import Queue
from time import sleep
from ibapi.ticktype import TickType, TickTypeEnum
from ibapi.client import *
from ibapi.wrapper import *

BN = 'Python_Excel.xlsx'
LD = 'Live_Data'

def createBook():
    bk = xw.Book()
    bk.save(BN)
    bk.activate(BN)
    bk.sheets.add(name=LD)
    bk.sheets('Sheet1').delete()
    bk.save(BN)

def buildHeaders():
    q.put([LD,'A1',"Symbol"])
    for i in range(9):
        x = letterIncr(i)
        q.put([LD,'{}1'.format(x),TickTypeEnum.toStr(i)])

def letterIncr(letter_int):
    incr_letter = chr(ord('@')+letter_int+2)
    return incr_letter

def writeToWorkbook():
    while True:
        params = q.get()
        book = params[0]
        cell = params[1]
        content = params[2]
        xw.Book(BN).sheets[book].range(cell).value = content
        print(params)

q = Queue()

for i in range(50):
    t = Thread(
        target=writeToWorkbook, 
        daemon=True
      ).start()
q.join()


try:
    xw.Book(BN)
except FileNotFoundError:
    createBook()

if __name__ == "__main__":
    buildHeaders()