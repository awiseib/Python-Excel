from decimal import Decimal
import xlwings as xw
from threading import Thread
import time
from queue import Queue

from ibapi.common import BarData, TickerId
from ibapi.client import *
from ibapi.wrapper import *
from ibapi.ticktype import TickType, TickTypeEnum

'''
This file is used as a variant of the pexcel.py, LiveData.py, and utils.py bundle by compiling all of the information into a single file.
There is no intention to update this file beyond what is currently available, so if there are any features of interest please see the spearate file structure. 
'''

port = 7496

LD = 'Live_Data'

class TestApp(EClient, EWrapper):
    def __init__(self):
        EClient.__init__(self, self)
        self.orderId = 0

    def nextValidId(self, orderId: OrderId):
        self.orderId = orderId
    
    def tickPrice(self, reqId: TickerId, tickType: TickType, price: float, attrib: TickAttrib,):
        if tickType < 9:
          col = letterIncr(tickType)
          q.put([LD,"{}{}".format(col, reqId) , price])
    
    def tickSize(self, reqId: TickerId, tickType: TickType, size: Decimal):
        if tickType < 9:
          col = letterIncr(tickType)
          q.put([LD,"{}{}".format(col, reqId), size])

    def historicalData(self, reqId: int, bar: BarData):
        print(reqId, bar)

    def error(self, reqId: TickerId, errorCode: int, errorString: str, advancedOrderRejectJson=""):
        q.put([LD,"J1", "{}: {}".format(errorCode, errorString)])

def buildHeaders():
    q.put([LD,'A1',"Symbol"])
    for i in range(9):
      x = letterIncr(i)
      q.put([LD,'{}1'.format(x),TickTypeEnum.toStr(i)])

def letterIncr(letterInt):
    incrLetter = chr(ord('@')+letterInt+2)
    return incrLetter

def write_to_workbook():
    while True:
        params = q.get()
        book = params[0]
        cell = params[1]
        content = params[2]
        xw.Book('Python_Sample.xlsx').sheets[book].range(cell).value = content
        print(params)

q = Queue()

for i in range(4):
    t = Thread(
        target=write_to_workbook, 
        daemon=True
      ).start()
q.join()

def main():
    buildHeaders()
    app = TestApp()
    app.connect("127.0.0.1", port, 100)

    time.sleep(1)
    Thread(target=app.run).start()
    time.sleep(3)

    mycontract = Contract()
    mycontract.exchange = "SMART"
    mycontract.secType = "STK"
    mycontract.currency = "USD"

    symbols = ['AAPL', 'IBKR', 'TSLA', 'MSFT']
    for enumer,symbol in enumerate(symbols):
        enumer += 2
        mycontract.symbol = symbol

        q.put([LD,"A%s" % enumer, symbol])

        app.reqMktData(
            reqId=enumer,
            contract=mycontract,
            genericTickList="",
            snapshot=False,
            regulatorySnapshot=False,
            mktDataOptions=[],
        )

if __name__ == '__main__':
    main()