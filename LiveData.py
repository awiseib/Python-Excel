from ibapi.client import *
from ibapi.wrapper import *
from ibapi.ticktype import TickType, TickTypeEnum

from utils import *

class TestApp(EClient, EWrapper):
    def __init__(self):
      EClient.__init__(self, self)
      self.orderId = 0

    def nextValidId(self, orderId: int):
      self.orderId = orderId

    def tickPrice(self, reqId: TickerId, tickType: TickType, price: float, attrib: TickAttrib,):
      if tickType < 9:
        col = letterIncr(tickType)
        q.put([LD,"{}{}".format(col, reqId) , str(price)])
    
    def tickSize(self, reqId: TickerId, tickType: TickType, size: float):
      if tickType < 9:
        col = letterIncr(tickType)
        q.put([LD,"{}{}".format(col, reqId), str(size)])
        
if __name__ == '__main__':
  app = TestApp()
  app.connect("127.0.0.1", 7497, 1000)
  app.run()