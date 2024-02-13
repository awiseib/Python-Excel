from LiveData import *

port = 7496

buildHeaders()
app = TestApp()
app.connect("127.0.0.1", port, 100)

sleep(1)
Thread(target=app.run).start()
sleep(3)

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