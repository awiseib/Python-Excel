# General Python Imports
from datetime import datetime
import time

# WEBAPI request imports
import json
import requests
import websocket
import ssl
import urllib3

# Excel utility imports
from utils import *

BASE_URL = "localhost:5001/v1/api" 
repeater = {}

# Ignore insecure error messages
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def init():
    # Initialize Session
    url = f'https://{BASE_URL}/iserver/auth/ssodh/init'
    headers = {"User-Agent": "python/3.11"}
    json_data = {
        "publish": True,
        "compete": True
    }
    init_request = requests.post(url=url, verify=False, headers=headers, json=json_data)
    response = json.dumps(init_request.json(), indent=2)
    print(response)

def scanner():
    # Call Market Scanner & Establish Conids
    method = 'POST'
    url = f'https://{BASE_URL}/hmds/scanner'
    headers = {"User-Agent": "python/3.11"}
    scan_body = {
        "instrument":"STK",
        "locations": "STK.US.MAJOR",
        "scanCode": "HOT_BY_VOLUME",
        "secType": "STK",
        "filters":[{}]
    }
    scanner_request = requests.post(url=url, verify=False, headers=headers, json=scan_body)
    contracts = scanner_request.json()["Contracts"]["Contract"]
    conids = []
    for i in range(len(contracts)):
        conids.append(str(contracts[i]["contractID"]))
    global CONIDS
    CONIDS = conids

def tickle():
    # Call Tickle for session token.
    url = f'https://{BASE_URL}/tickle'
    headers = {"User-Agent": "python/3.11"}
    tickle_request = requests.post(url=url, verify=False, headers=headers)
    return tickle_request.json()['session']


def on_message(ws, message):
    jmsg = json.loads(message.decode('utf-8'))
    print(jmsg)
    cid = jmsg["conidEx"]
    q.put([LD, 'B{}'.format(repeater[cid]), jmsg["84"]])
    q.put([LD, 'C{}'.format(repeater[cid]), jmsg["86"]])
    q.put([LD, 'D{}'.format(repeater[cid]), jmsg["31"]])

def on_error(ws, error):
    print(f"Error: {error}")

def on_close(ws, er1, er2):
    print("##CLOSED##")
    print(er1, er2)

def on_open(ws):
    print("Opened Connection")
    time.sleep(3)
    conids = CONIDS

    for row, conid in enumerate(conids):
        repeater[conid] = row+2
        q.put([LD, 'A{}'.format(row+2), conid])
        ws.send('smd+'+conid+'+{"fields":["31","84","86"]}')


if __name__ == "__main__":
    buildHeaders()
    init()
    session_token = tickle()
    scanner()

    ws = websocket.WebSocketApp(
        url=f"wss://{BASE_URL}/ws",
        on_open=on_open,
        on_message=on_message,
        on_error=on_error,
        on_close=on_close,
        header=["User-Agent: python/3.11"],
        cookie=f"api={session_token}"
    )
    ws.run_forever(sslopt={"cert_reqs": ssl.CERT_NONE})