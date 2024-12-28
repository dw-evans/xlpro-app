from __future__ import annotations
from pathlib import Path

import threading
import asyncio

import pythoncom
import win32com.client
import win32com.server.util
import win32com.server.policy
import win32api
import pywintypes
import win32event

from server import xlproServerAsync

import logging

wd = Path(__file__).parent


logging.basicConfig(
    filename= wd / 'log.log',   # The file where logs will be saved
    level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)


logger = logging.getLogger(__name__)

# the background loop to keep the process alive
loop = asyncio.new_event_loop()

def _start_background_loop():
    # Background loop keeps the server live forever.
    # pythoncom.CoInitialize()
    asyncio.set_event_loop(loop)
    loop.run_forever()

def serve():
    pythoncom.CoInitialize()
    clsid = pywintypes.IID(xlproServerAsync._reg_clsid_)

    BaseDefaultPolicy = win32com.server.policy.DefaultPolicy
    class MyPolicy(BaseDefaultPolicy):
        def _CreateInstance_(self, reqClsid, reqIID):
            if reqClsid == clsid:
                return win32com.server.util.wrap(xlproServerAsync(), reqIID)
            else:
                return BaseDefaultPolicy._CreateInstance_(self, clsid, reqIID)
    # class MyPolicy(BaseDefaultPolicy):
    #     def _CreateInstance_(self, reqClsid, reqIID):
    #         global singleton
    #         if reqClsid == clsid:
    #             singleton = win32com.server.util.wrap(xlproServerAsync(), reqIID)
    #             return singleton
    #         else:
    #             return singleton

    win32com.server.policy.DefaultPolicy = MyPolicy

    factory = pythoncom.MakePyFactory(clsid)

    clsctx = pythoncom.CLSCTX_LOCAL_SERVER
    flags = pythoncom.REGCLS_MULTIPLEUSE | pythoncom.REGCLS_SUSPENDED
    revokeId = pythoncom.CoRegisterClassObject(clsid, factory, clsctx, flags)

    pythoncom.EnableQuitMessage(win32api.GetCurrentThreadId())
    pythoncom.CoResumeClassObjects()

    print("Loop starting, ctrl+c to exit")
    while True:
        try:
            rc = win32event.MsgWaitForMultipleObjects(
                (), 0, win32event.INFINITE, win32event.QS_ALLEVENTS
            )
            if rc == win32event.WAIT_OBJECT_0:
                if pythoncom.PumpWaitingMessages():
                    break  # wm_quit
        except KeyboardInterrupt:
            break

    pythoncom.CoRevokeClassObject(revokeId)
    pythoncom.CoUninitialize()

if __name__ == "__main__":
    serve()
