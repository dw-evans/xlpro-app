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


def serve():
    pythoncom.CoInitialize()

    # we register everything dynamically using the clsid only.
    # the progid must be registered separately with admin elevation
    clsid = pywintypes.IID(xlproServerAsync._reg_clsid_)

    # overwrite the win32com server policy. Leaves in room to dispatch other objects...?
    # credit to xlwings library for this
    BaseDefaultPolicy = win32com.server.policy.DefaultPolicy
    class ServerWrapPolicy(BaseDefaultPolicy):
        def _CreateInstance_(self, reqClsid, reqIID):
            if reqClsid == clsid:
                # fyi we wrap the clsid IID object (a com-compatible interface) around our COM server
                return win32com.server.util.wrap(xlproServerAsync(), reqIID)
            else:
                # return BaseDefaultPolicy._CreateInstance_(self, clsid, reqIID)
                # I don't actually know how we would even get in here...?
                raise Exception 
    win32com.server.policy.DefaultPolicy = ServerWrapPolicy

    factory = pythoncom.MakePyFactory(clsid)

    # definitely need to register as a multipleuse local server
    # so Dispatch calls return the same object. Seems to work fine with Dispatch calls, but getactiveobject
    # is the more elegant solution in theory. idc
    
    # Note that the class needs to be configured as a singleton regardless of the below settings
    clsctx = pythoncom.CLSCTX_LOCAL_SERVER
    flags = pythoncom.REGCLS_MULTIPLEUSE | pythoncom.REGCLS_SUSPENDED
    revokeId = pythoncom.CoRegisterClassObject(clsid, factory, clsctx, flags)

    pythoncom.EnableQuitMessage(win32api.GetCurrentThreadId()) # not sure why we would need this.
    pythoncom.CoResumeClassObjects() # I think this cancels the suspended operation

    # XXX fix the main loop exit seq
    print("Loop starting, ctrl+c to exit (Probably doesn't work)")
    while True:
        try:
            rc = win32event.MsgWaitForMultipleObjects(
                (), 0, win32event.INFINITE, win32event.QS_ALLEVENTS
            )
            if rc == win32event.WAIT_OBJECT_0:
                pwm = pythoncom.PumpWaitingMessages()
                # if pwm:
                #     break  # wm_quit
        except KeyboardInterrupt:
            break
        except Exception:
            pass

    pythoncom.CoRevokeClassObject(revokeId)
    pythoncom.CoUninitialize()

if __name__ == "__main__":
    serve()
