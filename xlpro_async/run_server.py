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

from server import xlproServerAsync, ServerClosedException

import logging

import load_config
import sys

config = load_config.load_config()

import os

import file_lock
import utils

wd = Path(__file__).parent

__logging_dir = Path(config.logging_path).parent
if not __logging_dir.exists():
    __logging_dir.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    # filename= wd / 'log.log',   # The file where logs will be saved
    filename=config.logging_path,   # The file where logs will be saved
    # level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    level=getattr(logging, config.logging_level),          # The log level (DEBUG, INFO, WARNING, etc.)
    format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)

logger = logging.getLogger(__name__)

# the background loop to keep the process alive
loop = asyncio.new_event_loop()

def should_close_server():
    server = xlproServerAsync()
    return server._is_pending_close

def serve():
    try:
        lock_file_handle = file_lock.acquire_file_and_write_pid(config.xlpro_lock_path)
    except PermissionError as e:
        print("Could not acquire lock on file. Checking validity")
        if not file_lock.check_existing_lock_and_pid(config.xlpro_lock_path):
            print("The process with the lock file is not alive. ")
            raise Exception(f"Error in lock file '{config.xlpro_lock_path}' please correct manually.")
        print("The process appears to be alive.")
        utils.show_warning(
            "xlpro",
            f"""WARNING: Could not acquire the file lock.
  - Another xlpro instance appears to be running.
  - Delete {config.xlpro_lock_path} if this issue persists.
  - This will not have affected your current session if xlpro was already running.""")
        sys.exit(1)


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
    print(f"Loop starting on PID:{os.getpid()}")
    while True:
        try:
            # wait with a 1 sec timeout before checking for closedown signal
            rc = win32event.MsgWaitForMultipleObjects(
                (), 0, 1000, win32event.QS_ALLEVENTS
            )
            if rc == win32event.WAIT_OBJECT_0:
                # message loop is mandatory
                pwm = pythoncom.PumpWaitingMessages()
            if should_close_server():
                raise ServerClosedException
        except ServerClosedException:
            logger.info("ServerClosedException encountered. Closing the server...")
            logger.info(f"Releasing lock file '{config.xlpro_lock_path}' handle: '{lock_file_handle}'...")
            file_lock.close_file(handle=lock_file_handle)
            break

    pythoncom.CoRevokeClassObject(revokeId)
    pythoncom.CoUninitialize()

    logger.info("Graceful exit")
    sys.exit(1)

if __name__ == "__main__":
    serve()
