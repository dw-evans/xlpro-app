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
import argparse


import sys
import os
import logging

from xlpro import config

config = config.load()

from xlpro.server import xlproServer
from xlpro import file_lock
from xlpro import _utils
import psutil
from xlpro import errors
import debugpy


wd = Path(__file__).parent

__logging_dir = Path(config.logging_path).parent
if not __logging_dir.exists():
    __logging_dir.mkdir(parents=True, exist_ok=True)


logging.basicConfig(
    # filename= wd / 'log.log',   # The file where logs will be saved
    # filename=config.logging_path,   # The file where logs will be saved
    stream=sys.stdout,
    # level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    level=getattr(logging, config.logging_level),          # The log level (DEBUG, INFO, WARNING, etc.)
    format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)

logger = logging.getLogger()

# # Create a handler to output logs to stdout
# handler = logging.StreamHandler(sys.stdout)
# handler.setLevel(logging.DEBUG)  # Set the handler's log level

# # Create a formatter and attach it to the handler
# formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# handler.setFormatter(formatter)

# # Add the handler to the logger
# logger.addHandler(handler)

# the background loop to keep the process alive
loop = asyncio.new_event_loop()

SERVER:xlproServer = None

def is_server_pending_close():
    global SERVER
    # TODO - XXX - Not sure why this loop is called so much...
    try:
        return SERVER._is_pending_close
    except Exception as e:
        logger.warning(f"Exception encountered while checking is_pending_close: {e}")
        SERVER = xlproServer()
        return SERVER._is_pending_close

def is_parent_process_closed(pid):
    # Check if parent process still exists
    try:
        parent = psutil.Process(pid)
        if parent.status() == psutil.STATUS_ZOMBIE:
            return True
    except psutil.NoSuchProcess:
        return True

    return False


def serve():
    global DEBUGPY_PORT
    logger.debug(f"serve() being run at root directory: {os.getcwd()}")

    debugpy.listen((config.debug_ip, DEBUGPY_PORT),)

    logger.info(f"ready to receive connection to debugger at {(config.debug_ip, DEBUGPY_PORT)}...")
    # debugpy.wait_for_client()
    # logger.info(f"Client connected successfully at {(config.debug_ip, config.debug_port)}")

    xlpro_lock_fp = file_lock.get_xlpro_lockfile_path()
    if not xlpro_lock_fp.parent.exists():
        logger.warning(f"{xlpro_lock_fp.parent} does not exist, making parents")
        xlpro_lock_fp.parent.mkdir()

    pass
    global PARENT_PID
    global SERVER
    try:
        lock_file_handle = file_lock.acquire_file_and_write_pid(str(xlpro_lock_fp))
    except PermissionError as e:
        print("Could not acquire lock on file. Checking validity")
        pid = file_lock.check_existing_lock_and_pid(xlpro_lock_fp)
        if not pid:
            print("The process with the lock file is not alive.")
            raise Exception(f"Error in lock file '{xlpro_lock_fp}' please correct manually.")
        print("The process appears to be alive.")
        _utils.show_warning(
            "xlpro",
            f"""WARNING: Could not acquire the file lock.
  - Another xlpro instance appears to be running at PID: {pid}
  - Delete {xlpro_lock_fp} if this issue persists.
  - This will not have affected your current session if xlpro was already running.""")
        sys.exit(1)


    pythoncom.CoInitialize()

    # we register everything dynamically using the clsid only.
    # the progid must be registered separately with admin elevation
    clsid = pywintypes.IID(xlproServer._reg_clsid_)

    # overwrite the win32com server policy. Leaves in room to dispatch other objects...?
    # credit to xlwings library for this
    BaseDefaultPolicy = win32com.server.policy.DefaultPolicy
    class ServerWrapPolicy(BaseDefaultPolicy):
        def _CreateInstance_(self, reqClsid, reqIID):
            global SERVER
            if reqClsid == clsid:
                # fyi we wrap the clsid IID object (a com-compatible interface) around our COM server
                SERVER = xlproServer()
                return win32com.server.util.wrap(SERVER, reqIID)
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
    logger.info(f"xlpro server starting on PID: {os.getpid()}")
    SERVER = xlproServer()

    def tidy_up_lock_file():
        logger.info(f"Releasing lock file '{xlpro_lock_fp}' handle: '{lock_file_handle}'...")
        file_lock.close_file(handle=lock_file_handle)
        logger.info(f"Removing lock file '{xlpro_lock_fp}' handle: '{lock_file_handle}'...")
        os.remove(xlpro_lock_fp)
        pass

    while True:
        try:
            # wait with a timeout before checking for closedown signal
            rc = win32event.MsgWaitForMultipleObjects(
                (), 0, 10_000, win32event.QS_ALLEVENTS
            )
            if rc == win32event.WAIT_OBJECT_0:
                # message loop is mandatory
                pwm = pythoncom.PumpWaitingMessages()
            if is_server_pending_close():
                raise errors.ServerClosedException
            if PARENT_PID is not None:
                if is_parent_process_closed(PARENT_PID):
                    raise psutil.NoSuchProcess(PARENT_PID)
        except errors.ServerClosedException:
            logger.info("errors.ServerClosedException encountered. Closing the server...")
            tidy_up_lock_file()
            break
        except psutil.NoSuchProcess:
            logger.info("psutil.NoSuchProcess encountered. Parent process has closed. Closing the server...")
            tidy_up_lock_file()
            break
        except KeyboardInterrupt:
            logger.info("KeyboardInterrupt encountered. Closing the server...")
            tidy_up_lock_file()
            break

    pythoncom.CoRevokeClassObject(revokeId)
    pythoncom.CoUninitialize()

    logger.info("Graceful exit")
    sys.exit()




def main():
    global PARENT_PID
    global DEBUGPY_PORT

    parser = argparse.ArgumentParser(description="Run the xlpro COM server.")

    parser.add_argument("--debugpy_port", type=int, required=False, help="The port to configure for debugpy debugging")
    parser.add_argument("--parent_pid", type=int, required=False, help="The parent pid of the process for the script to monitor")

    args = parser.parse_args()
    PARENT_PID = args.parent_pid if args.parent_pid else None
    DEBUGPY_PORT = args.debugpy_port if args.debugpy_port else 5678
    
    try:
        serve()
    except Exception as e:
        print(e)
        input("Fatal error encountered. Press enter to exit")

if __name__ == "__main__":
    main()