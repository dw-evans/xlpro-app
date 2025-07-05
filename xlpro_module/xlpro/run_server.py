from __future__ import annotations

def disable_quickedit():
    '''
    Disable quickedit mode on Windows terminal. quickedit prevents script to
    run without user pressing keys..'''
    import os
    if not os.name == 'posix':
        try:
            import msvcrt
            import ctypes
            kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            device = r'\\.\CONIN$'
            with open(device, 'r') as con:
                hCon = msvcrt.get_osfhandle(con.fileno())
                kernel32.SetConsoleMode(hCon, 0x0080)
        except Exception as e:
            print('Cannot disable QuickEdit mode! ' + str(e))
            print('.. As a consequence the script might be automatically\
            paused on Windows terminal')

    pass

disable_quickedit()


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

import contextlib
import io
import rich




@contextlib.contextmanager
def suppress_output():
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        yield

wd = Path(__file__).parent

# __logging_dir = Path(config.logging_path).parent
# if not __logging_dir.exists():
#     __logging_dir.mkdir(parents=True, exist_ok=True)


logging.basicConfig(
    # filename= wd / 'log.log',   # The file where logs will be saved
    # filename=config.logging_path,   # The file where logs will be saved
    stream=sys.stdout,
    # level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    level=getattr(logging, config.LOGGING_LEVEL),          # The log level (DEBUG, INFO, WARNING, etc.)
    # format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    format='%(levelname)s - %(message)s',  # The format of log messages
    # datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
    datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)
logger = logging.getLogger()

# from rich.logging import RichHandler

# # Basic configuration for logging
# logging.basicConfig(
#     level="DEBUG",  # or INFO, WARNING, etc.
#     format="%(message)s",
#     datefmt="[%X]",
#     handlers=[RichHandler()]
# )

# logger = logging.getLogger("rich_logger")




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

def get_server():
    global SERVER
    try:
        return SERVER
    except Exception as e:
        SERVER = xlproServer()
        return SERVER


def is_server_pending_close():
    return get_server()._is_pending_close

def is_parent_process_closed(pid):
    # Check if parent process still exists
    try:
        parent = psutil.Process(pid)
        if parent.status() == psutil.STATUS_ZOMBIE:
            return True
    except psutil.NoSuchProcess:
        return True

    return False


def quickedit(enabled=1): # This is a patch to the system that sometimes hangs
        import ctypes
        '''
        Enable or disable quick edit mode to prevent system hangs, sometimes when using remote desktop
        Param (Enabled)
        enabled = 1(default), enable quick edit mode in python console
        enabled = 0, disable quick edit mode in python console
        '''
        # -10 is input handle => STD_INPUT_HANDLE (DWORD) -10 | https://learn.microsoft.com/en-us/windows/console/getstdhandle
        # default = (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x200)
        # 0x40 is quick edit, #0x20 is insert mode
        # 0x8 is disabled by default
        # https://learn.microsoft.com/en-us/windows/console/setconsolemode
        kernel32 = ctypes.windll.kernel32
        if enabled:
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))
            print("Console Quick Edit Enabled")
        else:
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x00|0x100))
            print("Console Quick Edit Disabled")

        pass

def serve():
    global DEBUGPY_PORT
    global CLSID
    global WORKBOOK_NAME

    import xlpro
    s = f"""

        ██╗  ██╗██╗     ██████╗ ██████╗  ██████╗ 
        ╚██╗██╔╝██║     ██╔══██╗██╔══██╗██╔═══██╗
         ╚███╔╝ ██║     ██████╔╝██████╔╝██║   ██║
         ██╔██╗ ██║     ██╔═══╝ ██╔══██╗██║   ██║
        ██╔╝ ██╗███████╗██║     ██║  ██║╚██████╔╝
        ╚═╝  ╚═╝╚══════╝╚═╝     ╚═╝  ╚═╝ ╚═════╝ 

            xlpro-server v{xlpro.__version__}
            Copyright (c) 2025 Daniel Evans
            License: MIT. Free for commercial use.

    """
    print(f"{s}")
    # from rich import print
    # print(f"{s}", style="#526cfe")

    logger.info(f"Starting xlpro server...")
    logger.info(f"Server parameters: CLSID='{CLSID}', DEBUGPY_PORT={DEBUGPY_PORT}, WORKBOOK_NAME='{WORKBOOK_NAME}'")
    logger.debug(f"xlpro.run_server.main() being run with CLSID='{CLSID}', DEBUGPY_PORT={DEBUGPY_PORT}, WORKBOOK_NAME='{WORKBOOK_NAME}'")

    # logger.debug(f"serve() being run at root directory: {os.getcwd()}")

    debugpy.listen(('localhost', DEBUGPY_PORT))

    logger.info(f"Ready to receive connection to debugger at {("localhost", DEBUGPY_PORT)}...")

    xlpro_lock_fp = file_lock.get_xlpro_lockfile_path_parent() / f"{WORKBOOK_NAME}.xlpro.lock"

    if not xlpro_lock_fp.parent.exists():
        logger.debug(f"{xlpro_lock_fp.parent} does not exist, making parents")
        xlpro_lock_fp.parent.mkdir()

    pass
    # global PARENT_PID
    global SERVER
    try:
        lock_file_handle = file_lock.acquire_file_and_write_datas(str(xlpro_lock_fp), guid=CLSID, debugpy_port=DEBUGPY_PORT)
    except PermissionError as e:
        logger.debug("Could not acquire lock on file. Checking validity")
        lockfile_contents_dict = file_lock.check_lockfile_get_contents_as_dict_if_alive(xlpro_lock_fp)
        if not lockfile_contents_dict:
            logger.debug("The process with the lock file is not alive.")
            raise Exception(f"Error in lock file '{xlpro_lock_fp}' please correct manually.")
        logger.debug("The process appears to be alive.")
        pid, guid, debugpy_port = [lockfile_contents_dict.get(x) for x in ("pid", "guid", "debugpy_port")]
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
    # clsid = pywintypes.IID(xlproServer._reg_clsid_)
    clsid = pywintypes.IID(CLSID)

    # overwrite the win32com server policy. Leaves in room to dispatch other objects...?
    # credit to xlwings library for this
    BaseDefaultPolicy = win32com.server.policy.DefaultPolicy
    class ServerWrapPolicy(BaseDefaultPolicy):
        def _CreateInstance_(self, reqClsid, reqIID):
            global SERVER
            if reqClsid == clsid:
                # fyi we wrap the clsid IID object (a com-compatible interface) around our COM server
                SERVER = xlproServer()
                xlproServer._reg_clsid_ = clsid
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
    logger.debug(f"xlpro server starting on PID: {os.getpid()}")
    SERVER = xlproServer()

    # print("XLPROSTART_TRIGGER_OK")
    logger.info("Startup OK, ready for synchronisation. Sending signal")
    sys.stderr.write("XLPROSTART_TRIGGER_OK\n")
    sys.stderr.flush()
    sys.stdout.write("XLPROSTART_TRIGGER_OK\n")
    sys.stdout.flush()
    logger.info("Signal sent.")

    def tidy_up_lock_file():
        logger.info("Cleaning up lock file...")
        logger.debug(f"Releasing lock file '{xlpro_lock_fp}' handle: '{lock_file_handle}'...")
        file_lock.close_file(handle=lock_file_handle)
        logger.debug(f"Removing lock file '{xlpro_lock_fp}' handle: '{lock_file_handle}'...")
        os.remove(xlpro_lock_fp)
        logger.info("Successfully cleaned up lockfile")
        pass

    while True:
        try:
            # wait with a timeout before checking for closedown signal
            rc = win32event.MsgWaitForMultipleObjects(
                (), 0, 10_000, win32event.QS_ALLEVENTS
            )
            if rc == win32event.WAIT_OBJECT_0:
                # message loop is mandatory
                with suppress_output():
                    pythoncom.PumpWaitingMessages()
            if is_server_pending_close():
                raise errors.ServerClosedException
        except errors.ServerClosedException:
            logger.critical("errors.ServerClosedException encountered. Closing the server...")
            break
        except psutil.NoSuchProcess:
            logger.critical("psutil.NoSuchProcess encountered. Parent process has closed. Closing the server...")
            break
        except KeyboardInterrupt:
            logger.critical("KeyboardInterrupt encountered. Closing the server...")
            break
        except Exception as e:
            logger.critical(f"Uncaught exception: {e}. Closing the server...")
            break
    
    try:
        tidy_up_lock_file()
    except:
        logger.warning("Error during lockfile cleanup, investigate if issues reloading persist.")

    pythoncom.CoRevokeClassObject(revokeId)
    pythoncom.CoUninitialize()

    logger.info("Graceful shutdown. Program exiting...")
    input("Press Enter to exit")
    sys.exit()


def main():
    try:

        # global PARENT_PID
        global DEBUGPY_PORT
        global CLSID
        global WORKBOOK_NAME

        parser = argparse.ArgumentParser(description="Run the xlpro COM server.")
        
        # default_clsid = '{122BB48A-57EF-4775-A28C-3F71ED0D02A7}'

        parser.add_argument("--workbook_path", type=str, required=True, help="The workbook")
        # parser.add_argument("--guid", type=str, required=False, help="The CLSID to run the server on", default=default_clsid)
        parser.add_argument("--debugpy_port", type=int, required=True, help="The port to configure for debugpy debugging")
        # parser.add_argument("--parent_pid", type=int, required=False, help="The parent pid of the process for the script to monitor", default=None)

        args = parser.parse_args()
        # PARENT_PID = args.parent_pid
        DEBUGPY_PORT = args.debugpy_port
        CLSID = pythoncom.CreateGuid()
        WORKBOOK_NAME = Path(args.workbook_path).name
    
        serve()
    except Exception as e:
        logger.critical(f"Fatal Exception encountered: {e}")
        logger.critical(f"Closing down...")
        input("Press enter to exit")

    finally:
        sys.exit()

if __name__ == "__main__":
    # sys.argv = ["run_server.py", "--workbook_path", r"C:\Users\Daniel Evans\projects\xlpro\xlpro_examples\test.xlsx", "--debugpy_port", "5678"]
    main()