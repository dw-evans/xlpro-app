import pythoncom
import win32com.server
from win32com.server.exception import COMException
import threading

class PersistingCOMServer:
    _reg_clsid_ = '{299E805E-168A-4FA6-BFF4-17C190EFD35F}'  # Replace with your GUID
    _reg_desc_ = "Persisting COM Server"
    _reg_progid_ = "Python.PersistingCOMServer"
    _public_methods_ = ['store_data', 'get_data', 'clear_data']

    def __init__(self):
        # Data to persist across calls
        self.data = {}
        print("PersistingCOMServer initialized.")

    def store_data(self, key, value):
        """ Store data in the server's state """
        self.data[key] = value
        return f"Stored: {key} = {value}"

    def get_data(self, key):
        """ Retrieve data from the server's state """
        return self.data.get(key, None)

    def clear_data(self):
        """ Clear all data in the server """
        self.data.clear()
        return "Cleared all stored data."

    def run_server(self):
        """ Run the server to keep it alive """
        pythoncom.PumpMessages()



if __name__ == "__main__":
    # Register the COM server
    import sys
    if any([x in sys.argv for x in ('--register', '--unregister', '--debug')]):
        import win32com.server.register
        win32com.server.register.UseCommandLine(PersistingCOMServer)
        sys.exit()

    # Create an instance of the server object
    server_instance = PersistingCOMServer()
    # Keep the COM server alive and responsive
    print("COM server is now running and waiting for requests...")
    pythoncom.PumpMessages()
