import winreg as reg
from server import xlproServer
import sys

"""
Completes the registry mappings for the dynamic COM server.

Run this script with administrator rights to configure the registry key.
This is only required because we cannot register a dynamic local COM server
with a progid.
"""

def register_progid_to_clsid(progid, clsid):
    try:
        # Map ProgID to CLSID
        progid_key_path = f"{progid}\\CLSID"
        with reg.CreateKey(reg.HKEY_CLASSES_ROOT, progid_key_path) as progid_key:
            reg.SetValue(progid_key, "", reg.REG_SZ, clsid)
            print(f"Successfully mapped ProgID '{progid}' to CLSID '{clsid}'. {progid_key_path}")
        
        # Map CLSID to ProgID (Optional)
        clsid_key_path = f"CLSID\\{clsid}\\ProgID"
        with reg.CreateKey(reg.HKEY_CURRENT_USER, clsid_key_path) as clsid_key:
            reg.SetValue(clsid_key, "", reg.REG_SZ, progid)
            print(f"Successfully mapped CLSID '{clsid}' to ProgID '{progid}'. {clsid_key_path}")
        
        print(f"Successfully registered ProgID '{progid}' to CLSID '{clsid}'.")

    except Exception as e:
        print(f"Error registering ProgID to CLSID: {e}")

def unregister_progid_to_clsid(progid, clsid):
    try:
        # Map ProgID to CLSID
        progid_key_path = f"{progid}\\CLSID"
        try:
            reg.DeleteKey(reg.HKEY_CURRENT_USER, progid_key_path) 
            print(f"Successfully removed ProgID '{progid}' to CLSID '{clsid}'. {progid_key_path}")
        except Exception as e:
            print(f"{e}")
        
        # Map CLSID to ProgID (Optional)
        clsid_key_path = f"CLSID\\{clsid}\\ProgID"
        try:
            reg.DeleteKey(reg.HKEY_CURRENT_USER, clsid_key_path)
        except Exception as e:
            print(f"{e}")
        
        print(f"Successfully deregistered ProgID '{progid}' to CLSID '{clsid}'.")
        

    except Exception as e:
        print(f"Error registering ProgID to CLSID: {e}")

# Example usage
if __name__ == "__main__":
    progid = xlproServer._reg_progid_
    clsid = xlproServer._reg_clsid_
    if "--unregister" in sys.argv:
        unregister_progid_to_clsid(progid, clsid)
    else:
        register_progid_to_clsid(progid, clsid)
    input("Press enter to continue...")
