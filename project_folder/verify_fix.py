import win32com.client as win32
import sys

try:
    print("Attempting to initialize Excel.Application...")
    # This forces early binding if the cache is working
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # Since Excel might be running in background, we just close it
    excel.Quit()
    print("Success: Excel initialized and closed correctly!")
except Exception as e:
    print(f"Failed: {e}")
    sys.exit(1)
