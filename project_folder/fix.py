import win32com
import shutil
import os
import sys

try:
    # Get the real path of gen_py
    gen_path = win32com.__gen_path__
    print("--- check path ---")
    print(f"Path: {gen_path}")

    if os.path.exists(gen_path):
        # Remove the entire folder
        shutil.rmtree(gen_path)
        print("Success: Deleted gen_py folder!")
    else:
        print("Not found: Path does not exist")

except Exception as e:
    print(f"Error: {e}")
    print("Try running CMD as Administrator")

print("--- Done ---")