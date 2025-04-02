import pandas as pd
import subprocess
import os
import sys

def convert_and_open_excel(csv_path):
    # Read the CSV file
    df = pd.read_csv(csv_path, encoding='utf-8')
    
    # Save back with UTF-8-BOM encoding
    df.to_csv(csv_path, encoding='utf-8-sig', index=False)
    
    # Open with Excel and quit script
    apple_script = f'''
    tell application "Microsoft Excel"
        launch
        delay 0.5
        activate
        open "{csv_path}"
    end tell
    '''
    
    try:
        subprocess.run(['osascript', '-e', apple_script])
        print(f"Converted to UTF-8 with BOM and opened: {csv_path}")
        sys.exit(0)  # Exit after successful opening
    except Exception as e:
        print(f"Error opening Excel: {e}")
        sys.exit(1)  # Exit if there's an error

if __name__ == '__main__':
    # Handle files dropped onto the app
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]
    else:
        # If no file is dropped, show file picker
        from tkinter import filedialog, Tk
        root = Tk()
        root.withdraw()  # Hide the main window
        csv_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        
    if csv_file and os.path.exists(csv_file):
        convert_and_open_excel(csv_file)
