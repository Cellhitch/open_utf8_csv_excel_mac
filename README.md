## open_utf8_csv_excel_mac
Fixes Excel’s issue with opening UTF-8 CSV files by using a Mac automation + Python script
A simple tool to convert CSV files to UTF-8 with BOM encoding and open them directly in Microsoft Excel on macOS.

## ⚙️ Prerequisites
- macOS
- Python 3.x
- Microsoft Excel for Mac
- Required Python packages:
  ```bash
  pip install pandas

## 🚀 Quick Setup
1. Copy the script to a local folder
2. Configure Mac automator tool (Automator -> new -> Application -> Add "Run Shell Script" action ) 
![image](https://github.com/user-attachments/assets/940f024d-0309-4bb7-b397-e62daac85236)

- Set "Pass input" to "as arguments"
- Enter the script:
  ```bash
  for f in "$@"
  do
  /path/to/your/python3 "/path/to/your/csv2excel.py" "$f"
  done

3. Save the application (e.g., as "CSV2Excel")
4. Right click on csv file -> Open with -> Other -> your automator tool. 


## 📝 Usage
1. Find any CSV file
2. Right-click → Get Info
3. Under "Open with", select your saved Automator app
4. Click "Change All..." to make it the default for CSV files

## 🛠 Troubleshooting
- If Excel shows encoding errors, ensure pandas is installed correctly
- If the script doesn't run, check Python path in Automator
- For permission issues, grant necessary permissions in System Preferences
