# python_utils
Python_utils repository contains some scirpt utils written by python.  
The first util is reayExcelToWrod_v1.py which read the Excel file(.xlsx .xls) and get the value of specific cells. Then, write the values to format strings and save the strings to Word file. The word file name is automatically generated by date that read from the Excel cell.
## Set up python environment
Install python env on your windows system https://www.python.org/downloads/windows/  
The script support python3 env.
## Install python library
Some python libraries(openpyxl, python-docx) need to be installed in your system for supporting python script to run.
### Linux/Mac OS
pip3 install openpyxl  
pip3 install python-docx  
### Windows
py -m pip install openpyxl  
py -m pip install python-docx
## How to run
python <python_file_name> <Excel_file_path>
### Example
#### Windows
python .\readExceclToWord_v1.py "C:\Users\robhuang\Downloads\test_py\MTM 2024.7.3.xlsx"
#### Linux/Mac OS
python3 readExceclToWord_v1.py /Users/hrp/Downloads/MTM\ 2024.7.3.xlsx


