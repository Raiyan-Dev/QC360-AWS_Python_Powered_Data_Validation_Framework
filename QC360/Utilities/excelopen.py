#import os
#from pathlib import Path
#absolutePath = Path('C:\\Users\\kzhl524\\Downloads\\QCTool_1\\reporting_config.xlsm').resolve()
#os.system(f'start excel.exe "{absolutePath}"')

# import os, os.path
# import win32com.client

# if os.path.exists("C:\Users\kzhl524\Downloads\QC Tool\QC Tool\reporting_config.xlsm"):
    # xl=win32com.client.Dispatch("Excel.Application")
    # xl.Workbooks.Open(os.path.abspath("C:\Users\kzhl524\Downloads\QC Tool\QC Tool\reporting_config.xlsm"), ReadOnly=1)
    # xl.Application.Run("reporting_config.xlsm!RunPython")
# ##    xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    # xl.Application.Quit() # Comment this out if your excel script closes
    # del xl

from openpyxl import load_workbook
wb = load_workbook(filename='/home/kzhl524/Tools/QCTool_1/reporting_config.xlsm', read_only=False, keep_vba=True)

