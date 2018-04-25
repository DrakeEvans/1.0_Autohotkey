from win32com import client

xl = client.GetActiveObject("Excel.Application")
xl.ActiveWorkbook.ActiveSheet.Range("A1:B2").Select()