import xlwings as xw
import os

try:
    app = xw.App(visible=False)
    wb = app.books.add()
    ws = wb.sheets[0]
    ws.range('A1').formula = "='[temp_transformed.xlsx]Current_Portfolio'!A1"
    print("Before:", ws.range('A1').formula)
    
    # Try replacing
    ws.api.Cells.Replace("[temp_transformed.xlsx]", "")
    print("After:", ws.range('A1').formula)
    
    wb.close()
    app.quit()
except Exception as e:
    print("Error:", e)
