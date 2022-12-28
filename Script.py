from win32com import client

app = client.DispatchEx("Excel.Application")
app.interactive = False
app.Visible = False

#we need to take input for path of excel file

path = input ("Introduce la ruta del archivo")
print("Conirtiendo a PDF")
workbook = app.Workbooks.Open(path)
workbook.ActiveSheet.ExportAsFixedFormat(0,path)
workbook.Close()

print("Listo!!")