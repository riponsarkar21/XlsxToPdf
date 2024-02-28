from win32com import client

app = client.Dispatch("Excel.Application")
app.Interactive = False
app.Visible = False

path = input("Enter the path of excel here: ")
print ("Converting. please wait....!")

workbook = app.Workbooks.Open(path)
workbook.ActiveSheet.ExportAsFixedFormat(0,path)
workbook.Close()
    
print("completed!!")
