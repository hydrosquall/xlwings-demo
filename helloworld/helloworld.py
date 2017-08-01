import xlwings as xw


def hello_xlwings():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"

# VBA Script
# Once you have written a method, you can go to the developer tools
# and create a button which calls this method every time the button is clicked.


'''
Sub SampleCall()
    mymodule = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
    RunPython ("import " & mymodule & ";" & mymodule & ".hello_xlwings()")
End Sub
'''
