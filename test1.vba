Option Explicit

On Error Resume Next

Test

Sub Test()
    '   create workbook
    
    Dim excelAPP
    Dim excelBook
    
    ' set them
    Set excelAPP = CreateObject("Excel.Application")
    Set excelBook = excelApp.WorkBooks.Open("C:\Users\kedwa\Desktop\test.xls", 0, True)
    
    excelAPP.Quit
    
    Set excelAPP = Nothing
    Set excelBook = Nothing
        
End Sub
