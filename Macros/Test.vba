Sub Test()
    '   create workbook
    
    Dim excelAPP
    Dim excelBook
    
    ' set them
    Set excelAPP = CreateObject("Excel.Application")
    Set excelBook = excelAPP.WorkBooks.Open("C:\Users\kedwa\Desktop\test.xlsm", 0, True)
    
    MsgBox "MyMacro is running..."
    
    'excelAPP.Quit
    
    Set excelAPP = Nothing
    Set excelBook = Nothing
        
End Sub
