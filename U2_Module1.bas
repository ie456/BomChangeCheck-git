Attribute VB_Name = "U2_Module1"
Sub CreatDefaulFile(fileName As Variant, filePath As Variant)
    
 
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    
    Set Wb = Application.Workbooks.add(1)
    
    
    Wb.Sheets(1).Cells(1, 1).value = "test"
    
    
    
    Wb.SaveAs fileName:=filePath & "\" & fileName & ".csv", FileFormat:=xlCSV
    Wb.Close
    
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    
    MsgBox "Done"
    
    Unload UserForm2
    
    Call Module1.start
    
End Sub
