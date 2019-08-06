Attribute VB_Name = "subModule"
Sub Recovery()
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
End Sub
Sub open_ExcelFile(fileName As Variant, sheetName As Variant, activeBook As Variant, sheetSel As Variant)
    
    Workbooks.Open fileName:=fileName
    openActive = ActiveWorkbook.name
   
   If Sheets.Count = 1 Then
    Sheets.add after:=Sheets(Worksheets.Count)
   End If
   
   
   
    Workbooks(openActive).Sheets(sheetSel).Move after:=Workbooks(activeBook).Sheets(Workbooks(activeBook).Worksheets.Count)
    
    Workbooks(activeBook).Sheets(Worksheets.Count).name = sheetName
    Workbooks(activeBook).Sheets(Worksheets.Count).Tab.ColorIndex = 41
    
    tempAry = Split(fileName, "\")
    Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
    
End Sub
Sub open_ExcelFile_Gene(fileName As Variant, sheetName As Variant, activeBook As Variant, sheetSel As Variant, color As Integer)
    
    Workbooks.Open fileName:=fileName
    openActive = ActiveWorkbook.name
   
   If Sheets.Count = 1 Then
    Sheets.add after:=Sheets(Worksheets.Count)
   End If
   
   
   
    Workbooks(openActive).Sheets(sheetSel).Move after:=Workbooks(activeBook).Sheets(Workbooks(activeBook).Worksheets.Count)
    
    Workbooks(activeBook).Sheets(Worksheets.Count).name = sheetName
    Workbooks(activeBook).Sheets(Worksheets.Count).Tab.ColorIndex = color
    
    tempAry = Split(fileName, "\")
    Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
    
End Sub

Sub open_ConceptFile(fileName As Variant, sheetName As Variant, activeBook As Variant)
     
     On Error GoTo exit1
    Workbooks.Open fileName:=fileName
    
    'Rows("1:6").EntireRow.Delete
    
    Range("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1)), TrailingMinusNumbers:=True
    Range("A1").Select
    
    Sheets(1).Move after:=Workbooks(activeBook).Sheets(Workbooks(activeBook).Worksheets.Count)
    Workbooks(activeBook).Sheets(Worksheets.Count).name = sheetName
    Workbooks(activeBook).Sheets(Worksheets.Count).Tab.ColorIndex = 41
    
Exit Sub
    
exit1:

    MsgBox "Please check path : " & vbCrLf & _
    vbCrLf & _
    fileName & vbCrLf & _
    vbCrLf & _
    "After correct path, please reload."
End Sub

Sub open_OrcadFile()

End Sub
'------------------------------------------------------
'---This function is uesed for fill in vale on cells---
'------------------------------------------------------
Sub saveData(sheetName As Variant, data As Variant, location As Variant)

With Workbooks(tempWorkBookName)

    unProtectSheet (sheetName)
    
    .Worksheets(sheetName).Range(location).value = data

    protectSheet (sheetName)

End With

End Sub

Sub saveData_2(bookName As Variant, sheetName As Variant, data As Variant, location As Variant)

With Workbooks(bookName)

    unProtectSheet (sheetName)
    
    .Worksheets(sheetName).Range(location).value = data

    protectSheet (sheetName)

End With

End Sub

Sub getSavePath()  'For setting output file path

    Dim myFolder As FileDialog
    Set myFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    myFolder.InitialFileName = ThisWorkbook.Path
    myFolder.Show
    On Error GoTo exitEnd
    UserForm1.Label17 = myFolder.SelectedItems.Item(1)
    
    Call saveData("Main", myFolder.SelectedItems.Item(1), "L30")
    
exitEnd:
    
End Sub

Sub getFilePath_Multi() 'Sample for select file path
    functionModule.unProtectSheet ("Main")
    
    With Application.FileDialog(msoFileDialogOpen)
        '.InitialFileName = "D\¤å¥óÀÉ\ªþÄÝÀÉ®×"
        .AllowMultiSelect = True
        .Show
        For i = 1 To .SelectedItems.Count
            Worksheets("Main").Cells(i, 1) = .SelectedItems(i)
            MsgBox .SelectedItems(i)
        Next
    End With
    functionModule.protectSheet ("Main")
End Sub


Sub compareString(ByRef String1 As Variant, ByRef String2 As Variant, ignorType As Variant)
  
  Dim tempList1, tempList2
  
  
  tempAry = Split(String1, ignorType)
  Set tempList1 = aryToList(tempAry)
  
  tempAry = Split(String2, ignorType)
  Set tempList2 = aryToList(tempAry)
  
  Set tempList1Clone = tempList1.Clone
  
  
  
  For Each tempValue In tempList1
  
    If tempList2.Contains(tempValue) Then
    
        tempList2.removeat (tempList2.indexof(tempValue, 0))
        tempList1Clone.removeat (tempList1Clone.indexof(tempValue, 0))
    End If
  
  Next
  
  String1 = aryToString(tempList1Clone)
  
  String2 = aryToString(tempList2)
  
  
End Sub

Private Function aryToList(arr As Variant) As Object
    
    Set aryToList = CreateObject("system.collections.arraylist")
    
    For Each tempValue In arr
        aryToList.add tempValue
    Next
    
End Function

Private Function aryToString(arr As Variant) As String
    temp = ""
    
    For Each tempValue In arr
        temp = temp & tempValue & ","
    Next
    
    If temp <> "" Then temp = Left(temp, Len(temp) - 1)
    
    aryToString = temp
    
    
End Function

Sub saveSetData()

    Dim checkValue, tempCheckValue
    
    
    Set checkValue = CreateObject("system.collections.arraylist")
    Set tempCheckValue = CreateObject("system.collections.arraylist")
    
     With Workbooks(tempWorkBookName)
    
    
    tempCheckValue.add .Worksheets("Main").Range("A30").value
    tempCheckValue.add .Worksheets("Main").Range("A31").value
    tempCheckValue.add .Worksheets("Main").Range("A32").value
    tempCheckValue.add .Worksheets("Main").Range("A33").value
    
    tempCheckValue.add .Worksheets("Main").Range("B30").value
    tempCheckValue.add .Worksheets("Main").Range("B31").value
    tempCheckValue.add .Worksheets("Main").Range("B32").value
    tempCheckValue.add .Worksheets("Main").Range("B33").value
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    On Error GoTo exit1
     If (64 < Asc(UserForm1.TextBox5.value) And Asc(UserForm1.TextBox5.value) < 91) And (64 < Asc(UserForm1.TextBox6.value) And Asc(UserForm1.TextBox6.value) < 91) _
        And (64 < Asc(UserForm1.TextBox7.value) And Asc(UserForm1.TextBox7.value) < 91) Then
            functionModule.unProtectSheet ("Main")
            
            .Worksheets("Main").Range("A31").value = UserForm1.TextBox5.value
            .Worksheets("Main").Range("A32").value = UserForm1.TextBox6.value
            .Worksheets("Main").Range("A33").value = UserForm1.TextBox7.value
            functionModule.protectSheet ("Main")
            
            functionModule.updateUserFormValue 2
        End If
            
        If IsNumeric(UserForm1.TextBox19.value) Then
            If Int(UserForm1.TextBox19.value) / UserForm1.TextBox19.value = 1 Then
            Call subModule.saveData("Main", UserForm1.TextBox19.value, "A30")
            functionModule.updateUserFormValue 8
            End If
        End If
        
        
        If (64 < Asc(UserForm1.TextBox13.value) And Asc(UserForm1.TextBox13.value) < 91) And (64 < Asc(UserForm1.TextBox14.value) And Asc(UserForm1.TextBox14.value) < 91) _
        And (64 < Asc(UserForm1.TextBox15.value) And Asc(UserForm1.TextBox15.value) < 91) Then
            functionModule.unProtectSheet ("Main")
            
            .Worksheets("Main").Range("B31").value = UserForm1.TextBox13.value
            .Worksheets("Main").Range("B32").value = UserForm1.TextBox14.value
            .Worksheets("Main").Range("B33").value = UserForm1.TextBox15.value
            
            functionModule.protectSheet ("Main")
            
            functionModule.updateUserFormValue 3
        End If
        If IsNumeric(UserForm1.TextBox21.value) Then
            If Int(UserForm1.TextBox21.value) / UserForm1.TextBox21.value = 1 Then
            
                Call subModule.saveData("Main", UserForm1.TextBox21.value, "B30")
                
                functionModule.updateUserFormValue 9
            End If
        Else
            GoTo exit1
        End If
        
    tempCheckValue.add .Worksheets("Main").Range("A30").value
    tempCheckValue.add .Worksheets("Main").Range("A31").value
    tempCheckValue.add .Worksheets("Main").Range("A32").value
    tempCheckValue.add .Worksheets("Main").Range("A33").value
    
    tempCheckValue.add .Worksheets("Main").Range("B30").value
    tempCheckValue.add .Worksheets("Main").Range("B31").value
    tempCheckValue.add .Worksheets("Main").Range("B32").value
    tempCheckValue.add .Worksheets("Main").Range("B33").value
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
  
    For i = 0 To UBound(checkValue(0))
    
        If checkValue(0)(i) <> checkValue(1)(i) Then
            
            Call subModule.saveData("Main", False, "M30")
            'updateUserFormValue 7
            MsgBox "Saved!!"
            Exit For
        End If
        
    
    Next
    
    
    
    End With
 
    
     
     
exit1:

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub
Sub saveNewSetData()

    Dim checkValue, tempCheckValue
    
    
    Set checkValue = CreateObject("system.collections.arraylist")
    Set tempCheckValue = CreateObject("system.collections.arraylist")
    
     With Workbooks(tempWorkBookName)
    
    
    tempCheckValue.add .Worksheets("Main").Range("A30").value
    tempCheckValue.add .Worksheets("Main").Range("A31").value
    tempCheckValue.add .Worksheets("Main").Range("A32").value
    tempCheckValue.add .Worksheets("Main").Range("A33").value
    
    
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    On Error GoTo exit1
     If (64 < Asc(UserForm1.TextBox5.value) And Asc(UserForm1.TextBox5.value) < 91) And (64 < Asc(UserForm1.TextBox6.value) And Asc(UserForm1.TextBox6.value) < 91) _
        And (64 < Asc(UserForm1.TextBox7.value) And Asc(UserForm1.TextBox7.value) < 91) Then
            functionModule.unProtectSheet ("Main")
            
            .Worksheets("Main").Range("A31").value = UserForm1.TextBox5.value
            .Worksheets("Main").Range("A32").value = UserForm1.TextBox6.value
            .Worksheets("Main").Range("A33").value = UserForm1.TextBox7.value
            functionModule.protectSheet ("Main")
            
            functionModule.updateUserFormValue 2
        End If
            
        If IsNumeric(UserForm1.TextBox19.value) Then
            If Int(UserForm1.TextBox19.value) / UserForm1.TextBox19.value = 1 Then
            Call subModule.saveData("Main", UserForm1.TextBox19.value, "A30")
            functionModule.updateUserFormValue 8
            End If
        Else
            GoTo exit1
        End If
        
        
       
        
    tempCheckValue.add .Worksheets("Main").Range("A30").value
    tempCheckValue.add .Worksheets("Main").Range("A31").value
    tempCheckValue.add .Worksheets("Main").Range("A32").value
    tempCheckValue.add .Worksheets("Main").Range("A33").value
   
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
  
    For i = 0 To UBound(checkValue(0))
    
        If checkValue(0)(i) <> checkValue(1)(i) Then
            
            Call subModule.saveData("Main", False, "M30")
            'updateUserFormValue 7
            MsgBox "Saved!!"
            Exit For
        End If
        
    
    Next
    
    
    
    End With
 
    
     
     
exit1:

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub
Sub saveOldSetData()

    Dim checkValue, tempCheckValue
    
    
    Set checkValue = CreateObject("system.collections.arraylist")
    Set tempCheckValue = CreateObject("system.collections.arraylist")
    
     With Workbooks(tempWorkBookName)
    
    
    
    
    tempCheckValue.add .Worksheets("Main").Range("B30").value
    tempCheckValue.add .Worksheets("Main").Range("B31").value
    tempCheckValue.add .Worksheets("Main").Range("B32").value
    tempCheckValue.add .Worksheets("Main").Range("B33").value
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    On Error GoTo exit1
     
        
        
        If (64 < Asc(UserForm1.TextBox13.value) And Asc(UserForm1.TextBox13.value) < 91) And (64 < Asc(UserForm1.TextBox14.value) And Asc(UserForm1.TextBox14.value) < 91) _
        And (64 < Asc(UserForm1.TextBox15.value) And Asc(UserForm1.TextBox15.value) < 91) Then
            functionModule.unProtectSheet ("Main")
            
            .Worksheets("Main").Range("B31").value = UserForm1.TextBox13.value
            .Worksheets("Main").Range("B32").value = UserForm1.TextBox14.value
            .Worksheets("Main").Range("B33").value = UserForm1.TextBox15.value
            
            functionModule.protectSheet ("Main")
            
            functionModule.updateUserFormValue 3
        End If
        If IsNumeric(UserForm1.TextBox21.value) Then
            If Int(UserForm1.TextBox21.value) / UserForm1.TextBox21.value = 1 Then
            
                Call subModule.saveData("Main", UserForm1.TextBox21.value, "B30")
                
                functionModule.updateUserFormValue 9
            End If
        Else
            GoTo exit1
        End If
        
    
    
    tempCheckValue.add .Worksheets("Main").Range("B30").value
    tempCheckValue.add .Worksheets("Main").Range("B31").value
    tempCheckValue.add .Worksheets("Main").Range("B32").value
    tempCheckValue.add .Worksheets("Main").Range("B33").value
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
  
    For i = 0 To UBound(checkValue(0))
    
        If checkValue(0)(i) <> checkValue(1)(i) Then
            
            Call subModule.saveData("Main", False, "M31")
            'updateUserFormValue 7
            MsgBox "Saved!!"
            Exit For
        End If
        
    
    Next
    
    
    
    End With
 
    
     
     
exit1:

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub
Sub saveChangeSetData()

    Dim checkValue, tempCheckValue
    
    
    Set checkValue = CreateObject("system.collections.arraylist")
    Set tempCheckValue = CreateObject("system.collections.arraylist")
    
     With Workbooks(tempWorkBookName)
    
    
    tempCheckValue.add .Worksheets("Main").Range("C30").value
    tempCheckValue.add .Worksheets("Main").Range("C31").value
    tempCheckValue.add .Worksheets("Main").Range("C32").value
    tempCheckValue.add .Worksheets("Main").Range("C33").value
    
    
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    On Error GoTo exit1
        If (64 < Asc(UserForm1.TextBox28.value) And Asc(UserForm1.TextBox28.value) < 91) _
        And (64 < Asc(UserForm1.TextBox24.value) And Asc(UserForm1.TextBox24.value) < 91) _
        And (64 < Asc(UserForm1.TextBox26.value) And Asc(UserForm1.TextBox26.value) < 91) _
        And (64 < Asc(UserForm1.TextBox30.value) And Asc(UserForm1.TextBox30.value) < 91) Then
            functionModule.unProtectSheet ("Main")
            
            .Worksheets("Main").Range("C30").value = UserForm1.TextBox28.value
            .Worksheets("Main").Range("C31").value = UserForm1.TextBox24.value
            .Worksheets("Main").Range("C32").value = UserForm1.TextBox26.value
            .Worksheets("Main").Range("C33").value = UserForm1.TextBox30.value
            
            functionModule.protectSheet ("Main")
            
            functionModule.updateUserFormValue 11
        Else
            GoTo exit1
        End If
            
       
        
        
       
        
    tempCheckValue.add .Worksheets("Main").Range("C30").value
    tempCheckValue.add .Worksheets("Main").Range("C31").value
    tempCheckValue.add .Worksheets("Main").Range("C32").value
    tempCheckValue.add .Worksheets("Main").Range("C33").value
   
    
    checkValue.add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
  
    For i = 0 To UBound(checkValue(0))
    
        If checkValue(0)(i) <> checkValue(1)(i) Then
            
            Call subModule.saveData("Main", False, "M32")
            'updateUserFormValue 7
            MsgBox "Saved!!"
            Exit For
        End If
        
    
    Next
    
    
    
    End With
 
    
     
     
exit1:

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub
Sub addBTN(location As Variant, sheetName As String, onAct As String, captionValue As String)



    Dim btn As Button
    Dim tempRange As Range
    
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    
    With Workbooks(tempWorkBookName).Worksheets(sheetName)
    
    .Buttons.Delete
    
    Set tempRange = .Range(location)
    Set btn = .Buttons.add(tempRange.Left, tempRange.Top, tempRange.Width, tempRange.Height)


    With btn
    
        .OnAction = onAct
        .caption = captionValue
        .name = "Split_Integrat"
    End With
    
    
    
    End With
    
    
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    


End Sub

Sub splitLocation()
 
 Dim tempACT, tempPartNumber, tempQTY, tempLocation, tempDescription, tempNote1, tempfontValue
 Dim ACT, PartNumber, QTY, location, DESCRIPTION, Note1, fontValue
 Dim sheetName As String
    
  Set ACT = CreateObject("system.collections.arraylist")
  Set PartNumber = CreateObject("system.collections.arraylist")
  Set QTY = CreateObject("system.collections.arraylist")
  Set location = CreateObject("system.collections.arraylist")
  Set DESCRIPTION = CreateObject("system.collections.arraylist")
  Set Note1 = CreateObject("system.collections.arraylist")
  Set fontValue = CreateObject("system.collections.arraylist")
 
 
  Set tempACT = CreateObject("system.collections.arraylist")
  Set tempPartNumber = CreateObject("system.collections.arraylist")
  Set tempQTY = CreateObject("system.collections.arraylist")
  Set tempLocation = CreateObject("system.collections.arraylist")
  Set tempDescription = CreateObject("system.collections.arraylist")
  Set tempNote1 = CreateObject("system.collections.arraylist")
  Set tempfontValue = CreateObject("system.collections.arraylist")
  
  
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
 
  Index = 6
  sheetName = ActiveSheet.name
  
  
    With Workbooks(tempWorkBookName).Worksheets(sheetName)
  
    Do While .Cells(Index, 1) <> ""
    
        tempACT.add .Cells(Index, 1).value
        tempPartNumber.add .Cells(Index, 2).value
        tempQTY.add .Cells(Index, 3).value
        tempLocation.add .Cells(Index, 4).value
        tempDescription.add .Cells(Index, 5).value
        tempNote1.add .Cells(Index, 6).value
        tempfontValue.add .Cells(Index, 4).Font.Strikethrough
    
        Index = Index + 1
    
    Loop
    
    
    
    
    
    For i = 0 To tempPartNumber.Count - 1
    
        If tempLocation(i) = "" Then
            
            ACT.add tempACT(i)
            PartNumber.add tempPartNumber(i)
            QTY.add Abs(tempQTY(i))
            location.add tempLocation(i)
            DESCRIPTION.add tempDescription(i)
            Note1.add tempNote1(i)
            fontValue.add tempfontValue(i)
            
            
        Else
            tempAry = Split(tempLocation(i), ",")
            
            
            For Each tempLoc In tempAry
            
                ACT.add tempACT(i)
                PartNumber.add tempPartNumber(i)
                QTY.add 1
                location.add tempLoc
                DESCRIPTION.add tempDescription(i)
                Note1.add tempNote1(i)
                fontValue.add tempfontValue(i)
                
            Next
            
        End If
    
    
    
    Next
    
    
    
    .Range(Cells(6, 1), Cells(Index, 6)).Delete
    
    Call printDataInSheet(6, sheetName, ACT, PartNumber, QTY, location, DESCRIPTION, Note1, fontValue)
    
    
  End With

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    
    ACT.Clear
    PartNumber.Clear
    QTY.Clear
    location.Clear
    DESCRIPTION.Clear
    Note1.Clear
    fontValue.Clear
    
    tempACT.Clear
    tempPartNumber.Clear
    tempQTY.Clear
    tempLocation.Clear
    tempDescription.Clear
    tempNote1.Clear
    tempfontValue.Clear
    
    
  Call addBTN("D4", sheetName, "integrateLocation", "Integrate Location")
 
    
 
 
End Sub

Sub integrateLocation()

 Dim temp
 Dim ACT, PartNumber, QTY, location, DESCRIPTION, Note1, fontValue
Dim sheetName As String
    
    
  Set ACT = CreateObject("system.collections.arraylist")
  Set PartNumber = CreateObject("system.collections.arraylist")
  Set QTY = CreateObject("system.collections.arraylist")
  Set location = CreateObject("system.collections.arraylist")
  Set DESCRIPTION = CreateObject("system.collections.arraylist")
  Set Note1 = CreateObject("system.collections.arraylist")
  Set fontValue = CreateObject("system.collections.arraylist")
 
 
  Set temp = CreateObject("system.collections.arraylist")
  
  
  
  
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
 
  Index = 6
  sheetName = ActiveSheet.name
  
    With Workbooks(tempWorkBookName).Worksheets(sheetName)
  
    Do While .Cells(Index, 1).value <> ""
    
    
        
        combi = .Cells(Index, 1).value & "/" & .Cells(Index, 2).value & "/" & .Cells(Index, 5).value & "/" & .Cells(Index, 6).value & "/" & .Cells(Index, 4).Font.Strikethrough
        
        If temp.Contains(combi) Then
        
            aryListIndex = temp.indexof(combi, 0)
        
            tempString = location(aryListIndex) & "," & .Cells(Index, 4).value
            tempNum = QTY(aryListIndex) + Abs(.Cells(Index, 3).value)
            location.removeat aryListIndex
            QTY.removeat aryListIndex
            
            location.Insert aryListIndex, tempString
            QTY.Insert aryListIndex, tempNum
            
            
            
        Else
            temp.add combi
            QTY.add Abs(.Cells(Index, 3).value)
            location.add .Cells(Index, 4).value
        End If
        
    
        Index = Index + 1
    
    Loop
    
    
    
    
    
    For i = 0 To temp.Count - 1
        
        tempAry = Split(temp(i), "/")
        
        ACT.add (tempAry(0))
        PartNumber.add (tempAry(1))
        DESCRIPTION.add (tempAry(2))
        Note1.add (tempAry(3))
        fontValue.add (tempAry(4))
    
    
    Next
    
    
    
    .Range(Cells(6, 1), Cells(Index, 6)).Delete
    
    Call printDataInSheet(6, sheetName, ACT, PartNumber, QTY, location, DESCRIPTION, Note1, fontValue)
    
    
  End With

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    
    ACT.Clear
    PartNumber.Clear
    QTY.Clear
    location.Clear
    DESCRIPTION.Clear
    Note1.Clear
    fontValue.Clear
    
    temp.Clear




    
    Call addBTN("D4", sheetName, "splitLocation", "Split Location")
End Sub


Private Sub printDataInSheet(startROW As Integer, sheetName As String, ACT As Variant, PartNumber As Variant, QTY As Variant, location As Variant, DESCRIPTION As Variant, Note1 As Variant, fontValue As Variant)
    
    Index = startROW
    
    
  
    With Workbooks(tempWorkBookName).Worksheets(sheetName)
    
        For i = 0 To ACT.Count - 1
        
            .Cells(Index + i, 1) = ACT(i)
            .Cells(Index + i, 2) = PartNumber(i)
            
            .Cells(Index + i, 4) = location(i)
            .Cells(Index + i, 5) = DESCRIPTION(i)
            .Cells(Index + i, 6) = Note1(i)
            
            If fontValue(i) Then
                .Cells(Index + i, 3) = QTY(i) * (-1)
                .Range(Cells(Index + i, 1), Cells(Index + i, 6)).Font.color = RGB(255, 0, 0)
                .Range(Cells(Index + i, 4), Cells(Index + i, 4)).Font.Strikethrough = fontValue(i)
            Else
            
                    .Cells(Index + i, 3) = QTY(i)

            End If
        
        Next i
    
    End With
    
End Sub

Sub enableManuChang(enableValue As Boolean)



    With Workbooks(tempWorkBookName)
    
    
    originNewBOM = .Worksheets("Main").Range("G40").value
    modifySheet = originNewBOM & "_Modify"
    manualSheet = .Worksheets("Main").Range("I40").value
    main = "Main"
    
    
    
        If enableValue Then
             temp = functionModule.creatSheet(manualSheet, 0, 255)
             
             .Worksheets(manualSheet).Range("A1").value = "No."
             .Worksheets(manualSheet).Range("B1").value = "ACT"
             .Worksheets(manualSheet).Range("C1").value = "ORIGINAL PART_NUMBER"
             .Worksheets(manualSheet).Range("D1").value = "NEW PART_NUMBER"
             .Worksheets(manualSheet).Range("E1").value = "Ref Des(If empty change all of this item)"
             .Worksheets(manualSheet).Range("F1").value = "QTY"
             .Worksheets(manualSheet).Columns("A:F").AutoFit
             
           
             
        Else
            
             temp = functionModule.delSheet(manualSheet, 0, 255)
             If functionModule.IsInWorksheet(modifySheet) Then
                Application.DisplayAlerts = False
                .Sheets(modifySheet).Delete
                Application.DisplayAlerts = True
             End If
             
             Call subModule.saveData(main, originNewBOM, "G41")
             .Worksheets(originNewBOM).Tab.ColorIndex = 41
        End If
        
    End With

End Sub


Sub manuChageList()

   
    
    Dim OriQPN, NewQPN, ACT, changeLoc, tempLoc, status, QTY
    'Dim DIP_SMD, COMPINENT_TYPE, DESCRIPTION
    Dim addQPN, addLoc
    
    Dim modifySheet As String
    Dim changeBomSheet As String
    Dim manuSheetName As String
    Dim cnt As Integer
    
    
    Set ACT = CreateObject("system.collections.arraylist")
    Set OriQPN = CreateObject("system.collections.arraylist")
    Set NewQPN = CreateObject("system.collections.arraylist")
    Set changeLoc = CreateObject("system.collections.arraylist")
    Set tempLoc = CreateObject("system.collections.arraylist")
    Set QTY = CreateObject("system.collections.arraylist")
    'Set status = CreateObject("system.collections.arraylist")
    
    Set addQPN = CreateObject("system.collections.arraylist")
    Set addLoc = CreateObject("system.collections.arraylist")
    
    'Set DIP_SMD = CreateObject("system.collections.arraylist")
    'Set COMPINENT_TYPE = CreateObject("system.collections.arraylist")
    'Set DESCRIPTION = CreateObject("system.collections.arraylist")
    
    'manuSheetName = "Manual Change"
    'changeBomSheet = "NewBOM"
    'modifySheet = "NewBom_Modify"
    
    
   
    
    
    cnt = UserForm1.TextBox20.value
    col_Partnumber = UserForm1.TextBox8.value
    col_QTY = UserForm1.TextBox9.value
    col_Location = UserForm1.TextBox10.value
    
    
    col_No = UserForm1.TextBox29.value
    col_DIP_SMD = UserForm1.TextBox25.value
    col_Com_TYP = UserForm1.TextBox27.value
    col_Descripiton = UserForm1.TextBox31.value
    
    'tempColAry = arrary(col_No, col_Partnumber, col_QTY, col_Location, col_Descripiton, col_DIP_SMD, col_Com_TYP)
    
    maxLow = 0


With Workbooks(tempWorkBookName)
    
    
    manuSheetName = .Worksheets("Main").Range("I40").value
    changeBomSheet = .Worksheets("Main").Range("G40").value
    modifySheet = changeBomSheet & "_Modify"
    
        
        With .Application
                .ScreenUpdating = False
                .DisplayAlerts = False
                .EnableEvents = False
        End With
    
          i = 2
    
        Do While Replace(.Worksheets(manuSheetName).Cells(i, 2).value, " ", "") <> ""
            
            'Insert active
            ACT.add (Replace(.Worksheets(manuSheetName).Cells(i, 2).value, " ", ""))
            .Worksheets(manuSheetName).Cells(i, 1).value = i - 1
            'Insert origin QPN /// only for change
     
            OriQPN.add (Replace(.Worksheets(manuSheetName).Cells(i, 3).value, " ", ""))

            'Insert newQPN
            NewQPN.add (Replace(.Worksheets(manuSheetName).Cells(i, 4).value, " ", ""))
            

            
            
            'Insert Location
            If Replace(.Worksheets(manuSheetName).Cells(i, 5).value, " ", "") = "" Then
                     changeLoc.add ("empty")
            Else
                    changeLoc.add (Replace(.Worksheets(manuSheetName).Cells(i, 5).value, " ", ""))
            End If
            
            
            'Insert QTY
            
            If Replace(.Worksheets(manuSheetName).Cells(i, 6).value, " ", "") = "" Then
                     QTY.add ("")
            Else
                    QTY.add (Replace(.Worksheets(manuSheetName).Cells(i, 6).value, " ", ""))
            End If
            
            
            'If Replace(.Worksheets(manuSheetName).Cells(i, 2).value, " ", "") = "A" Then
            '    If Replace(.Worksheets(manuSheetName).Cells(i, 5).value, " ", "") = "" And Replace(.Worksheets(manuSheetName).Cells(i, 6).value, " ", "") = "" Then
            '
            '        MsgBox "F" & i & " fill out  QTY if this item haven't location."
            '        Exit Sub
            '
            '    ElseIf Not IsNumeric(.Worksheets(manuSheetName).Cells(i, 6).value) Then
            '
            '        MsgBox "F" & i & " fill out a number."
            '        Exit Sub
            '    Else
            '
            '        If Int(.Worksheets(manuSheetName).Cells(i, 6).value) <> .Worksheets(manuSheetName).Cells(i, 6).value Then
            '            MsgBox "F" & i & " fill out  a Interger."
            '            Exit Sub
            '        Else
            '            QTY.add (Replace(.Worksheets(manuSheetName).Cells(i, 6).value, " ", ""))
            '        End If
            '
            '    End If
            'Else
            '
            '         QTY.add (Replace(.Worksheets(manuSheetName).Cells(i, 6).value, " ", ""))
            '
            'End If
            
               
                
            i = i + 1
            
        Loop
        
        .Worksheets(manuSheetName).Range("G1:G" & ACT.Count + 1).Clear
        .Worksheets(manuSheetName).Range("G1").value = "Result"
        
        .Worksheets(manuSheetName).Cells.ClearFormats
        .Worksheets(manuSheetName).Cells.Font.name = "Calibri"
        .Worksheets(manuSheetName).Cells.Font.Size = 12
        
        .Worksheets(manuSheetName).Range("A1:G" & i - 1).Borders.LineStyle = xlContinuous
        
        .Worksheets(manuSheetName).Range("A1:G1").Font.color = RGB(255, 255, 255)
        .Worksheets(manuSheetName).Range("A1:G1").Interior.color = RGB(0, 204, 255)
        
        'table check
        '
        '
        '
        
        
        
        'worksheet check need has NewBom sheet
        
        'check there is NewBow sheet or not
        '
        '
         If ACT.Count = 0 Then Exit Sub
         If Not changeTableCheck(manuSheetName, ACT, OriQPN, NewQPN, changeLoc, QTY, 2) Then GoTo exit_end:
        
        
        
        
        'functionModule.IsInWorksheet (changeBomSheet)
        
            If Not functionModule.IsInWorksheet(changeBomSheet) Then
               GoTo Exception1
            End If
            
            
            
            If functionModule.IsInWorksheet(modifySheet) Then
                Application.DisplayAlerts = False
                .Sheets(modifySheet).Delete
                Application.DisplayAlerts = True
            End If
             
             
             
             .Worksheets(changeBomSheet).Copy after:=.Worksheets(changeBomSheet)
             ActiveSheet.name = modifySheet
             .Worksheets(modifySheet).Cells.ClearFormats
             .Worksheets(modifySheet).Cells.Font.name = "Calibri"
             .Worksheets(modifySheet).Tab.ColorIndex = 41
             .Worksheets(changeBomSheet).Tab.ColorIndex = 47
             
             
             'Get table column
             
             tempDataCol = 0
             Do While Replace(.Worksheets(modifySheet).Cells(cnt, tempDataCol + 1).value, " ", "") <> ""
                tempDataCol = tempDataCol + 1
             Loop
             
             
             'Get data row
             'tempDataRow = .Worksheets(modifySheet).Range("B" & cnt).End(xlDown).Row
             
             
             tempDataRow = cnt
             Do While Replace(.Worksheets(modifySheet).Cells(tempDataRow, 2).value, " ", "") <> ""
                tempDataRow = tempDataRow + 1
             Loop
             
             
             
             
             If tempDataCol = 0 Or tempDataRow = 0 Then
             Exit Sub
             End If
             
             
             tempDataRow = tempDataRow - 1
             
             
             
             
             
             
            
             
             
        'Case Infomation
        '1. C: "C"hange item
        '2. A: "A"dd item
        '3. D: "D"elet item
        For i = 0 To ACT.Count - 1
        


            'If i = 28 Then Stop
        
            Select Case UCase(ACT(i))
    
            Case "C"
                
            Call changeItem(modifySheet, manuSheetName, OriQPN(i), NewQPN(i), changeLoc(i), col_Partnumber, col_QTY, col_Location, cnt, i, tempDataRow _
              , col_DIP_SMD, col_Com_TYP, col_Descripiton)
                 
            Case "A"
                Call addItem(modifySheet, manuSheetName, NewQPN(i), changeLoc(i), QTY(i), col_Partnumber, col_QTY, col_Location, cnt, i, tempDataRow _
              , col_DIP_SMD, col_Com_TYP, col_Descripiton)
              
            Case "D"
                Call delItem(modifySheet, manuSheetName, OriQPN(i), changeLoc(i), QTY(i), col_Partnumber, col_QTY, col_Location, cnt, i, tempDataRow)
            Case Else
                
            End Select
        
            
            '.Worksheets(manuSheetName).Cells(i + 2, 1).value = i + 1
        
        Next
        
        '----------------------------------------------
        ' Set form
        
        'Get Column
         On Error GoTo Exception2
         
        tempColValue = 0
        
        Do While Replace(.Worksheets(modifySheet).Cells(cnt - 1, tempColValue + 1).value, " ", "") <> ""
            tempColValue = tempColValue + 1
            
            
            If tempColValue > 1000 Then Exit Do
        Loop
        
        If tempColValue <> 0 Then
            MaxCol = Chr$(64 + tempColValue)
            
        .Worksheets(modifySheet).Cells.Font.name = "Calibri"
        .Worksheets(modifySheet).Cells.HorizontalAlignment = xlLeft
        
        .Worksheets(modifySheet).Range("A" & cnt - 1 & ":" & MaxCol & tempDataRow).Borders.LineStyle = xlContinuous
        
        
        .Worksheets(modifySheet).Range("A" & cnt - 1 & ":" & MaxCol & cnt - 1).Font.color = RGB(255, 255, 255)
        .Worksheets(modifySheet).Range("A" & cnt - 1 & ":" & MaxCol & cnt - 1).Interior.color = RGB(0, 204, 255)
            
            If UCase(.Worksheets(modifySheet).Cells(cnt - 1, col_QTY).value) = "QTY" Then
                    sumQTY = 0
                
                For i = 0 To (tempDataRow - cnt)
                    .Worksheets(modifySheet).Cells(cnt + i, col_No).value = i
                    sumQTY = sumQTY + .Worksheets(modifySheet).Cells(cnt + i, col_QTY).value
                Next
                
                .Worksheets(modifySheet).Cells(tempDataRow + 1, col_QTY).value = sumQTY
                .Worksheets(modifySheet).Cells(tempDataRow + 1, col_No).value = "TOTAL"
            End If
            
            
            
        Else
        
        End If
        
        
        
        
       
        Call subModule.saveData("Main", modifySheet, "G41")
        
        
        
        
        
GoTo exit_end
        
        
        
        
    
    





    
Exception1:

    MsgBox "Please load file first."
    GoTo exit_end
    
    
Exception2:

    MsgBox "Error Column."
    GoTo exit_end
    
    
    
    
exit_end:
       
        

        With .Application
                .ScreenUpdating = True
                .DisplayAlerts = True
                .EnableEvents = True
        End With

    
End With


     ACT.Clear
     OriQPN.Clear
     NewQPN.Clear
     changeLoc.Clear
     tempLoc.Clear
     QTY.Clear
     addQPN.Clear
     addLoc.Clear




End Sub


Private Sub changeItem(modifySheet As String, manuSheetName As String, _
                        OriQPN As Variant, NewQPN As Variant, changeLoc As Variant, _
                        col_Partnumber As Variant, col_QTY As Variant, col_Location As Variant, cnt As Integer, _
                        itemIndex As Variant, ByRef tempDataRow As Variant, _
                        col_DIP_SMD As Variant, col_Com_TYP As Variant, col_Descripiton As Variant)
                        

                Dim bolFindIndex As Boolean
                Dim fullChangeCount As Boolean

                tempStatus = ""
                fullChangeCount = 0
                SameString = ""
                tempcnt = cnt
                
                
                bolFindIndex = False
                fullChangeCount = False
                
                tempstring_1 = changeLoc
                
                
           With Workbooks(tempWorkBookName)
                
                Do While Replace(.Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).value, " ", "") <> ""
            
                
            
                    If OriQPN = .Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).value Then
                      
                      If .Worksheets(modifySheet).Cells(tempcnt, col_Descripiton).value Like "* DIP *" Then
                        tempDesc = "* DIP *"
                      Else
                        tempDesc = "?"
                      End If
                      
                      tempDIP = .Worksheets(modifySheet).Cells(tempcnt, col_DIP_SMD).value
                      tempTYP = .Worksheets(modifySheet).Cells(tempcnt, col_Com_TYP).value
                      
                      
                      bolFindIndex = True
                      
                      
                      '////////////////////////////////////////////////////////////////////////
                      
                      If changeLoc = "empty" Then  'Change All Item to another one
                        .Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).value = NewQPN
                        .Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).Font.color = RGB(255, 255, 255)
                        .Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).Interior.color = RGB(255, 0, 0)
                        fullChangeCount = True
                      Else
                        
                        'Compare Location
                        tempString_2 = .Worksheets(modifySheet).Cells(tempcnt, col_Location).value
                        tempSameLoc = functionModule.fcompareString(tempstring_1, tempString_2, ",")
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        
                        
                        
                     
                        
                        
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                                If tempString_2 = "" And tempSameLoc <> "" Then 'All Location meet
                                    With .Worksheets(modifySheet).Cells(tempcnt, col_Location)
                                        .value = tempSameLoc
                                        .Font.color = RGB(255, 255, 255)
                                        .Interior.color = RGB(255, 0, 0)
                                    End With
                                    
                                    With .Worksheets(modifySheet).Cells(tempcnt, col_Partnumber)
                                        .value = NewQPN
                                        .Font.color = RGB(255, 255, 255)
                                        .Interior.color = RGB(255, 0, 0)
                                    End With
    
                                ElseIf tempString_2 <> "" And tempSameLoc <> "" Then 'Some meet
                                
                                    With .Worksheets(modifySheet).Cells(tempcnt, col_Location)
                                        .value = tempString_2
                                        .Font.color = RGB(255, 255, 255)
                                        .Interior.color = RGB(255, 0, 0)
                                    End With
                                    
                                    With .Worksheets(modifySheet).Cells(tempcnt, col_QTY)
                                        .value = UBound(Split(tempString_2, ",")) + 1
                                        .Font.color = RGB(255, 255, 255)
                                        .Interior.color = RGB(255, 0, 0)
                                    End With
                                    
                                    
                                    
                                    'Save Same Location
                                    If SameString = "" Then
                                        SameString = tempSameLoc
                                    Else
                                        SameString = SameString & "," & tempSameLoc
                                    End If
                                    
                                    
                                    
                                    
                                Else
        
                                End If
                        
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                            If tempstring_1 = "" Then Exit Do
                        
                        
                        
                        
                      End If
                      '////////////////////////////////////////////////////////////////////////
                    Else
                    
                        'tempStatus = "Can't find" & NewQPN(i)
                    End If
                
                    tempcnt = tempcnt + 1
                
                Loop
                
                If SameString <> "" Then
                            
                            'find index
                            Dim bolCount As Boolean
                            
                            tempDataRowIndex = cnt
                            bolCount = True
                            
                            'If NewQPN = "CH31006KB18" Then Stop
                            
                            For i = tempDataRowIndex To tempDataRow
                                
                                If NewQPN = .Worksheets(modifySheet).Cells(i, col_Partnumber).value Then
                                    tempDataRowIndex = i
                                    bolCount = False
                                    Exit For
                                End If
                            Next
                            
                            If bolCount Then
                                tempDataRow = tempDataRow + 1
                                tempDataRowIndex = tempDataRow
                                
                                
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_Partnumber)
                                .value = NewQPN
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_Location)
                                .value = SameString
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_QTY)
                                .value = UBound(Split(SameString, ",")) + 1
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                        
                            
                            
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_Descripiton)
                                .value = tempDesc
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                            
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_DIP_SMD)
                                .value = tempDIP
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                            
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_Com_TYP)
                                .value = tempTYP
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                                
                                
                            ElseIf Not bolCount Then
                                
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_Partnumber)
                               
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                
                                End With
                
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_Location)
                                .value = .value & "," & SameString
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                
                                With .Worksheets(modifySheet).Cells(tempDataRowIndex, col_QTY)
                                .value = .value + UBound(Split(SameString, ",")) + 1
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                                End With
                            Else
                                
                                
                            End If
                            
                            
                            
                        
                            
                     
                    'tempDataRow = tempDataRow + 1
                    
                End If
                
                   'tempDesc = .Worksheets(modifySheet).Cells(tempcnt, col_Descripiton).value
                      'tempDIP = .Worksheets(modifySheet).Cells(tempcnt, col_DIP_SMD).value
                      'tempTYP = .Worksheets(modifySheet).Cells(tempcnt, col_Com_TYP).value
                
                
                'status.add tempStatus
                
                'Status
            If bolFindIndex Then
                If fullChangeCount Then
                    .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = "Done"
                Else
                
                    If tempstring_1 = "" Then
                        .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = "Done"
                    Else
                        .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = "Can't find " & tempstring_1
                    End If
                
                End If
                
            Else
                .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = "Can't find this PartNumber"
            End If
            
            
            'If fullChangeCount > 0 Or tempstring_1 = "" Then
            '
            '    .Worksheets(manuSheetName).Cells(itemIndex + 2, 6).value = "Done"
            'ElseIf Not fullChangeCount Then
            '    .Worksheets(manuSheetName).Cells(itemIndex + 2, 6).value = "Can't find this PartNumber"
            'ElseIf tempstring_1 <> "" Then
            '    .Worksheets(manuSheetName).Cells(itemIndex + 2, 6).value = "Can't find " & tempstring_1
            'Else
            'End If
            
            
        End With
                
                
End Sub


Private Sub addItem(modifySheet As String, manuSheetName As String, _
                        NewQPN As Variant, changeLoc As Variant, QTY As Variant, _
                        col_Partnumber As Variant, col_QTY As Variant, col_Location As Variant, cnt As Integer, _
                        itemIndex As Variant, ByRef tempDataRow As Variant, _
                        col_DIP_SMD As Variant, col_Com_TYP As Variant, col_Descripiton As Variant)
                        
                        
                Dim bolFindIndex As Boolean
                Dim fullChangeCount As Boolean
                Dim bolError
                Dim bolAddTrig As Boolean

                bolFindIndex = False
                fullChangeCount = False
                bolAddTrig = False
                bolError = False
                
                errorIndex = ""
                tempcnt = cnt
                tempstring1 = changeLoc
                        
          With Workbooks(tempWorkBookName)
                        
                        
                Do While Replace(.Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).value, " ", "") <> ""
            
                
            
                    If NewQPN = .Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).value Then
                      
                            bolFindIndex = True
                      
                      
                            If changeLoc = "empty" Then
                            
                              
                              
                                  If Replace(.Worksheets(modifySheet).Cells(tempcnt, col_Location).value, " ", "") <> "" Then
                                      bolError = True
                                      errorIndex = "0x1000" 'Add item no location but BOM has.
                                      Exit Do
                                  Else
                                      
                                      
                                      With .Worksheets(modifySheet).Cells(tempcnt, col_QTY)
                                        .value = .value + QTY
                                        .Font.color = RGB(255, 255, 255)
                                        .Interior.color = RGB(255, 0, 0)
                                      End With
                                      
                                      
                                      
                                      
                                      fullChangeCount = True
                                  End If
                              
                              
                            
                            Else
                            
                                  
                                  tempString2 = .Worksheets(modifySheet).Cells(tempcnt, col_Location).value
                                  tempSameLoc = functionModule.fcompareString(tempstring1, tempString2, ",")
                                  
                                  
                                  If tempstring1 <> "" Then
                                      bolAddTrig = True
                                      
                                      
                                      With .Worksheets(modifySheet)
                                        .Cells(tempcnt, col_Location).value = .Cells(tempcnt, col_Location).value & "," & tempstring1
                                        .Cells(tempcnt, col_Location).Font.color = RGB(255, 255, 255)
                                        .Cells(tempcnt, col_Location).Interior.color = RGB(255, 0, 0)
                                        
                                        .Cells(tempcnt, col_QTY).value = UBound(Split(.Cells(tempcnt, col_Location).value, ",")) + 1
                                        .Cells(tempcnt, col_QTY).Font.color = RGB(255, 255, 255)
                                        .Cells(tempcnt, col_QTY).Interior.color = RGB(255, 0, 0)
                                        
                                        
                                      End With
                                      
                                      
                                      
                                  Else
                                      bolError = True
                                      errorIndex = "0x1001" 'There was those locatin .
                                  End If
                                  
                                  Exit Do
                              
                            
                            End If

                    End If
                      
                
                    tempcnt = tempcnt + 1
                    
                'timeout
                If tempcnt > 10000 Then
                    Exit Do
                End If
                
                
                
                Loop
                
                                
                If Not bolFindIndex Then
                
                        
                            With .Worksheets(modifySheet).Cells(tempDataRow + 1, col_Partnumber)
                                .value = NewQPN
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                            End With
                            
                            If changeLoc <> "empty" Then
                                With .Worksheets(modifySheet).Cells(tempDataRow + 1, col_Location)
                                    .value = changeLoc
                                    .Font.color = RGB(255, 255, 255)
                                    .Interior.color = RGB(255, 0, 0)
                                End With
                
                                With .Worksheets(modifySheet).Cells(tempDataRow + 1, col_QTY)
                                    .value = UBound(Split(changeLoc, ",")) + 1
                                    .Font.color = RGB(255, 255, 255)
                                    .Interior.color = RGB(255, 0, 0)
                                End With
                            Else
                                 With .Worksheets(modifySheet).Cells(tempDataRow + 1, col_QTY)
                                    .value = QTY
                                    .Font.color = RGB(255, 255, 255)
                                    .Interior.color = RGB(255, 0, 0)
                                End With
                            End If
                            
                            
                            With .Worksheets(modifySheet).Cells(tempDataRow + 1, col_Descripiton)
                                .value = "?"
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                            End With
                            
                            With .Worksheets(modifySheet).Cells(tempDataRow + 1, col_DIP_SMD)
                                .value = "?"
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                            End With
                            
                            With .Worksheets(modifySheet).Cells(tempDataRow + 1, col_Com_TYP)
                                .value = "?"
                                .Font.color = RGB(255, 255, 255)
                                .Interior.color = RGB(255, 0, 0)
                            End With
                     
                    tempDataRow = tempDataRow + 1
                    
                End If
                
                
                'Work Status
                
                
                If bolError Then
                
                    .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = functionModule.ErrorInfo(errorIndex)
                Else
                
                    .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = functionModule.ErrorInfo("0x0000")
                
                End If
                
                    
                
        End With
                
    
End Sub

Sub delItem(modifySheet As String, manuSheetName As String, _
            OriQPN As Variant, changeLoc As Variant, QTY As Variant, _
            col_Partnumber As Variant, col_QTY As Variant, col_Location As Variant, cnt As Integer, _
            itemIndex As Variant, ByRef tempDataRow As Variant)
            
            
            
            tempcnt = cnt
            tempstring1 = changeLoc
            
             With Workbooks(tempWorkBookName)
                        
                        
                Do While Replace(.Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).value, " ", "") <> ""
                    
                    If OriQPN = .Worksheets(modifySheet).Cells(tempcnt, col_Partnumber).value Then
                      
                            bolFindIndex = True
                            
                            '--------------------------------------------------------------------------
                            
                            If changeLoc = "empty" Then
                            
                              fullChangeCount = True
  
                              .Worksheets(modifySheet).Rows(tempcnt).Delete Shift:=xlUp
                              tempDataRow = tempDataRow - 1
                              tempcnt = tempcnt - 1
                            
                            Else
                            
                                  
                                 
                                  
                                          tempString2 = .Worksheets(modifySheet).Cells(tempcnt, col_Location).value
                                          tempSameLoc = functionModule.fcompareString(tempstring1, tempString2, ",")
                                          
                                          
                                          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                
                                                If tempString2 = "" And tempSameLoc <> "" Then 'All Location meet
                                                   
                                                   
                                                    .Worksheets(modifySheet).Rows(tempcnt).Delete Shift:=xlUp
                                                    tempDataRow = tempDataRow - 1
                                                    tempcnt = tempcnt - 1
                                                    
                          
                                                ElseIf tempSameLoc <> "" Then
                                                
                                                    With .Worksheets(modifySheet).Cells(tempcnt, col_Location)
                                                        .value = tempString2
                                                        .Font.color = RGB(255, 255, 255)
                                                        .Interior.color = RGB(255, 0, 0)
                                                    End With
                                                    
                                                    With .Worksheets(modifySheet).Cells(tempcnt, col_QTY)
                                                        .value = UBound(Split(tempString2, ",")) + 1
                                                        .Font.color = RGB(255, 255, 255)
                                                        .Interior.color = RGB(255, 0, 0)
                                                    End With
                                                Else
                                                    'Nothing
                                                End If
                                        
                                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                        
                                                If tempstring1 = "" Then Exit Do
                                        
                                        
                                       
                              
                            
                            End If
                            
                            '--------------------------------------------------------------------------
                            
                    End If
                    
                    
                    
                    tempcnt = tempcnt + 1
                    'timeout
                        If tempcnt > 10000 Then Exit Do
                    
                    
                    
                Loop
                
                'Status
                If bolFindIndex Then
                    If fullChangeCount Then
                        .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = functionModule.ErrorInfo("0x0000")
                    Else
                         If tempstring1 = "" Then
                            .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = functionModule.ErrorInfo("0x0000")
                         Else
                            .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = "Can't find " & tempstring1
                         End If
                    
                    End If
                Else
                    .Worksheets(manuSheetName).Cells(itemIndex + 2, 7).value = functionModule.ErrorInfo("0x0001")
                End If
                
                
                
                
            End With
            
    
End Sub

Sub manualTableCheck(sheetName As Variant, col_act As Variant, col_OriQPN As Variant, col_NewQPN As Variant)

    

End Sub


