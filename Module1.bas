Attribute VB_Name = "Module1"
Public bomSumArry(1)
Public tempWorkBookName
Public tempWorkBookNamePath

Sub GetBookName(name As String)
    tempWorkBookName = name
End Sub

Sub start()

    
    tempWorkBookName = ActiveWorkbook.name
    tempWorkBookNamePath = ActiveWorkbook.Path
    
    
    
    'New feature
    
    
    'If functionModule.findFile(Replace(tempWorkBookName, ".xlsm", ""), tempWorkBookNamePath) Then
    '    UserForm1.Show vbModeless
    'Else
    '    UserForm2.Show vbModeless
    'End If
    
    
    
    UserForm1.Show vbModeless
    UserForm1.MultiPage1.value = 0
    
   
    'Stop
    
    functionModule.updateUserFormValue 0

    
    
End Sub

Sub load()

    Dim temprotectSheet
    Dim newBom, oldBom
    Dim bomArry, fileNameAry, bomSheetSel ', bomSumArry(1)
   
     Set temprotectSheet = CreateObject("system.collections.arraylist")
    
    
    
    
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With

With Workbooks(tempWorkBookName)
   
        
    '------------------------INIT------------------------
    
    newFileName = UserForm1.TextBox1.value
    oldFileName = UserForm1.TextBox2.value
    
    
    
    If Dir(newFileName, vbDirectory) = vbNullString Then
        MsgBox "New File Path Not Exist: " & vbCrLf & newFileName
        Exit Sub
        
    End If
    
    If Dir(oldFileName, vbDirectory) = vbNullString Then
        MsgBox "Old File Path Not Exist: " & vbCrLf & oldFileName
        Exit Sub
        
    End If
    
    
    
    fileNameAry = Array(newFileName, oldFileName)


    'modifyManu = "Manual Change"
    'newBom = "NewBom"
    'oldBom = "OldBom"
    modifyManu = .Worksheets("Main").Range("I40").value
    newBom = .Worksheets("Main").Range("G40").value
    oldBom = .Worksheets("Main").Range("H40").value
    
    
    
    bomArry = Array(newBom, oldBom)
    
    newSheetSel = UserForm1.ComboBox1.Text
    oldSheetSel = UserForm1.ComboBox2.Text
    bomSheetSel = Array(newSheetSel, oldSheetSel)
    
    col_Partnumber = Array(UserForm1.TextBox8.value, UserForm1.TextBox16.value)
    col_QTY = Array(UserForm1.TextBox9.value, UserForm1.TextBox17.value)
    col_Location = Array(UserForm1.TextBox10.value, UserForm1.TextBox18.value)
    
    'Micro_book = Application.ActiveWorkbook.Name
    Micro_book = tempWorkBookName
    '------------------------INIT------------------------
    
    
    temprotectSheet.add "Main"
    
     'Clear unused sheet
    If UserForm1.CheckBox1.value Then temprotectSheet.add newBom
    
    If UserForm1.CheckBox2.value Then temprotectSheet.add modifyManu
    
    
        functionModule.unUsedSheet (temprotectSheet.toArray())
    
    
    
    
    
    
    
    
    For i = 0 To 1
        
        If i = 0 Then
            cnt = .Worksheets("Main").Range("A30").value
        ElseIf i = 1 Then
            cnt = .Worksheets("Main").Range("B30").value
        Else
        End If
    
        If fileNameAry(i) <> "" And (i Or Not UserForm1.CheckBox1.value) Then
                
                
                tempValue = functionModule.checkFileType(fileNameAry(i))
                
                Select Case tempValue
                Case 0 'open_ExcelFile(fileName As String, sheetSel As String, sheetName As String, activeBook As String)
                    Call subModule.open_ExcelFile(fileNameAry(i), bomArry(i), Micro_book, bomSheetSel(i))
                Case 1
                    Call subModule.open_ConceptFile(fileNameAry(i), bomArry(i), Micro_book)
                'Case 2
                    'Call subModule.open_OrcadFile
                Case Else
                
                End Select
                
                 
           
        
            
            
            
            
                
        End If
        
    
    Next


    tempBol = functionModule.checkSheetForEnBTN(fileNameAry)

   

    Call subModule.saveData("Main", UserForm1.TextBox1.value, "A34")
    Call subModule.saveData("Main", UserForm1.ComboBox1.Text, "C34")
    Call subModule.saveData("Main", UserForm1.TextBox2.value, "A35")
    Call subModule.saveData("Main", UserForm1.ComboBox1.Text, "C35")
    'Call subModule.saveData("Main", True, "L35") 'Load finish
    
    
    Call subModule.saveData("Main", bomArry(0), "G41")
    Call subModule.saveData("Main", bomArry(1), "H41")
    
    
End With
    
exit1:
    
    
    temprotectSheet.Clear
    
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub

' General feature

Sub load_Gene(makeSheetName As Variant, textBoxValue As Variant, comboBoxTxt As Variant)

    Dim protectSheet
    Dim newBom, oldBom
    Dim bomArry, fileNameAry, bomSheetSel ', bomSumArry(1)
   
    
    
    
    
    
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With


    
    'Micro_book = Application.ActiveWorkbook.Name
    Micro_book = tempWorkBookName
    
        Dummy11 = functionModule.delSheet(makeSheetName, False, 12)
    
        If textBoxValue <> "" Then
                
                
                
                
                Select Case functionModule.checkFileType(textBoxValue)
                Case 0 'open_ExcelFile(fileName As String, sheetSel As String, sheetName As String, activeBook As String)
                    Call subModule.open_ExcelFile_Gene(textBoxValue, makeSheetName, Micro_book, comboBoxTxt, 15)
                Case 1
                    'Call subModule.open_ConceptFile(textBoxValue, makeSheetName, Micro_book)
                'Case 2
                    'Call subModule.open_OrcadFile
                Case Else
                
                End Select
                
                 
           
        
            
            
            
            
                
        End If




   

    
    
exit1:
    
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub
Sub load_sum()


    Dim protectSheet
    Dim newBom, oldBom
    Dim bomArry, fileNameAry, bomSheetSel ', bomSumArry(1)
    Dim newPartNumber, newQTY, newLocation
    Dim oldPartNumber, oldQTY, oldLocation
    Dim tempPartNumber, tempQTY, tempLocation
    
    
    
    
    Set newPartNumber = CreateObject("system.collections.arraylist")
    Set newQTY = CreateObject("system.collections.arraylist")
    Set newLocation = CreateObject("system.collections.arraylist")
    
    Set oldPartNumber = CreateObject("system.collections.arraylist")
    Set oldQTY = CreateObject("system.collections.arraylist")
    Set oldLocation = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber = CreateObject("system.collections.arraylist")
    Set tempQTY = CreateObject("system.collections.arraylist")
    Set tempLocation = CreateObject("system.collections.arraylist")
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With



    With Workbooks(tempWorkBookName)
    
    
        '------------------------INIT------------------------
        newFileName = .Worksheets("Main").Range("A34").value
        oldFileName = .Worksheets("Main").Range("A35").value
        fileNameAry = Array(newFileName, oldFileName)
        
        
        
        If UserForm1.CheckBox2.value Then
            newBom = "NewBom_Modify"
        Else
            newBom = "NewBom"
        End If
        
        oldBom = "OldBom"
        bomArry = Array(newBom, oldBom)
        
        newSheetSel = UserForm1.ComboBox1.Text
        oldSheetSel = UserForm1.ComboBox2.Text
        bomSheetSel = Array(newSheetSel, oldSheetSel)
        
        col_Partnumber = Array(UserForm1.TextBox8.value, UserForm1.TextBox16.value)
        col_QTY = Array(UserForm1.TextBox9.value, UserForm1.TextBox17.value)
        col_Location = Array(UserForm1.TextBox10.value, UserForm1.TextBox18.value)
        
        'Micro_book = Application.ActiveWorkbook.Name
        Micro_book = tempWorkBookName
        '------------------------INIT------------------------
        For i = 0 To 1
            
            If i = 0 Then
                cnt = .Worksheets("Main").Range("A30").value
            ElseIf i = 1 Then
                cnt = .Worksheets("Main").Range("B30").value
            Else
            End If
        
            If fileNameAry(i) <> "" Then
                    
                     
                'functionModule.removeTemp (Worksheets(tempAry).Cells(Index, 1).value)
            
                Do While functionModule.removeTemp(.Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value) <> ""  'remove space
                
                    If Not tempPartNumber.Contains(functionModule.removeTemp(.Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value)) Then  'First QPN
                    
                        tempPartNumber.add (functionModule.removeTemp(.Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value))
                        tempQTY.add (.Worksheets(bomArry(i)).Cells(cnt, col_QTY(i)).value)
                        tempLocation.add (.Worksheets(bomArry(i)).Cells(cnt, col_Location(i)).value)
                                        
                    Else 'Duplicat QPN
                    
                        indexNum = tempPartNumber.indexof(functionModule.removeTemp(.Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value), 0)
                        
                        tmpValueQty = tempQTY(indexNum)
                        tmpValueLocation = tempLocation(indexNum)
                        
                        tmpValueQty = tmpValueQty + .Worksheets(bomArry(i)).Cells(cnt, col_QTY(i)).value
                        tmpValueLocation = tmpValueLocation & "," & .Worksheets(bomArry(i)).Cells(cnt, col_Location(i)).value
                        
                        tempQTY.removeat (indexNum)
                        tempLocation.removeat (indexNum)
                        
                        tempQTY.Insert indexNum, tmpValueQty
                        tempLocation.Insert indexNum, tmpValueLocation
                    
                    End If
                    cnt = cnt + 1
                Loop
                
                
                bomSumArry(i) = functionModule.addSheet(bomArry(i), 1, tempPartNumber, tempQTY, tempLocation)
                
            
                tempPartNumber.Clear
                tempQTY.Clear
                tempLocation.Clear
                    
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
Sub load_Old()  'for refer


    Dim protectSheet
    Dim newBom, oldBom
    Dim bomArry, fileNameAry, bomSheetSel ', bomSumArry(1)
    Dim newPartNumber, newQTY, newLocation
    Dim oldPartNumber, oldQTY, oldLocation
    Dim tempPartNumber, tempQTY, tempLocation
    
    
    
    
    Set newPartNumber = CreateObject("system.collections.arraylist")
    Set newQTY = CreateObject("system.collections.arraylist")
    Set newLocation = CreateObject("system.collections.arraylist")
    
    Set oldPartNumber = CreateObject("system.collections.arraylist")
    Set oldQTY = CreateObject("system.collections.arraylist")
    Set oldLocation = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber = CreateObject("system.collections.arraylist")
    Set tempQTY = CreateObject("system.collections.arraylist")
    Set tempLocation = CreateObject("system.collections.arraylist")
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With


    'Clear unused sheet
    protectSheet = Array("Main")
    functionModule.unUsedSheet (protectSheet)

    '------------------------INIT------------------------
    newFileName = UserForm1.TextBox1.value
    oldFileName = UserForm1.TextBox2.value
    fileNameAry = Array(newFileName, oldFileName)

    newBom = "NewBom"
    oldBom = "OldBom"
    bomArry = Array(newBom, oldBom)
    
    newSheetSel = UserForm1.ComboBox1.Text
    oldSheetSel = UserForm1.ComboBox2.Text
    bomSheetSel = Array(newSheetSel, oldSheetSel)
    
    col_Partnumber = Array(UserForm1.TextBox8.value, UserForm1.TextBox16.value)
    col_QTY = Array(UserForm1.TextBox9.value, UserForm1.TextBox17.value)
    col_Location = Array(UserForm1.TextBox10.value, UserForm1.TextBox18.value)
    
    Micro_book = Application.ActiveWorkbook.name
    '------------------------INIT------------------------
    For i = 0 To 1
        
        If i = 0 Then
            cnt = Worksheets("Main").Range("A30").value
        ElseIf i = 1 Then
            cnt = Worksheets("Main").Range("B30").value
        Else
        End If
    
        If fileNameAry(i) <> "" Then
                
                
                tempValue = functionModule.checkFileType(fileNameAry(i))
                
                Select Case tempValue
                Case 0 'open_ExcelFile(fileName As String, sheetSel As String, sheetName As String, activeBook As String)
                    Call subModule.open_ExcelFile(fileNameAry(i), bomArry(i), Micro_book, bomSheetSel(i))
                Case 1
                    Call subModule.open_ConceptFile(fileNameAry(i), bomArry(i), Micro_book)
                'Case 2
                    'Call subModule.open_OrcadFile
                Case Else
                
                End Select
                
                 
            'functionModule.removeTemp (Worksheets(tempAry).Cells(Index, 1).value)
        
            Do While functionModule.removeTemp(Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value) <> ""  'remove space
            
                If Not tempPartNumber.Contains(Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value) Then  'First QPN
                
                    tempPartNumber.add (functionModule.removeTemp(Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value))
                    tempQTY.add (Worksheets(bomArry(i)).Cells(cnt, col_QTY(i)).value)
                    tempLocation.add (Worksheets(bomArry(i)).Cells(cnt, col_Location(i)).value)
                                    
                Else 'Duplicat QPN
                
                    indexNum = tempPartNumber.indexof(Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value, 0)
                    
                    tmpValueQty = tempQTY(indexNum)
                    tmpValueLocation = tempLocation(indexNum)
                    
                    tmpValueQty = tmpValueQty + Worksheets(bomArry(i)).Cells(cnt, col_QTY(i)).value
                    tmpValueLocation = tmpValueLocation & "," & Worksheets(bomArry(i)).Cells(cnt, col_Location(i)).value
                    
                    tempQTY.removeat (indexNum)
                    tempLocation.removeat (indexNum)
                    
                    tempQTY.Insert indexNum, tmpValueQty
                    tempLocation.Insert indexNum, tmpValueLocation
                
                End If
                cnt = cnt + 1
            Loop
            
            
            bomSumArry(i) = functionModule.addSheet(bomArry(i), 1, tempPartNumber, tempQTY, tempLocation)
            
        
            tempPartNumber.Clear
            tempQTY.Clear
            tempLocation.Clear
                
        End If
        
    
    Next


    tempBol = functionModule.checkSheetForEnBTN(bomSumArry)



    Call subModule.saveData("Main", UserForm1.TextBox1.value, "A34")
    Call subModule.saveData("Main", UserForm1.TextBox2.value, "A35")
    
exit1:
    
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
            
End Sub

Sub simpleCompare(sheetAry As Variant)

    Dim getArry
    Dim newPartNumber, newQTY, newLocation
    Dim oldPartNumber, oldQTY, oldLocation
    Dim tempPartNumber, tempQTY, tempLocation
    Dim tempPartNumber1, tempQTY1, tempLocation1 'for calulate N-O and O-N
    Dim tempPartNumber2, tempQTY2, tempLocation2 'for calulate N-O and O-N
    Dim tempPartNumber_ND, tempQTY_ND, tempLocation_ND    'for save new delete Item
    Dim tempState, tempState1, tempState2
    Dim IndexRow, tempIndexRow As Integer
 
    
    
    Set tempPartNumber = CreateObject("system.collections.arraylist")
    Set tempQTY = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber_ND = CreateObject("system.collections.arraylist")
    Set tempQTY_ND = CreateObject("system.collections.arraylist")
   
    
    Set tempState = CreateObject("system.collections.arraylist")
    Set tempState2 = CreateObject("system.collections.arraylist")
    
    Index_InfoRow = 2


    With Workbooks(tempWorkBookName)

        comSheetName = "Compare_Simple"
        mainSheetName = "Main"
        
        
     
        
        newFi = sheetAry(0)
        oldFi = sheetAry(1)
        getNewFile = newFi & "_SUM"
        getOldFile = oldFi & "_SUM"
        getArry = Array(getNewFile, getOldFile)
        
        functionModule.creatSheet comSheetName, True, 45
        
            '---------------------------------------V9_3 add--------------------------------------------------
        
            .Worksheets(comSheetName).Cells(1, 1).value = "TimeStamp : " & Format(Now, "yyyy/mm/dd/hh:mm")
            
        
            .Worksheets(comSheetName).Cells(Index_InfoRow, 1).value = "No."
            .Worksheets(comSheetName).Cells(Index_InfoRow, 2).value = "Item"
            .Worksheets(comSheetName).Cells(Index_InfoRow, 3).value = "File"
            .Worksheets(comSheetName).Cells(Index_InfoRow, 4).value = "SheetName"
            '.Worksheets(comSheetName).Cells(Index_InfoRow, 5).value = "Path"
            
         Index_InfoRow = Index_InfoRow + 1
            .Worksheets(comSheetName).Cells(Index_InfoRow, 1).value = Index_InfoRow - 2
            .Worksheets(comSheetName).Cells(Index_InfoRow, 2).value = newFi
            
            tempAryFileName = Split(.Worksheets(mainSheetName).Range("A34").value, "\")
            .Worksheets(comSheetName).Cells(Index_InfoRow, 3).value = tempAryFileName(UBound(tempAryFileName))
            .Worksheets(comSheetName).Cells(Index_InfoRow, 4).value = .Worksheets(mainSheetName).Range("C34").value
            '.Worksheets(comSheetName).Cells(Index_InfoRow, 5).value = UserForm1.TextBox1
            
        
        Index_InfoRow = Index_InfoRow + 1
           
        If newFi Like "*_Modify" Then
            .Worksheets(comSheetName).Cells(Index_InfoRow, 1).value = Index_InfoRow - 2
            .Worksheets(comSheetName).Cells(Index_InfoRow, 2).value = "Modify Source"
            
            tempAryFileName = Split(.Worksheets(mainSheetName).Range("A36").value, "\")
            .Worksheets(comSheetName).Cells(Index_InfoRow, 3).value = tempAryFileName(UBound(tempAryFileName))
            .Worksheets(comSheetName).Cells(Index_InfoRow, 4).value = .Worksheets(mainSheetName).Range("C36").value
            '.Worksheets(comSheetName).Cells(Index_InfoRow, 5).value = UserForm1.TextBox23
        
            Index_InfoRow = Index_InfoRow + 1
        End If
            .Worksheets(comSheetName).Cells(Index_InfoRow, 1).value = Index_InfoRow - 2
            .Worksheets(comSheetName).Cells(Index_InfoRow, 2).value = oldFi
            
            tempAryFileName = Split(.Worksheets(mainSheetName).Range("A35").value, "\")
            .Worksheets(comSheetName).Cells(Index_InfoRow, 3).value = tempAryFileName(UBound(tempAryFileName))
            .Worksheets(comSheetName).Cells(Index_InfoRow, 4).value = .Worksheets(mainSheetName).Range("C35").value
            '.Worksheets(comSheetName).Cells(Index_InfoRow, 5).value = UserForm1.TextBox2

        
        
        .Worksheets(comSheetName).Range("A2:D" & Index_InfoRow).Borders.LineStyle = xlContinuous
        
        
        .Worksheets(comSheetName).Range("A2:D2").Font.color = RGB(0, 0, 0)
        .Worksheets(comSheetName).Range("A2:D2").Interior.color = RGB(255, 255, 0)
        
        '---------------------------------------V9_3 add--------------------------------------------------
        
        
        IndexRow = 6
        IndexRow = Index_InfoRow + 3
        tempIndexRow = IndexRow
        
        .Worksheets(comSheetName).Cells(IndexRow - 1, 1).value = "ACT"
        .Worksheets(comSheetName).Cells(IndexRow - 1, 2).value = "Part_Number"
        .Worksheets(comSheetName).Cells(IndexRow - 1, 3).value = "QTY"
        
        .Worksheets(comSheetName).Range("A" & IndexRow - 1 & ":C" & IndexRow - 1).Font.color = RGB(255, 255, 255) 'V9_3 add
        
        .Worksheets(comSheetName).Range("A" & IndexRow - 1 & ":C" & IndexRow - 1).Interior.color = RGB(0, 204, 255) 'V9_3 add
        
        
        'get NewBom Data
         'On Error GoTo exit1
         For Each tempAry In getArry
         
            Index = 2
            
            Do While .Worksheets(tempAry).Cells(Index, 1).value <> ""
                
                tempPartNumber.add (.Worksheets(tempAry).Cells(Index, 1).value)
                tempQTY.add (.Worksheets(tempAry).Cells(Index, 2).value)
                tempState.add ("-")
                Index = Index + 1
            Loop
            
            
            Select Case Application.Match(tempAry, getArry, 0) - 1
                
            Case 0
                Set newPartNumber = tempPartNumber.Clone
                Set newQTY = tempQTY.Clone
                'Set tempState1 = tempState.Clone
            Case 1
                Set oldPartNumber = tempPartNumber.Clone
                Set oldQTY = tempQTY.Clone
                Set tempState1 = tempState.Clone
            Case Else
            
            End Select
            
            tempPartNumber.Clear
            tempQTY.Clear
            
            
         
         Next
         
         
         RemoveItem = "DELETE"
         addItem = "Add"
         
         For i = 0 To 1
         
            If i Then 'OLD-NEW
                Set tempPartNumber1 = oldPartNumber.Clone
                Set tempQTY1 = oldQTY.Clone
                Set tempPartNumber2 = newPartNumber.Clone
                Set tempQTY2 = newQTY.Clone
                tempString = RemoveItem
                Negative = "-"
                
            Else    'NEW-OLD
                Set tempPartNumber1 = newPartNumber.Clone
                Set tempQTY1 = newQTY.Clone
                Set tempPartNumber2 = oldPartNumber.Clone
                Set tempQTY2 = oldQTY.Clone
                tempString = addItem
                Negative = ""
            End If
            
            
            
            For Each tempPN In tempPartNumber1
            
                If tempPartNumber2.Contains(tempPN) Then
                    
                    If Not tempString = RemoveItem Then
                    
                        tempQTY_cal = tempQTY1(tempPartNumber1.indexof(tempPN, 0)) - tempQTY2(tempPartNumber2.indexof(tempPN, 0))
                        
                        tempQTY2.removeat tempPartNumber2.indexof(tempPN, 0)
                        tempQTY2.Insert tempPartNumber2.indexof(tempPN, 0), tempQTY_cal
                        
                        Select Case tempQTY_cal
                            Case 0
                                
                            Case Is > 0
                                tempState1.removeat tempPartNumber2.indexof(tempPN, 0)
                                tempState1.Insert tempPartNumber2.indexof(tempPN, 0), "Change(+)"
                                
                            Case Is < 0
                                tempState1.removeat tempPartNumber2.indexof(tempPN, 0)
                                tempState1.Insert tempPartNumber2.indexof(tempPN, 0), "Change(-)"
                            Case Else
                            
                        End Select
                        
                    End If
                    
                Else
                    
                    
                    
                    tempPartNumber_ND.add (tempPN)
                    tempQTY_ND.add (Negative & tempQTY1(tempPartNumber1.indexof(tempPN, 0)))
                    tempState2.add (tempString)
                    
                End If
                
            
            Next
            
           
           
            
            
            If i = 0 Then
                tempIndexRow = functionModule.inserData(tempIndexRow, comSheetName, tempPartNumber2, tempQTY2, tempState1, 1, "-")
                
                'sort
                .Activate
                .Worksheets(comSheetName).Select
                'LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
               .Worksheets(comSheetName).Range("A" & IndexRow & ":C" & tempIndexRow - 1).Sort key1:=.Worksheets(comSheetName).Range("A" & IndexRow & ":A" & tempIndexRow - 1), order1:=xlAscending, Header:=xlNo
            End If
            
            
            
            
            tempPartNumber1.Clear
            tempQTY1.Clear
            tempPartNumber2.Clear
            tempQTY2.Clear
         
         Next
        
    
            tempIndexRow = functionModule.inserData(tempIndexRow, comSheetName, tempPartNumber_ND, tempQTY_ND, tempState2, 0, "")
        
         '---------------------------------------V9_3 add--------------------------------------------------
        
        .Worksheets(comSheetName).Range("A" & Index_InfoRow + 2 & ":C" & tempIndexRow - 1).Borders.LineStyle = xlContinuous
        
        .Worksheets(comSheetName).Columns("A:E").AutoFit
        .Worksheets(comSheetName).Columns("A").ColumnWidth = 9
        .Worksheets(comSheetName).Columns("B").ColumnWidth = 20
        .Worksheets(comSheetName).Columns("C").ColumnWidth = 6
        
        .Worksheets(comSheetName).Cells.Font.name = "Calibri"
        .Worksheets(comSheetName).Cells.HorizontalAlignment = xlLeft
        
         '---------------------------------------V9_3 add--------------------------------------------------
    
    End With
    

    
exit1:
    
    tempPartNumber.Clear
    tempQTY.Clear
    
    
    tempPartNumber_ND.Clear
    tempQTY_ND.Clear
    
    
    tempState.Clear
    tempState2.Clear
    
    
        'MsgBox "Done."
    
End Sub

Sub compare(sheetAry As Variant)
    
    Dim IndexRow, tempIndexRow As Integer
    
    Dim newPartNumber, newQTY, newLocation, newState
    Dim oldPartNumber, oldQTY, oldLocation, oldState
    Dim tempPartNumber, tempQTY, tempLocation, tempState
    Dim tempPartNumber1, tempQTY1, tempLocation1, tempState1
    Dim tempPartNumber2, tempQTY2, tempLocation2, tempState2
    Dim tempPartNumber3, tempQTY3, tempLocation3, tempState3
    
    Set newPartNumber = CreateObject("system.collections.arraylist")
    Set newQTY = CreateObject("system.collections.arraylist")
    Set newLocation = CreateObject("system.collections.arraylist")
    Set newState = CreateObject("system.collections.arraylist")
    
    Set oldPartNumber = CreateObject("system.collections.arraylist")
    Set oldQTY = CreateObject("system.collections.arraylist")
    Set oldLocation = CreateObject("system.collections.arraylist")
    Set oldState = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber = CreateObject("system.collections.arraylist")
    Set tempQTY = CreateObject("system.collections.arraylist")
    Set tempLocation = CreateObject("system.collections.arraylist")
    Set tempState = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber1 = CreateObject("system.collections.arraylist")
    Set tempQTY1 = CreateObject("system.collections.arraylist")
    Set tempLocation1 = CreateObject("system.collections.arraylist")
    Set tempState1 = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber2 = CreateObject("system.collections.arraylist")
    Set tempQTY2 = CreateObject("system.collections.arraylist")
    Set tempLocation2 = CreateObject("system.collections.arraylist")
    Set tempState2 = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber3 = CreateObject("system.collections.arraylist")
    Set tempQTY3 = CreateObject("system.collections.arraylist")
    Set tempLocation3 = CreateObject("system.collections.arraylist")
    Set tempState3 = CreateObject("system.collections.arraylist")
    
    comSheetName = "Compare"
    'newFi = "NewBom"
    'oldFi = "OldBom"
    'getNewFile = "NewBom_SUM"
    'getOldFile = "OldBom_SUM"
    
    newFi = sheetAry(0)
    oldFi = sheetAry(1)
    getNewFile = newFi & "_SUM"
    getOldFile = oldFi & "_SUM"
    
    getArry = Array(getNewFile, getOldFile)
    
    functionModule.creatSheet comSheetName, True, 45
    
    Worksheets(comSheetName).Cells(1, 1).value = newFi
    Worksheets(comSheetName).Cells(1, 2).value = UserForm1.TextBox1
    
    
    Worksheets(comSheetName).Cells(2, 1).value = oldFi
    Worksheets(comSheetName).Cells(2, 2).value = UserForm1.TextBox2
    
    
    IndexRow = 6
    tempIndexRow = IndexRow
    
    Worksheets(comSheetName).Cells(IndexRow - 1, 1).value = "ACT"
    Worksheets(comSheetName).Cells(IndexRow - 1, 2).value = "Part_Number"
    Worksheets(comSheetName).Cells(IndexRow - 1, 3).value = "QTY"
    Worksheets(comSheetName).Cells(IndexRow - 1, 4).value = "Location"
    
    ''''''''''''''''''''''''''''''''''''
    'Get NewBom and OldBom value to List
    ''''''''''''''''''''''''''''''''''''
    For Each tempAry In getArry
     
        Index = 2
        
        Do While Worksheets(tempAry).Cells(Index, 1).value <> ""
            
            tempPartNumber.add (Worksheets(tempAry).Cells(Index, 1).value)
            tempQTY.add (Worksheets(tempAry).Cells(Index, 2).value)
            tempLocation.add (Worksheets(tempAry).Cells(Index, 3).value)
            tempState.add ("-")
            Index = Index + 1
        Loop
        
        
        
        '0>>get new BOM
        '1>>get old BOM
        
        Select Case Application.Match(tempAry, getArry, 0) - 1
            
        Case 0
            Set newPartNumber = tempPartNumber.Clone
            Set newQTY = tempQTY.Clone
            Set newLocation = tempLocation.Clone
            Set newState = tempState.Clone
        Case 1
            Set oldPartNumber = tempPartNumber.Clone
            Set oldQTY = tempQTY.Clone
            Set oldLocation = tempLocation.Clone
            Set oldState = tempState.Clone
        Case Else
        
        End Select
        
        tempPartNumber.Clear
        tempQTY.Clear
        tempLocation.Clear
        tempState.Clear
        
        
     
     Next
    ''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''


    
    RemoveItem = "DELETE"
    addItem = "Add"
    
    For i = 0 To 1
    
        If i Then 'Delete
            Set tempPartNumber = oldPartNumber.Clone
        
            Set tempPartNumber1 = oldPartNumber.Clone
            Set tempQTY1 = oldQTY.Clone
            Set tempLocation1 = oldLocation.Clone
            Set tempState1 = oldState.Clone
            
            Set tempPartNumber2 = newPartNumber.Clone
            Set tempQTY2 = newQTY.Clone
            Set tempLocation2 = newLocation.Clone
            Set tempState2 = newState.Clone
            
            tempString = RemoveItem
            tempInt = -1
            
        Else    'Change(+) Change(-) Add
            Set tempPartNumber = newPartNumber.Clone
        
            Set tempPartNumber1 = newPartNumber.Clone
            Set tempQTY1 = newQTY.Clone
            Set tempLocation1 = newLocation.Clone
            Set tempState1 = newState.Clone
            
            Set tempPartNumber2 = oldPartNumber.Clone
            Set tempQTY2 = oldQTY.Clone
            Set tempLocation2 = oldLocation.Clone
            Set tempState2 = oldState.Clone
            
            tempString = addItem
            tempInt = 1
            
        End If
        
        
        
        For Each tempPN In tempPartNumber
        
            If tempPartNumber2.Contains(tempPN) And tempState1(tempPartNumber1.indexof(tempPN, 0)) = "-" Then
                
                
                
                
                    tempstring_1 = tempLocation1(tempPartNumber1.indexof(tempPN, 0))
                    tempString_2 = tempLocation2(tempPartNumber2.indexof(tempPN, 0))

                    Call subModule.compareString(tempstring_1, tempString_2, ",")
                    
                    
                    '
                    'temp1
                    '
                    tempIndex = tempPartNumber1.indexof(tempPN, 0)
    
                    If UBound(Split(tempstring_1, ",")) + 1 = 0 Then
                    
                        tempPartNumber1.removeat tempIndex
                        tempQTY1.removeat tempIndex
                        tempLocation1.removeat tempIndex
                        tempState1.removeat tempIndex
                                        
                    Else
                    
                        tempLocation1.removeat tempIndex
                        tempLocation1.Insert tempIndex, tempstring_1
                        tempQTY1.removeat tempIndex
                        tempQTY1.Insert tempIndex, UBound(Split(tempstring_1, ",")) + 1
                        tempState1.removeat tempIndex
                        tempState1.Insert tempIndex, "Change(+)"
                    
                    End If
                    
                    
                    
                    '
                    'temp2
                    '
                    tempIndex = tempPartNumber2.indexof(tempPN, 0)
                    
                    If UBound(Split(tempString_2, ",")) + 1 = 0 Then
                    
                        tempPartNumber2.removeat tempIndex
                        tempQTY2.removeat tempIndex
                        tempLocation2.removeat tempIndex
                        tempState2.removeat tempIndex
                                        
                    Else
                    
                        tempLocation2.removeat tempIndex
                        tempLocation2.Insert tempIndex, tempString_2
                        tempQTY2.removeat tempIndex
                        tempQTY2.Insert tempIndex, (UBound(Split(tempString_2, ",")) + 1) * (-1)
                        tempState2.removeat tempIndex
                        tempState2.Insert tempIndex, "Change(-)"
                    
                    End If
                
                
                
            Else
                    
                If tempState1(tempPartNumber1.indexof(tempPN, 0)) = "-" Then
                    tempIndex = tempPartNumber1.indexof(tempPN, 0)
                    
                    tempPartNumber3.add tempPartNumber1(tempIndex)
                    tempQTY3.add tempQTY1(tempIndex) * tempInt
                    tempLocation3.add tempLocation1(tempIndex)
                    tempState3.add tempString
                    
                    'If tempQTY3.Count = 21 Then Stop
                    
                    
                    
                    
                    tempPartNumber1.removeat tempIndex
                    tempQTY1.removeat tempIndex
                    tempLocation1.removeat tempIndex
                    tempState1.removeat tempIndex
                    
                End If
            
            
            End If
        
         DoEvents
        
        Next
        
        
        If i = 0 Then
           'Delete'Change(+) Change(-) Add
        
            Set newPartNumber = tempPartNumber1.Clone
            Set newQTY = tempQTY1.Clone
            Set newLocation = tempLocation1.Clone
            Set newState = tempState1.Clone
            
            Set oldPartNumber = tempPartNumber2.Clone
            Set oldQTY = tempQTY2.Clone
            Set oldLocation = tempLocation2.Clone
            Set oldState = tempState2.Clone
         Else
            Set oldPartNumber = tempPartNumber1.Clone
            Set oldQTY = tempQTY1.Clone
            Set oldLocation = tempLocation1.Clone
            Set oldState = tempState1.Clone
            
        End If
        
        
        tempPartNumber1.Clear
        tempQTY1.Clear
        tempLocation1.Clear
        tempState1.Clear
        tempPartNumber2.Clear
        tempQTY2.Clear
        tempLocation2.Clear
        tempState2.Clear
        
        
         DoEvents
         
    Next
    
    
    
    
    tempIndexRow = functionModule.inserData2(tempIndexRow, comSheetName, newPartNumber, newQTY, newState, newLocation, 0, "", 0)   'Insert change(+)
    tempIndexRow = functionModule.inserData2(tempIndexRow, comSheetName, oldPartNumber, oldQTY, oldState, oldLocation, 0, "", 1)   'Insert change(-)
    
    'SORT
    Worksheets(comSheetName).Select
            'LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    Range("A" & IndexRow & ":D" & tempIndexRow - 1).Sort key1:=Range("B" & IndexRow & ":B" & tempIndexRow - 1), order1:=xlAscending, Header:=xlNo
    
    tempIndexRow = functionModule.inserData2(tempIndexRow, comSheetName, tempPartNumber3, tempQTY3, tempState3, tempLocation3, 0, "", 1) 'Insert new or remove component

    
    .Worksheets(comSheetName).Columns("A:D").AutoFit
    .Worksheets(comSheetName).Columns("B").ColumnWidth = 20
    .Worksheets(comSheetName).Columns("D").ColumnWidth = 20
    
    
    
exit1:
        newPartNumber.Clear
        newQTY.Clear
        newLocation.Clear
        newState.Clear

        oldPartNumber.Clear
        oldQTY.Clear
        oldLocation.Clear
        oldState.Clear
        
        tempPartNumber.Clear
        tempQTY.Clear
        tempLocation.Clear
        tempState.Clear
    
        tempPartNumber1.Clear
        tempQTY1.Clear
        tempLocation1.Clear
        tempState1.Clear
        
        tempPartNumber2.Clear
        tempQTY2.Clear
        tempLocation2.Clear
        tempState2.Clear
        
        tempPartNumber3.Clear
        tempQTY3.Clear
        tempLocation3.Clear
        tempState3.Clear
    
    
    
    
End Sub
Sub compare_2(sheetAry As Variant)
    
    Dim IndexRow, tempIndexRow As Integer
    Dim comSheetName As String
    
    Dim newPartNumber, newQTY, newLocation, newState
    Dim oldPartNumber, oldQTY, oldLocation, oldState
    Dim tempPartNumber, tempQTY, tempLocation, tempState
    Dim tempPartNumber1, tempQTY1, tempLocation1, tempState1
    
    Dim tempPartNumber3, tempQTY3, tempLocation3, tempState3
    Dim tempChangeLocation, tempChangeQPN, tempChangeState, tempChangeQTY
    
    Set newPartNumber = CreateObject("system.collections.arraylist")
    Set newQTY = CreateObject("system.collections.arraylist")
    Set newLocation = CreateObject("system.collections.arraylist")
    Set newState = CreateObject("system.collections.arraylist")
    
    Set oldPartNumber = CreateObject("system.collections.arraylist")
    Set oldQTY = CreateObject("system.collections.arraylist")
    Set oldLocation = CreateObject("system.collections.arraylist")
    Set oldState = CreateObject("system.collections.arraylist")
    
    Set tempPartNumber = CreateObject("system.collections.arraylist")
    Set tempQTY = CreateObject("system.collections.arraylist")
    Set tempLocation = CreateObject("system.collections.arraylist")
    Set tempState = CreateObject("system.collections.arraylist")
    
    
    
    
    Set tempPartNumber3 = CreateObject("system.collections.arraylist")
    Set tempQTY3 = CreateObject("system.collections.arraylist")
    Set tempLocation3 = CreateObject("system.collections.arraylist")
    Set tempState3 = CreateObject("system.collections.arraylist")
    
    'tempChangeSave
    Set tempChangeQPN = CreateObject("system.collections.arraylist")
    Set tempChangeLocation = CreateObject("system.collections.arraylist")
    Set tempChangeState = CreateObject("system.collections.arraylist")
    Set tempChangeQTY = CreateObject("system.collections.arraylist")
    
    With Workbooks(tempWorkBookName)
    
        comSheetName = "Compare_Detail"
        'newFi = "NewBom"
        'oldFi = "OldBom"
        'getNewFile = "NewBom_SUM"
        'getOldFile = "OldBom_SUM"
        
        newFi = sheetAry(0)
        oldFi = sheetAry(1)
        getNewFile = newFi & "_SUM"
        getOldFile = oldFi & "_SUM"
        
        
        getArry = Array(getNewFile, getOldFile)
        
        functionModule.creatSheet comSheetName, True, 45
        
        .Worksheets(comSheetName).Cells(1, 1).value = newFi
        .Worksheets(comSheetName).Cells(1, 2).value = UserForm1.TextBox1
        
        
        .Worksheets(comSheetName).Cells(2, 1).value = oldFi
        .Worksheets(comSheetName).Cells(2, 2).value = UserForm1.TextBox2
        
        
        IndexRow = 6
        tempIndexRow = IndexRow
        
        .Worksheets(comSheetName).Cells(IndexRow - 1, 1).value = "ACT"
        .Worksheets(comSheetName).Cells(IndexRow - 1, 2).value = "Part_Number"
        .Worksheets(comSheetName).Cells(IndexRow - 1, 3).value = "QTY"
        .Worksheets(comSheetName).Cells(IndexRow - 1, 4).value = "Location"
        .Worksheets(comSheetName).Cells(IndexRow - 1, 5).value = "Description"
        .Worksheets(comSheetName).Cells(IndexRow - 1, 6).value = "Note1"
        
        '
        'Get NewBom and OldBom value to List
        '
        For Each tempAry In getArry
         
            Index = 2
            
            Do While .Worksheets(tempAry).Cells(Index, 1).value <> ""
                
                
                
                
                tempPartNumber.add (.Worksheets(tempAry).Cells(Index, 1).value)
                tempQTY.add (.Worksheets(tempAry).Cells(Index, 2).value)
                
                
                
                If Not .Worksheets(tempAry).Cells(Index, 3).value = "" Then
                    
                
                    
                    tempLocation.add (.Worksheets(tempAry).Cells(Index, 3).value)
                    
                Else  'add temp location if location is empty
                        
                    tempLocation.add creatTempLocaiton(.Worksheets(tempAry).Cells(Index, 2).value, .Worksheets(tempAry).Cells(Index, 1).value)
                End If
                
                tempState.add ("X")
                Index = Index + 1
                
                DoEvents
            Loop
            
            
            Select Case Application.Match(tempAry, getArry, 0) - 1
                
            Case 0
                Set newPartNumber = tempPartNumber.Clone
                Set newQTY = tempQTY.Clone
                Set newLocation = tempLocation.Clone
                Set newState = tempState.Clone
            Case 1
                Set oldPartNumber = tempPartNumber.Clone
                Set oldQTY = tempQTY.Clone
                Set oldLocation = tempLocation.Clone
                Set oldState = tempState.Clone
            Case Else
            
            End Select
            
            tempPartNumber.Clear
            tempQTY.Clear
            tempLocation.Clear
            tempState.Clear
            
            
         
         Next
        
        
        Set tempPartNumber = newPartNumber.Clone
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'Remove the same Item (Same QPN and same Locaiton)
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        
        For Each tempQPN In tempPartNumber
            
            'If tempQPN = "HCS6I019010" Then Stop
            
            
            If oldPartNumber.Contains(tempQPN) Then
            
                'compare two string
                
                        'test
                        

                        
                        index_1 = newPartNumber.indexof(tempQPN, 0)
                        index_2 = oldPartNumber.indexof(tempQPN, 0)
                        
                
                        tempstring_1 = newLocation(index_1)
                        tempString_2 = oldLocation(index_2)
                        
                        Call subModule.compareString(tempstring_1, tempString_2, ",")
                         
                         
                        'remove new BOM empty item
                        
                        If tempstring_1 = "" Then
                        
                            newPartNumber.removeat index_1
                            newQTY.removeat index_1
                            newLocation.removeat index_1
                            newState.removeat index_1
                        Else
                            newQTY.removeat index_1
                            newLocation.removeat index_1
                            newState.removeat index_1
                            
                            newQTY.Insert index_1, UBound(Split(tempstring_1, ",")) + 1
                            newLocation.Insert index_1, tempstring_1
                            newState.Insert index_1, "-"
                            
                        End If
                        
                        'remove old BOM empty item
                        
                          If tempString_2 = "" Then
                        
                            oldPartNumber.removeat index_2
                            oldQTY.removeat index_2
                            oldLocation.removeat index_2
                            oldState.removeat index_2
                        Else
                            oldQTY.removeat index_2
                            oldLocation.removeat index_2
                            oldState.removeat index_2
                            
                            oldQTY.Insert index_2, UBound(Split(tempString_2, ",")) + 1
                            oldLocation.Insert index_2, tempString_2
                            oldState.Insert index_2, "-"
                            
                        End If
                         
            Else
                
                'index_3 = newPartNumber.indexof(tempQPN, 0)
            
            
                'NewItemPartNumber.Add newPartNumber(index_3)
                'NewItemQTY.Add newQTY(index_3)
                'NewItemLocation.Add newLocation(tindex_3)
                
                'newPartNumber.removeat index_3
                'newQTY.removeat index_3
                'newLocation.removeat index_3
                'newState.removeat index_3
                
                
            End If
            
            DoEvents
        Next
        
        
    
        
        '''''''''''''''''''''''''''''''''''''''''''
        
        'Find change location
        
        
        Set tempPartNumber = newPartNumber.Clone
        Set tempPartNumber3 = oldPartNumber.Clone
        
        If tempPartNumber3.Count = 0 Then
        
            tempIndexRow = functionModule.inserData3(tempIndexRow, comSheetName, newPartNumber, newQTY, newState, newLocation, 0, "", 0, False)
        
        Else
           For Each tempQPN2 In tempPartNumber
           
           
                'If tempQPN2 = "HCS6I019010" Then Stop
               index_temp = newPartNumber.indexof(tempQPN2, 0)
               
               
               tempSaveLocation = ""
               index_lock = ""
               
               
               tempLocation = Split(newLocation(index_temp), ",")
               
               
               
               For Each tempL In tempLocation
               
                  
                   tempstring1 = tempL
                   
                   For Each tempQPN3 In tempPartNumber3
                   
                       index_temp3 = oldPartNumber.indexof(tempQPN3, 0)
                       
                       
                       
                       tempString2 = oldLocation(index_temp3)
                       
                       Call subModule.compareString(tempstring1, tempString2, ",")
                       
                       
                       ''''''''''''''''''''''''''''''''''''''''''''''
                      
                       If tempString2 = "" Then 'There is no Location in the old part number
                               
                               oldPartNumber.removeat index_temp3
                               oldQTY.removeat index_temp3
                               oldLocation.removeat index_temp3
                               oldState.removeat index_temp3
                           
                       Else
                               oldQTY.removeat index_temp3
                               oldLocation.removeat index_temp3
                               
                               oldLocation.Insert index_temp3, tempString2
                               oldQTY.Insert index_temp3, UBound(Split(tempString2, ",")) + 1
                       End If
                       
                       ''''''''''''''''''''''''''''''''''''''''''''''
                       
                       If tempstring1 = "" Then 'Have Find change
                           
                           If tempChangeQPN.Contains(tempQPN2 & "/" & tempQPN3) Then
                               
                               tempChangeLocString = _
                                   tempChangeLocation(tempChangeQPN.indexof(tempQPN2 & "/" & tempQPN3, 0)) & "," & tempL
                               
                               tempChangeLocation.removeat tempChangeQPN.indexof(tempQPN2 & "/" & tempQPN3, 0)
                               tempChangeQTY.removeat tempChangeQPN.indexof(tempQPN2 & "/" & tempQPN3, 0)
                               tempChangeLocation.Insert tempChangeQPN.indexof(tempQPN2 & "/" & tempQPN3, 0), tempChangeLocString
                               tempChangeQTY.Insert tempChangeQPN.indexof(tempQPN2 & "/" & tempQPN3, 0), UBound(Split(tempChangeLocString, ",")) + 1
                               
                           Else
                               tempChangeLocation.add tempL
                               tempChangeQTY.add UBound(Split(tempL, ",")) + 1
                               tempChangeQPN.add tempQPN2 & "/" & tempQPN3
                               tempChangeState.add newState(index_temp)
                           End If
                           
                               
                           
                           Exit For
                       End If
                       
                       ''''''''''''''''''''''''''''''''''''''''''''''''
                           
                       DoEvents
          
                   Next
                   
                   'save change in tempPN3
                   Set tempPartNumber3 = oldPartNumber.Clone
                   
                   
                   'temp String for save back can't find location
                   If tempstring1 <> "" Then
                       If tempSaveLocation = "" Then
                           tempSaveLocation = tempstring1
                       Else
                           tempSaveLocation = tempSaveLocation & "," & tempstring1
                       End If
              
                   End If
                   
               Next
               
               
               
               'save back can't find location
               
                   If tempSaveLocation = "" Then
                       newLocation.removeat index_temp
                       newPartNumber.removeat index_temp
                       newQTY.removeat index_temp
                       newState.removeat index_temp
                   Else
                       newLocation.removeat index_temp
                       newQTY.removeat index_temp
                       newLocation.Insert index_temp, tempSaveLocation
                       newQTY.Insert index_temp, UBound(Split(tempSaveLocation, ",")) + 1
                   End If
               
           
           Next
           
        '''''''''''''''''''''''''''''''''''''''''''
        
           tempIndexRow = functionModule.inserData3(tempIndexRow, comSheetName, newPartNumber, newQTY, newState, newLocation, 0, "", 0, False)  'Insert change(+)
           tempIndexRow = functionModule.inserData3(tempIndexRow, comSheetName, oldPartNumber, oldQTY, oldState, oldLocation, 0, "", 1, False)
           tempIndexRow = functionModule.inserData3(tempIndexRow, comSheetName, tempChangeQPN, tempChangeQTY, tempChangeState, tempChangeLocation, 0, "", 0, 1)
        End If
        
        
        
        'SORT
        '.Worksheets(comSheetName).Select
                'LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
        'Range("A" & IndexRow & ":F" & tempIndexRow - 1).Sort key1:=Range("B" & IndexRow & ":B" & tempIndexRow - 1), order1:=xlAscending, Header:=xlNo
        
        'SORT
        .Activate
        .Worksheets(comSheetName).Select
        .Worksheets(comSheetName).Range("A" & IndexRow & ":F" & tempIndexRow - 1).Sort key1:=.Worksheets(comSheetName).Range("B" & IndexRow & ":B" & tempIndexRow - 1), order1:=xlAscending, Header:=xlNo
        
        .Worksheets(comSheetName).Columns("A:F").AutoFit
        .Worksheets(comSheetName).Columns("B").ColumnWidth = 20
        .Worksheets(comSheetName).Columns("D").ColumnWidth = 20
 
        '.Worksheets(comSheetName).Select
    
    
    End With
    
    
    Call subModule.addBTN("D4", comSheetName, "splitLocation", "Split Location")
    
    
    
    
exit1:
        newPartNumber.Clear
        newQTY.Clear
        newLocation.Clear
        newState.Clear

        oldPartNumber.Clear
        oldQTY.Clear
        oldLocation.Clear
        oldState.Clear
        
        tempPartNumber3.Clear
        tempQTY3.Clear
        tempLocation3.Clear
        tempState3.Clear
        
        tempChangeQPN.Clear
        tempChangeQTY.Clear
        tempChangeLocation.Clear
        tempChangeState.Clear
        
         
   
        'MsgBox "Done."
    
    
    
    
    
    
    
    
End Sub
Sub printOut(filname As Variant)

    Dim tempListAry
    
     With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    If UserForm1.ListBox3.ListCount = 0 Then GoTo exit2
    
    
    With Workbooks(tempWorkBookName)
    'Set wb1 = ActiveWorkbook
    'Set wb1 = tempWorkBook
    Set Wb2 = Application.Workbooks.add(1)
    
    Set tempListAry = CreateObject("system.collections.arraylist")
    
    tary = UserForm1.ListBox3.List
    tempSheetName = Wb2.Worksheets(1).name
    
    
    For i = 0 To UserForm1.ListBox3.ListCount - 1
        
        tempListAry.add (UserForm1.ListBox3.List(i))
        Wb2.Worksheets.add(before:=Wb2.Worksheets(Wb2.Sheets.Count)).name = UserForm1.ListBox3.List(i)
        .Worksheets(UserForm1.ListBox3.List(i)).Cells.Copy
        Wb2.Worksheets(UserForm1.ListBox3.List(i)).Range("A1").PasteSpecial
        
    Next
    
    tempAry = tempListAry.toArray()
    'temp = "'NewBom_SUM', 'OldBom_SUM'"
    'tempAry = Array("NewBom_SUM", "OldBom_SUM")
    
    
    
    
    
    
    'Wb1.Worksheets(tempAry).Copy before:=Wb2.Sheets(1)
    Wb2.Sheets(Wb2.Sheets.Count).Delete
    'Wb2.Sheets(1).Name = "BOM"
    
    On Error GoTo exit3
    Wb2.SaveAs fileName:=UserForm1.Label17 & "\" & filname & ".xlsx", FileFormat:=51
    'Wb2.SaveAs fileName:=wb1.Path & "\" & filname & ".xlsx", FileFormat:=51
    Wb2.Close
 End With
 
 
MsgBox "Done!"
    

'EVENT no sheet select


exitEnd:


With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
End With

    Exit Sub




exit2:
MsgBox "Please select at least one sheet for output."
GoTo exitEnd

exit3:
MsgBox "Error: Please check the validaty of saving path"
Wb2.Close
Call subModule.getSavePath
GoTo exitEnd


End Sub


