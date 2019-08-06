Attribute VB_Name = "functionModule"

'---------------------------------------------------------------
'---This function is uesed for check value is in array or not---
'---------------------------------------------------------------
Function IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
  'IsInArray = UBound(filter(arr, stringToBeFound)) > -1
  Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
Function IsInWorksheet(name As Variant) As Boolean
    
    Dim tempArry
    
    Set tempArry = CreateObject("system.collections.arraylist")


         With Workbooks(tempWorkBookName)
            
            For Each tempSheet In .Worksheets
                tempArry.add tempSheet.name
            Next
            
            IsInWorksheet = IsInArray(name, tempArry.toArray)
            
            tempArry.Clear
            
         End With

End Function
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

Function getPath(sheetName As String, saveLocaPath As String, saveLocaSheet As String, upSelNum As Integer) As String



    Dim isExcelFile As Boolean
    Dim excelFileType
    
    excelFileType = Array("xls", "xlsx", "xlsm")
    
    'tempValue = getPath_Ex
    tempValue = getPath_ExRptBom
    tempAry = Split(tempValue, "\")
    
    
    
    With Workbooks(tempWorkBookName)
    
    
        If tempValue = "" Then
            getPath = .Worksheets(sheetName).Range(saveLocaPath).value
        Else
                
                Application.ScreenUpdating = False
                functionModule.unProtectSheet (sheetName)
                .Worksheets(sheetName).Range(saveLocaPath).value = tempValue
                
                
                
                
                'Get worksheets list
               
                temp = ""
                
                If checkFileType(tempValue) = 0 Then
                
                    tempPathAry = Split(tempValue, "\")
                    
                    If functionModule.check_excelFile_open(tempPathAry(UBound(tempPathAry))) Then
                        MsgBox "'" & tempPathAry(UBound(tempPathAry)) & "' had opened, please save and close it. "
                        Exit Function
                    End If
                    
                    
                    
                    
                    Workbooks.Open fileName:=tempValue
                    
                    
                    For Each tempSheet In Worksheets
                    'ComboBox1.AddItem tempSheet.Name
                        If tempSheet.name = Worksheets.Item(1).name Then
                            temp = tempSheet.name
                        Else
                            temp = temp & "," & tempSheet.name
                        End If
                    Next
                    
                     Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
                    
                End If
                
                
                
               
                'ComboBox1.Text = ComboBox1.List(0)
        
                'tempAry = Split(tempValue, "\")
               
                
                .Worksheets(sheetName).Range(saveLocaSheet).value = temp
                functionModule.protectSheet (sheetName)
                Application.ScreenUpdating = True
                getPath = tempValue
                updateUserFormValue (upSelNum)
            
        End If
    
    
 
    End With

    
    

End Function



Function getPath_Ex() As String
    'Open nets file

    Micro_book = Application.ActiveWorkbook.name

    Dim filter, caption, datafilename, cmpsheet As String

    filter = "excel file (*.xls;*.xlsx;*.xlsm), *.xls;*.xlsx;*.xlsm"
    caption = "Select a NET file"
    datafilename = Application.GetOpenFilename(filter, , caption)

    'If datafilename = fail Then Exit Sub  'Do nothing when dont select file
    
    
    If datafilename = False Then
        getPath_Ex = ""
    Else
        getPath_Ex = datafilename
    End If
    

End Function

Function getPath_Rpt() As String
    'Open nets file

    Micro_book = Application.ActiveWorkbook.name

    Dim filter, caption, datafilename, cmpsheet As String

    filter = "BOM file (*.rpt), *.rpt"
    caption = "Select a NET file"
    datafilename = Application.GetOpenFilename(filter, , caption)

    'If datafilename = fail Then Exit Sub  'Do nothing when dont select file
    
    
    If datafilename = False Then
        getPath_Rpt = ""
    Else
        getPath_Rpt = datafilename
        
    End If
    
    

End Function

'Select Bom file
'*.rpt>> Concepte
'*.xls;*.xlsx;*.xlsm >> Excel file
'*.Bom;>> Orcad file


Function getPath_ExRptBom() As String

        'Micro_book = Application.ActiveWorkbook.Name
        Micro_book = tempWorkBookName
    

    Dim filter, caption, datafilename, cmpsheet As String
    
    
    filter = "BOM file (*.rpt;*.xls;*.xlsx;*.xlsm), *.rpt;*.xls;*.xlsx;*.xlsm"
    caption = "Select a NET file"
    datafilename = Application.GetOpenFilename(filter, , caption)

    'If datafilename = fail Then Exit Sub  'Do nothing when dont select file
    
    
    If datafilename = False Then
        getPath_ExRptBom = ""
    Else
        getPath_ExRptBom = datafilename
        Call subModule.saveData("Main", False, "M30")
        updateUserFormValue 7
    End If
    
    
    
End Function
Function getPath_Gene(sheetName As String, saveLocaPath As String, saveLocaSheet As String, upSelNum As Integer, filter As Variant) As String


    Dim isExcelFile As Boolean
    Dim excelFileType
    
    
    
    'tempValue = getPath_Ex
    tempValue = getPath_Gene_f(filter)
    tempAry = Split(tempValue, "\")
    
    
    
    With Workbooks(tempWorkBookName)
    
    
        If tempValue = "" Then
            getPath_Gene = .Worksheets(sheetName).Range(saveLocaPath).value
        Else
                
                Application.ScreenUpdating = False
                functionModule.unProtectSheet (sheetName)
                .Worksheets(sheetName).Range(saveLocaPath).value = tempValue
                
                
                
                
                'Get worksheets list
               
                temp = ""
                
                If checkFileType(tempValue) = 0 Then
                
                    tempPathAry = Split(tempValue, "\")
                    
                    If functionModule.check_excelFile_open(tempPathAry(UBound(tempPathAry))) Then
                        MsgBox "'" & tempPathAry(UBound(tempPathAry)) & "' had opened, please save and close it. "
                        Exit Function
                    End If
                
                
                    Workbooks.Open fileName:=tempValue
                    
                    
                    For Each tempSheet In Worksheets
                    'ComboBox1.AddItem tempSheet.Name
                        If tempSheet.name = Worksheets.Item(1).name Then
                            temp = tempSheet.name
                        Else
                            temp = temp & "," & tempSheet.name
                        End If
                    Next
                    
                     Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
                    
                End If
                
                
                
               
                'ComboBox1.Text = ComboBox1.List(0)
        
                'tempAry = Split(tempValue, "\")
               
                
                .Worksheets(sheetName).Range(saveLocaSheet).value = temp
                functionModule.protectSheet (sheetName)
                Application.ScreenUpdating = True
                getPath_Gene = tempValue
                updateUserFormValue (upSelNum)
            
        End If
    
    
 
    End With

    
    

End Function
Function getPath_Gene_f(filter As Variant) As String

        'Micro_book = Application.ActiveWorkbook.Name
        Micro_book = tempWorkBookName
    

    Dim caption, datafilename, cmpsheet As String
    
    
    'filter = "BOM file (*.rpt;*.xls;*.xlsx;*.xlsm), *.rpt;*.xls;*.xlsx;*.xlsm"
    caption = "Select a file"
    datafilename = Application.GetOpenFilename(filter, , caption)

    'If datafilename = fail Then Exit Sub  'Do nothing when dont select file
    
    
    If datafilename = False Then
        getPath_Gene_f = ""
    Else
        getPath_Gene_f = datafilename
        'Call subModule.saveData("Main", False, "M30")
        'updateUserFormValue 7
    End If
    
    
    
End Function

Function unUsedSheet(porTectAry As Variant)
    'Protect sheets

sheetProtect = porTectAry


'REMOVE UNUSED SHEETS
 With Workbooks(tempWorkBookName)

        For Each Worksheet In .Worksheets
            aa = Worksheet.name
            If functionModule.IsInArray(Worksheet.name, sheetProtect) = False Then
                Application.DisplayAlerts = False
                .Sheets(aa).Delete
                Application.DisplayAlerts = True
            End If
        Next

 End With

End Function


Function addSheet(sheetName As Variant, colorOnOff As Boolean, partNum As Variant, QTY As Variant, location As Variant) As String

 With Workbooks(tempWorkBookName)
        
        newSheetName = sheetName & "_SUM"
        
        For Each tempSheet In .Worksheets
        
            If tempSheet.name = newSheetName Then
            
                .Worksheets(newSheetName).Cells.Clear
                GoTo exit1
            
            End If
    
        Next
        
        
        .Sheets.add(after:=.Worksheets(.Worksheets.Count)).name = newSheetName
        'ActiveSheet.Name = newSheetName
    
    
exit1:
    
        .Worksheets(newSheetName).Cells(1, 1).value = "Part Number"
        .Worksheets(newSheetName).Cells(1, 2).value = "Qty"
        .Worksheets(newSheetName).Cells(1, 3).value = "Location"
                    

        
        
        If colorOnOff Then
                .Worksheets(newSheetName).Tab.ColorIndex = 33
        End If

        
        For i = 0 To partNum.Count - 1
        
            .Worksheets(newSheetName).Cells(i + 2, 1).value = partNum(i)
            .Worksheets(newSheetName).Cells(i + 2, 2).value = QTY(i)
            .Worksheets(newSheetName).Cells(i + 2, 3).value = location(i)
            
        Next
                   
    
    
         addSheet = newSheetName
         

End With


End Function



Function creatSheet(sheetName As Variant, colorOnOff As Boolean, colorCode As Integer)
    
    'check sheet name
    On Error GoTo exit1
    
    With Workbooks(tempWorkBookName)
    
                For Each tempSheet In .Worksheets
                    
                    If tempSheet.name = sheetName Then
                    
                        .Worksheets(sheetName).Cells.Clear
                        .Worksheets(sheetName).Buttons.Delete
                        GoTo exit1
                    
                    End If
                
                Next
                
                
                    .Sheets.add(after:=Worksheets(.Worksheets.Count)).name = sheetName
                    'ActiveSheet.Name = sheetName
            
            
exit1:
                    
                    If colorOnOff Then
                            .Worksheets(sheetName).Tab.ColorIndex = colorCode
                    End If
        
    End With

End Function

Function removeTemp(dataString As String) As String
    
    
    removeTemp = Replace(dataString, " ", "")
    
    
End Function
'functionModule.inserData(IndexRow, comSheetName, tempPartNumber2, tempQTY2, 1, "-")
Function inserData(indexNum As Integer, sheetName As Variant, PartNumber As Variant, QTY As Variant, state As Variant, ignorEnable As Boolean, ignorIndex As String) As Integer

With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
End With

With Workbooks(tempWorkBookName)

    For Each tempPN In PartNumber
    
        If ignorEnable And state(PartNumber.indexof(tempPN, 0)) = ignorIndex Then
        
        Else
            .Worksheets(sheetName).Cells(indexNum, 1).value = state(PartNumber.indexof(tempPN, 0))
            .Worksheets(sheetName).Cells(indexNum, 2).value = PartNumber(PartNumber.indexof(tempPN, 0))
            .Worksheets(sheetName).Cells(indexNum, 3).value = QTY(PartNumber.indexof(tempPN, 0))
            indexNum = indexNum + 1
        End If
        
        
        
        
        
    Next
    
    inserData = indexNum
End With


With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
End With

End Function

'functionModule.inserData(IndexRow, comSheetName, tempPartNumber2, tempQTY2, 1, "-")
Function inserData2(indexNum As Integer, sheetName As Variant, PartNumber As Variant, QTY As Variant, state As Variant, location As Variant, ignorEnable As Boolean, ignorIndex As String, enDeLine As Boolean) As Integer



    For Each tempPN In PartNumber
    
        If ignorEnable And state(PartNumber.indexof(tempPN, 0)) = ignorIndex Then
        
        Else
            Worksheets(sheetName).Cells(indexNum, 1).value = state(PartNumber.indexof(tempPN, 0))
            Worksheets(sheetName).Cells(indexNum, 2).value = PartNumber(PartNumber.indexof(tempPN, 0))
            Worksheets(sheetName).Cells(indexNum, 3).value = QTY(PartNumber.indexof(tempPN, 0))
            Worksheets(sheetName).Cells(indexNum, 4).value = location(PartNumber.indexof(tempPN, 0))
            
            If enDeLine And QTY(PartNumber.indexof(tempPN, 0)) < 0 Then
                Worksheets(sheetName).Cells(indexNum, 1).Font.color = RGB(255, 0, 0)
                Worksheets(sheetName).Cells(indexNum, 2).Font.color = RGB(255, 0, 0)
                Worksheets(sheetName).Cells(indexNum, 3).Font.color = RGB(255, 0, 0)
                Worksheets(sheetName).Cells(indexNum, 4).Font.Strikethrough = True
                Worksheets(sheetName).Cells(indexNum, 4).Font.color = RGB(255, 0, 0)
            Else
                Worksheets(sheetName).Cells(indexNum, 4).Font.Strikethrough = False
            End If
            indexNum = indexNum + 1
        End If
        
        
        
        
        
    Next
    
    inserData2 = indexNum
    

End Function

Function inserData3(indexNum As Integer, sheetName As Variant, PartNumber As Variant, QTY As Variant, state As Variant, location As Variant, _
ignorEnable As Boolean, ignorIndex As String, enDeLine As Boolean, enChange As Boolean) As Integer


With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
End With

            
 With Workbooks(tempWorkBookName)
        
    If enChange Then
            tempstring1 = "Change(+)"
            tempString1_1 = "Change components from "
            tempString2 = "Add"
            tempString2_2 = "New item change from "
    Else
        If enDeLine Then
            tempstring1 = "Change(-)"
            tempString1_1 = "Remove Components"
            tempString2 = "Delete"
            tempString2_2 = "Delete Item"
        Else
            tempstring1 = "Change(+)"
            tempString1_1 = "Add Components"
            tempString2 = "Add"
            tempString2_2 = "Add new item"
        End If
        
    End If
    


    For Each tempPN In PartNumber
    
        temp3 = ""
    
        If ignorEnable And state(PartNumber.indexof(tempPN, 0)) = ignorIndex Then  'ignor special
        
        Else
            
            If Not location(PartNumber.indexof(tempPN, 0)) Like "tmp-_*" Then
                .Worksheets(sheetName).Cells(indexNum, 4).value = location(PartNumber.indexof(tempPN, 0))
            End If
            
            If enChange Then
                
                tempQPN = Split(PartNumber(PartNumber.indexof(tempPN, 0)), "/")
                .Worksheets(sheetName).Cells(indexNum, 2).value = tempQPN(0)
                temp3 = tempQPN(1)
                .Worksheets(sheetName).Cells(indexNum, 6).value = tempQPN(1)
            Else
                .Worksheets(sheetName).Cells(indexNum, 2).value = PartNumber(PartNumber.indexof(tempPN, 0))
            
            End If
            
            
            
            
            Select Case state(PartNumber.indexof(tempPN, 0))
                Case "-"
                    .Worksheets(sheetName).Cells(indexNum, 1).value = tempstring1
                    .Worksheets(sheetName).Cells(indexNum, 5).value = tempString1_1 & temp3
                    
                Case "X"
                    .Worksheets(sheetName).Cells(indexNum, 1).value = tempString2
                    .Worksheets(sheetName).Cells(indexNum, 5).value = tempString2_2 & temp3
                
                Case Else
         
            End Select
            
            
            
            
            If enDeLine Then
                
                
                .Worksheets(sheetName).Cells(indexNum, 3).value = QTY(PartNumber.indexof(tempPN, 0)) * (-1)
                .Worksheets(sheetName).Cells(indexNum, 1).Font.color = RGB(255, 0, 0)
                .Worksheets(sheetName).Cells(indexNum, 2).Font.color = RGB(255, 0, 0)
                .Worksheets(sheetName).Cells(indexNum, 3).Font.color = RGB(255, 0, 0)
                .Worksheets(sheetName).Cells(indexNum, 4).Font.Strikethrough = True
                .Worksheets(sheetName).Cells(indexNum, 4).Font.color = RGB(255, 0, 0)
                .Worksheets(sheetName).Cells(indexNum, 5).Font.color = RGB(255, 0, 0)
            Else
                .Worksheets(sheetName).Cells(indexNum, 3).value = QTY(PartNumber.indexof(tempPN, 0))
                .Worksheets(sheetName).Cells(indexNum, 4).Font.Strikethrough = False
            End If
            
            indexNum = indexNum + 1
        
        End If
        
        
        
        
        
    Next
    
    inserData3 = indexNum
End With


With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
End With

End Function


Function protectSheet(sheetName As Variant)
With Workbooks(tempWorkBookName)
    .Worksheets(sheetName).Protect "isaac0516"
End With
End Function


Function unProtectSheet(sheetName As Variant)
With Workbooks(tempWorkBookName)
    .Worksheets(sheetName).Unprotect "isaac0516"
End With
End Function

'-----------------------------------------------------------------'
'Update Item

'For Update UserForm Data
'0-> UpdateAll
'1->UpdateAll SETTING PAGE
'2->Update SETTING PAGE NEW
'3->Update SETTING PAGE OLD


Function updateUserFormValue(caseValue As Integer)

    
    Select Case caseValue
    
    Case 0
        updateUserFormValue_FiPath
        
        updateUserFormValue_SetNewValue
        updateUserFormValue_SetOldValue
        updateUserFormValue_SetChangeValue
        updateUserFormValue_SetNewValue_row
        updateUserFormValue_SetOldValue_row
        updateUserFormValue_ComoNewList
        updateUserFormValue_ComoOldList
        updateList ("Main")
        updateBTNState
        updateOutPutPath
        updateManualChangeCheckBox
        updateUserFormValue_ComoChangeList
        updateManualChangebtnEn

    Case 1
        updateUserFormValue_SetNewValue
        updateUserFormValue_SetOldValue
    Case 11
        updateUserFormValue_SetChangeValue
    Case 2
        updateUserFormValue_SetNewValue
    Case 3
        updateUserFormValue_SetOldValue
    Case 4
        updateUserFormValue_ComoNewList
    Case 41
        updateUserFormValue_ComoChangeList
    Case 42
        updateManualChangebtnEn
    Case 5
        updateUserFormValue_ComoOldList
    Case 6
        updateList ("Main")
    Case 7
         updateBTNState
    Case 8
        updateUserFormValue_SetNewValue_row
    Case 9
        updateUserFormValue_SetOldValue_row
    
    Case Else
         
    End Select
    
    
   

    
   
End Function

Private Function updateUserFormValue_FiPath()

With Workbooks(tempWorkBookName)
    UserForm1.TextBox1 = .Worksheets("Main").Range("A34").value
    UserForm1.TextBox2 = .Worksheets("Main").Range("A35").value
    UserForm1.TextBox23 = .Worksheets("Main").Range("A36").value
    'UserForm1.OptionButton1.value = .Worksheets("Main").Range("N30").value
    'UserForm1.OptionButton2.value = .Worksheets("Main").Range("N31").value
    UserForm1.CheckBox3.value = .Worksheets("Main").Range("N30").value
    UserForm1.CheckBox4.value = .Worksheets("Main").Range("N31").value
    'UserForm1.OptionButton3.value = .Worksheets("Main").Range("N32").value
End With
End Function


Private Function updateUserFormValue_SetNewValue()
     'SETTING PAGE NEW LOAD EXCEL FILE
With Workbooks(tempWorkBookName)
    UserForm1.TextBox5.value = .Worksheets("Main").Range("A31").value
    UserForm1.TextBox6.value = .Worksheets("Main").Range("A32").value
    UserForm1.TextBox7.value = .Worksheets("Main").Range("A33").value
         
    UserForm1.TextBox8.value = .Worksheets("Main").Range("A31").value
    UserForm1.TextBox9.value = .Worksheets("Main").Range("A32").value
    UserForm1.TextBox10.value = .Worksheets("Main").Range("A33").value
    
End With
End Function
Private Function updateUserFormValue_SetNewValue_row()
     'SETTING PAGE NEW LOAD EXCEL FILE
With Workbooks(tempWorkBookName)
     
    UserForm1.TextBox19.value = .Worksheets("Main").Range("A30").value
         
    UserForm1.TextBox20.value = .Worksheets("Main").Range("A30").value
End With
End Function
Private Function updateUserFormValue_SetOldValue()
     'SETTING PAGE OLD LOAD EXCEL FILE
     
With Workbooks(tempWorkBookName)
    UserForm1.TextBox13.value = .Worksheets("Main").Range("B31").value
    UserForm1.TextBox14.value = .Worksheets("Main").Range("B32").value
    UserForm1.TextBox15.value = .Worksheets("Main").Range("B33").value
         
    UserForm1.TextBox16.value = .Worksheets("Main").Range("B31").value
    UserForm1.TextBox17.value = .Worksheets("Main").Range("B32").value
    UserForm1.TextBox18.value = .Worksheets("Main").Range("B33").value
    
End With
End Function
Private Function updateUserFormValue_SetOldValue_row()
     'SETTING PAGE NEW LOAD EXCEL FILE
     
With Workbooks(tempWorkBookName)
    UserForm1.TextBox21.value = .Worksheets("Main").Range("B30").value
         
    UserForm1.TextBox22.value = .Worksheets("Main").Range("B30").value
End With
End Function
Private Function updateUserFormValue_SetChangeValue()
     
With Workbooks(tempWorkBookName)
    UserForm1.TextBox28.value = .Worksheets("Main").Range("C30").value
    UserForm1.TextBox24.value = .Worksheets("Main").Range("C31").value
    UserForm1.TextBox26.value = .Worksheets("Main").Range("C32").value
    UserForm1.TextBox30.value = .Worksheets("Main").Range("C33").value
    
         
    UserForm1.TextBox29.value = .Worksheets("Main").Range("C30").value
    UserForm1.TextBox25.value = .Worksheets("Main").Range("C31").value
    UserForm1.TextBox27.value = .Worksheets("Main").Range("C32").value
    UserForm1.TextBox31.value = .Worksheets("Main").Range("C33").value
    
End With
End Function
Private Function updateUserFormValue_ComoNewList()
     
 With Workbooks(tempWorkBookName)
     
     UserForm1.ComboBox1.Clear
     tempList = Split(.Worksheets("Main").Range("B34"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox1.addItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox1.Text = UserForm1.ComboBox1.List(0)
End With
exit1:
     
End Function
Private Function updateUserFormValue_ComoChangeList()
     
 With Workbooks(tempWorkBookName)
     
     UserForm1.ComboBox3.Clear
     tempList = Split(.Worksheets("Main").Range("B36"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox3.addItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox3.Text = UserForm1.ComboBox3.List(0)
End With
exit1:
     
End Function
Private Function updateUserFormValue_ComoOldList()
    UserForm1.ComboBox2.Clear
    
With Workbooks(tempWorkBookName)
     tempList = Split(.Worksheets("Main").Range("B35"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox2.addItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox2.Text = UserForm1.ComboBox2.List(0)
End With
exit1:
     
End Function
Private Function updateList(igNoreSheet As Variant)
    
    UserForm1.ListBox2.Clear
    UserForm1.ListBox3.Clear
    For Each tempsh In Worksheets
    
        If Not tempsh.name = igNoreSheet Then
            UserForm1.ListBox2.addItem (tempsh.name)
        End If
    Next
    
End Function
Private Function updateBTNState()
   
   'Compare BTN
   UserForm1.CommandButton13.Enabled = Workbooks(tempWorkBookName).Worksheets("Main").Range("M30").value
   
    
End Function

Function updateOutPutPath()

    UserForm1.Label17.caption = Workbooks(tempWorkBookName).Worksheets("Main").Range("L30").value

End Function
Function updateManualChangeCheckBox()

    UserForm1.CheckBox2 = Workbooks(tempWorkBookName).Worksheets("Main").Range("L34").value

End Function
Function updateManualChangebtnEn()
   
   If UserForm1.CheckBox2.value = True Then
        UserForm1.CommandButton20.Enabled = True
        UserForm1.TextBox23.Enabled = True
        UserForm1.CommandButton21.Enabled = True
        UserForm1.CommandButton22.Enabled = True
        UserForm1.ComboBox3.Enabled = True
    Else
        UserForm1.CommandButton20.Enabled = False
        UserForm1.TextBox23.Enabled = False
        UserForm1.CommandButton21.Enabled = False
        UserForm1.CommandButton22.Enabled = False
        UserForm1.ComboBox3.Enabled = False
    End If
   

End Function



'-----------------------------------------------------------------'







'0 -> EXCEL FILE (.xls ; .xlsx ; .xlsm)
'1 -> CONCEPT FILE (.rpt)
'2 -> Orcad FILE (.BOM)

Function checkFileType(checkFileName As Variant) As Integer

    fileNamExt = Split(checkFileName, ".")
    
    
    Select Case LCase(fileNamExt(UBound(fileNamExt)))
    
    Case "xls", "xlsx", "xlsm"
        checkFileType = 0
    Case "rpt"
        checkFileType = 1
    Case "bom"
        checkFileType = 2
    Case Else
        checkFileType = 100
    End Select

End Function

Function checkSheetForEnBTN(checkAry As Variant) As Boolean
    
    Dim checkEn As Boolean
    
    checkEn = True
    
    For Each tempAry In checkAry
    
        checkEn = checkEn And (tempAry <> "")
    
    Next
    
    
    
    
    If checkEn Then
        UserForm1.CommandButton13.Enabled = True
        Call saveData("Main", True, "M30")
    Else
         UserForm1.CommandButton13.Enabled = False
         Call saveData("Main", False, "M30")
    End If
    
    
End Function


Function creatTempLocaiton(QTY As Integer, PN As String) As String
    
    'tempPN = Left(PN, 7)
    tempPN = PN
    
    If QTY = 1 Then
        creatTempLocaiton = "tmp-_" & tempPN & "_1"
    Else
    
        tempString = "tmp-_" & PN & "_1"
        
        For i = 2 To QTY
            tempString = tempString & "," & "tmp-_" & tempPN & "_" & i
        Next
        
        creatTempLocaiton = tempString
        
    End If
    
    
    
    


End Function

Function delSheet(sheetName As Variant, colorOnOff As Boolean, colorCode As Integer)
    
    temp = 0
    
    'check sheet name
    On Error GoTo exit1
    
    With Workbooks(tempWorkBookName)
    
                For Each tempSheet In .Worksheets
                    
                    If tempSheet.name = sheetName Then
                    
                        .Worksheets(sheetName).Delete
                        GoTo exit1
                    
                    End If
                
                Next
                
                
                    
            
            
            temp = 1
            
exit1:
                    
                    If colorOnOff Then
                            .Worksheets(sheetName).Tab.ColorIndex = colorCode
                    End If
                    
                delSheet = temp
        
    End With

End Function

Function ErrorInfo(errIndex As Variant)

    Select Case errIndex
    
    Case "0x0000"
        tempString = "Done"
    Case "0x0001"
        tempString = "Can't find this PartNumber"
    Case "0x1000"
        tempString = "ERROR.Adding Item has no location but BOM have."
    Case "0x1001"
        tempString = "ERROR.BOM have this QPN and Location.This change is not necessary"
    Case Else
        
    End Select
    
    ErrorInfo = tempString

End Function

Function fcompareString(ByRef String1 As Variant, ByRef String2 As Variant, ignorType As Variant)
  
  Dim tempList1, tempList2, sameLoc
  
  
  Set sameLoc = CreateObject("system.collections.arraylist")
  
  tempAry = Split(String1, ignorType)
  Set tempList1 = aryToList(tempAry)
  
  tempAry = Split(String2, ignorType)
  Set tempList2 = aryToList(tempAry)
  
  Set tempList1Clone = tempList1.Clone
  
  
  
  For Each tempValue In tempList1
  
    If tempList2.Contains(tempValue) Then
    
        sameLoc.add tempValue
        tempList2.removeat (tempList2.indexof(tempValue, 0))
        tempList1Clone.removeat (tempList1Clone.indexof(tempValue, 0))
    End If
  
  Next

  String1 = aryToString(tempList1Clone)
  
  String2 = aryToString(tempList2)
  
  fcompareString = aryToString(sameLoc)
  
  
End Function

Function load_sum()


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
    
         'Remove Old Compare sheet
        '------------------------INIT------------------------
                    For Each tempSheet In .Worksheets
                        
                        If tempSheet.name Like "Compare_*" Or tempSheet.name Like "*_SUM" Then
                        
                            .Worksheets(tempSheet.name).Delete
                        
                        End If
                    
                    Next
           
        '------------------------INIT------------------------
        newFileName = .Worksheets("Main").Range("A34").value
        oldFileName = .Worksheets("Main").Range("A35").value
        fileNameAry = Array(newFileName, oldFileName)
        
        
        
        'If UserForm1.CheckBox2.value Then
        '    newBom = "NewBom_Modify"
        'Else
        '    newBom = "NewBom"
        'End If
        
        'oldBom = "OldBom"
        
        newBom = .Worksheets("Main").Range("G41").value
        oldBom = .Worksheets("Main").Range("H41").value
        
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
                
                    If Not tempPartNumber.Contains(.Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value) Then  'First QPN
                    
                        
                        tempPartNumber.add (functionModule.removeTemp(.Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value))
                        tempQTY.add (.Worksheets(bomArry(i)).Cells(cnt, col_QTY(i)).value)
                        tempLocation.add functionModule.removeTemp(.Worksheets(bomArry(i)).Cells(cnt, col_Location(i)).value)
                                        
                    Else 'Duplicat QPN
                    
                        indexNum = tempPartNumber.indexof(.Worksheets(bomArry(i)).Cells(cnt, col_Partnumber(i)).value, 0)
                        
                        tmpValueQty = tempQTY(indexNum)
                        tmpValueLocation = tempLocation(indexNum)
                        
                        tmpValueQty = tmpValueQty + .Worksheets(bomArry(i)).Cells(cnt, col_QTY(i)).value
                        tempQTY.removeat (indexNum)
                        tempQTY.Insert indexNum, tmpValueQty
                        
                        
                        If functionModule.removeTemp(.Worksheets(bomArry(i)).Cells(cnt, col_Location(i)).value) <> "" Then
                            tmpValueLocation = tmpValueLocation & "," & .Worksheets(bomArry(i)).Cells(cnt, col_Location(i)).value
                            tempLocation.removeat (indexNum)
                            tempLocation.Insert indexNum, tmpValueLocation
                        End If
                        
                        
                        
                        
                    
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
    
    load_sum = bomArry
            
End Function

Function findFile(fileName As Variant, filePath As Variant) As Boolean

    
    Dim i As Integer
    
    
    myFile = Dir(filePath & "\" & fileName & ".csv")
    
    
    If myFile <> "" Then
        findFile = True
    Else
        findFile = False
    End If
    


End Function

Function changeTableCheck(manuSheetName As String, _
                          ACT As Variant, OriQPN As Variant, NewQPN As Variant, location As Variant, QTY As Variant, _
                          startROW As Integer) As Boolean


    Dim errorCount As Integer
    Dim errorMessage
    
    sheetName = manuSheetName
    Set errorMessage = CreateObject("system.collections.arraylist")
    
    For i = 0 To ACT.Count - 1
    
    
        Select Case UCase(ACT(i))
        
            Case "C"
                If changeItemCheck(manuSheetName, OriQPN(i), NewQPN(i), location(i), QTY(i), startROW + i) <> "" Then
                    errorMessage.add "No." & i + 1 & " : " & changeItemCheck(manuSheetName, OriQPN(i), NewQPN(i), location(i), QTY(i), startROW + i)
                End If
                
            Case "D"
                 If delItemCheck(manuSheetName, OriQPN(i), NewQPN(i), location(i), QTY(i), startROW + i) <> "" Then
                    errorMessage.add "No." & i + 1 & " : " & delItemCheck(manuSheetName, OriQPN(i), NewQPN(i), location(i), QTY(i), startROW + i)
                 End If
            Case "A"
                 If addItemCheck(manuSheetName, OriQPN(i), NewQPN(i), location(i), QTY(i), startROW + i) <> "" Then
                    errorMessage.add "No." & i + 1 & " : " & addItemCheck(manuSheetName, OriQPN(i), NewQPN(i), location(i), QTY(i), startROW + i)
                 End If
            Case Else
            
        
        End Select
    
    Next
    
    
    
    If errorMessage.Count = 0 Then
        changeTableCheck = True
    Else
        changeTableCheck = False
        Call errorReport(errorMessage)
    End If
    
    
End Function

Private Function changeItemCheck(manuSheetName As String, OriQPN As Variant, NewQPN As Variant, location As Variant, QTY As Variant, indexCell As Variant) As String


    With Workbooks(tempWorkBookName)
    

        If Replace(OriQPN, " ", "") = "" Or Replace(NewQPN, " ", "") = "" Then
            If Replace(OriQPN, " ", "") = "" Then .Worksheets(manuSheetName).Cells(indexCell, 3).Interior.color = RGB(255, 0, 0)
            If Replace(NewQPN, " ", "") = "" Then .Worksheets(manuSheetName).Cells(indexCell, 4).Interior.color = RGB(255, 0, 0)
            changeItemCheck = "Please fill out original / new QPN."
        Else
            
        End If
        
        
    End With
    
End Function

Private Function addItemCheck(manuSheetName As String, OriQPN As Variant, NewQPN As Variant, location As Variant, QTY As Variant, indexCell As Variant) As String

     With Workbooks(tempWorkBookName)
        
        If Replace(OriQPN, " ", "") <> "" And Replace(NewQPN, " ", "") <> "" Then
             .Worksheets(manuSheetName).Cells(indexCell, 3).Interior.color = RGB(255, 0, 0)
             .Worksheets(manuSheetName).Cells(indexCell, 4).Interior.color = RGB(255, 0, 0)
            addItemCheck = "Please check which QPN you want to add."
        'ElseIf Replace(OriQPN, " ", "") <> "" And Replace(NewQPN, " ", "") = "" Then
        ElseIf Replace(NewQPN, " ", "") = "" Then
            .Worksheets(manuSheetName).Cells(indexCell, 4).Interior.color = RGB(255, 0, 0)
            addItemCheck = "Please fill out QPN in New Part Number."
        ElseIf Replace(OriQPN, " ", "") = "" And Replace(NewQPN, " ", "") <> "" Then
        
            If Replace(location, " ", "") = "empty" And Replace(QTY, " ", "") = "" Then
                    '.Worksheets(manuSheetName).Cells(indexCell, 5).Interior.Color = RGB(255, 0, 0)
                    .Worksheets(manuSheetName).Cells(indexCell, 6).Interior.color = RGB(255, 0, 0)
                    addItemCheck = "Fill out  QTY if this item haven't location."
            ElseIf Replace(location, " ", "") = "empty" And Replace(QTY, " ", "") <> "" Then
                
                
                If Not IsNumeric(QTY) Then
                    .Worksheets(manuSheetName).Cells(indexCell, 6).Interior.color = RGB(255, 0, 0)
                    addItemCheck = "Fill out a number."
                Else
            
                   If Int(CDbl(QTY)) <> CDbl(QTY) Then
                        .Worksheets(manuSheetName).Cells(indexCell, 6).Interior.color = RGB(255, 0, 0)
                        addItemCheck = "Fill out  a Interger."
                    End If
            
                End If
            ElseIf Replace(location, " ", "") <> "empty" Then
                'Do nothing
            Else
                MsgBox "Exception! _functionModule_addItemCheck_2"
            End If
            
        Else
            MsgBox "Exception! _functionModule_addItemCheck_1"
        End If
     End With
End Function
Private Function delItemCheck(manuSheetName As String, OriQPN As Variant, NewQPN As Variant, location As Variant, QTY As Variant, indexCell As Variant) As String
     With Workbooks(tempWorkBookName)
     
        If Replace(OriQPN, " ", "") <> "" And Replace(NewQPN, " ", "") <> "" Then
             .Worksheets(manuSheetName).Cells(indexCell, 3).Interior.color = RGB(255, 0, 0)
             .Worksheets(manuSheetName).Cells(indexCell, 4).Interior.color = RGB(255, 0, 0)
            delItemCheck = "Please check which QPN you want to remove."
        ElseIf Replace(OriQPN, " ", "") = "" Then
            .Worksheets(manuSheetName).Cells(indexCell, 3).Interior.color = RGB(255, 0, 0)
            delItemCheck = "Please fill out QPN in Oringinal Part Number."
        ElseIf Replace(OriQPN, " ", "") <> "" And Replace(NewQPN, " ", "") = "" Then
            'Do nothing
        Else
            MsgBox "Exception! _functionModule_delItemCheck_1"
        End If
     End With
End Function

Private Sub errorReport(errorMessage As Variant)

    

    'Dim fso As Object
    'Dim oFile As Object
    
    'Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Set oFile = fso.CreateTextFile(tempWorkBookNamePath & "\" & Replace(tempWorkBookName, ".xlsm", "") & "_errorMsg.txt")
    
        'oFile.WriteLine "Please check the follow problem..."
    'For Each erMesg In errorMessage
    '    oFile.WriteLine erMesg
    'Next

    'Set fso = Nothing
    'Set oFile = Nothing

    UserForm3.Show vbModeless
    
    UserForm3.ListBox1.Clear
    
    UserForm3.ListBox1.addItem "Please check the follow problem..."
    
    For Each tempMsg In errorMessage
    
        UserForm3.ListBox1.addItem tempMsg
     
    Next
    
    
    
    
End Sub


Function check_excelFile_open(filePath As Variant) As Boolean
    
    
    For i = 1 To Application.Workbooks.Count
        If Workbooks(i).name = filePath Then
        
            check_excelFile_open = True
            Exit Function
            
        End If
    Next
    
    check_excelFile_open = False
    
            
    
    
End Function
