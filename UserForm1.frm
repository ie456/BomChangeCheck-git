VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4464
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8460.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox1_Click()
    
    
    If CheckBox1.value Then
        MsgBox "NOTICE : New Bom doesn't load now!! "
    Else
    
    End If
    
   
End Sub

Private Sub CheckBox2_Change()


 'If Me.MultiPage1.value = 3 Then
    Call subModule.enableManuChang(UserForm1.CheckBox2.value)
 'End If
 
    functionModule.updateUserFormValue (42)
 

End Sub

Private Sub CheckBox3_Click()
    If Not Me.CheckBox3.value And Not Me.CheckBox4.value Then
        Me.CheckBox3.value = True
    End If
End Sub

Private Sub CheckBox4_Click()
    If Not Me.CheckBox3.value And Not Me.CheckBox4.value Then
        Me.CheckBox4.value = True
    End If
End Sub



Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()

End Sub

'Userform1 EXCEL file select
'select EXCEL new file
'
Private Sub CommandButton1_Click()
        
        'getPath(sheetName As String, saveLocaPath As String, saveLocaSheet As String, upSelNum As Integer) As String
        'TextBox1.Enabled = True
        tempValue = functionModule.getPath("Main", "A34", "B34", 4)
        
        If tempValue <> "" Then
            TextBox1 = tempValue
        End If
        'TextBox1.Enabled = False
End Sub

'Userform1 EXCEL file select
'select EXCEL old file
'
Private Sub CommandButton2_Click()

    
        'getPath(sheetName As String, saveLocaPath As String, saveLocaSheet As String, upSelNum As Integer) As String
        'TextBox2.Enabled = True
        
        tempValue = functionModule.getPath("Main", "A35", "B35", 5)
        
        If tempValue <> "" Then
            TextBox2 = tempValue
        End If
        'TextBox2.Enabled = False
End Sub

Private Sub CommandButton20_Click() 'test

    Call subModule.manuChageList
    functionModule.updateUserFormValue (6)
    
End Sub

Private Sub CommandButton21_Click()

    tempFilter = " File (*.xls;*.xlsx;*.xlsm), *.xls;*.xlsx;*.xlsm"
    
    tempValue = functionModule.getPath_Gene("Main", "A36", "B36", 41, tempFilter)
    
    
    
    If tempValue <> "" Then
        TextBox23 = tempValue
    End If
End Sub

Private Sub CommandButton22_Click()
        sheetName = Workbooks(tempWorkBookName).Worksheets("Main").Range("I40").value
    Call Module1.load_Gene(sheetName, UserForm1.TextBox23.Text, UserForm1.ComboBox3.Text)
    Call subModule.saveData("Main", UserForm1.ComboBox3.Text, "C36")
End Sub








'Userform1 EXCEL file select
'select Load
'
Private Sub CommandButton3_Click()
    'Call Module1.load_Ex
    
    'On Error GoTo exit1
    
    Call Module1.load
    functionModule.updateUserFormValue (6)
    
    
    If UserForm1.CheckBox2.value Then
        With Workbooks(tempWorkBookName)
            .Worksheets(.Worksheets("Main").Range("I40").value).Move after:=.Worksheets(.Worksheets.Count)
        End With
    End If
    
    'If Me.CheckBox2.value And Me.CommandButton13.Enabled Then
    '    Me.MultiPage1.value = 3
    '
    '    Me.CommandButton13.Enabled = False
    '     Call saveData("Main", False, "M30")
    'End If
    
'exit1:
    
    
End Sub

Private Sub CommandButton13_Click()
          
          
     
          
          
    'Call Module1.load_sum
    tempAry = functionModule.load_sum
    
    'On Error GoTo exit2
    
     With Application
                .ScreenUpdating = False
                .DisplayAlerts = False
                .EnableEvents = False
    End With
    

     
    
    
    If Me.CheckBox3.value Then Call Module1.simpleCompare(tempAry)
    If Me.CheckBox4.value Then Call Module1.compare_2(tempAry)
    
    'If UserForm1.OptionButton1.value = True Then
    '    Call Module1.simpleCompare(tempAry)
    'ElseIf UserForm1.OptionButton2.value = True Then
    '    Call Module1.compare(tempAry)
    'ElseIf UserForm1.OptionButton3.value = True Then
    '    Call Module1.compare_2(tempAry)
    'Else
    '
    'End If
    
    functionModule.updateUserFormValue (6)
    
   
    
    'If UserForm1.CheckBox2.value Then
    '    With Workbooks(tempWorkBookName)
    '        .Worksheets(.Worksheets("Main").Range("I40").value).Move after:=.Worksheets(.Worksheets.Count)

    '    End With
    'End If
    
     
    
    MsgBox "Done."
    
    GoTo exit1
    
exit2:
    MsgBox "Please check *_SUM data validity"
    
exit1:

    With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
    End With
    
    
End Sub

Private Sub CommandButton15_Click() 'add
    
    
    For i = 0 To ListBox2.ListCount - 1
        
        If ListBox2.Selected(i) Then
        
            ListBox3.addItem (ListBox2.List(i))
        
        End If
    
    Next
    
    For i = ListBox2.ListCount - 1 To 0 Step -1

        If ListBox2.Selected(i) = True Then
            ListBox2.RemoveItem (i)
            Exit For
        End If
    
    Next
    
    
    
End Sub

Private Sub CommandButton16_Click() 'remove

    For i = 0 To ListBox3.ListCount - 1
        
        If ListBox3.Selected(i) Then
        
            ListBox2.addItem (ListBox3.List(i))
        
        End If
    
    Next
    
    For i = ListBox3.ListCount - 1 To 0 Step -1

        If ListBox3.Selected(i) = True Then

            ListBox3.RemoveItem (i)
            Exit For
        End If
    
    Next
End Sub

Private Sub CommandButton17_Click() 'add all
    For i = 0 To ListBox2.ListCount - 1
            ListBox3.addItem (ListBox2.List(i))
    Next
    
    For i = ListBox2.ListCount - 1 To 0 Step -1
      ListBox2.RemoveItem (i)
    Next
End Sub

Private Sub CommandButton18_Click() 'remove all

    functionModule.updateUserFormValue 0
    
    For i = ListBox3.ListCount - 1 To 0 Step -1


        ListBox3.RemoveItem (i)

    
    Next
End Sub

Private Sub CommandButton19_Click()

   ' On Error GoTo exit1
    temp = InputBox("Please key-in file name.", "SET FILE NAME", "DEFAULT")
    temp = Replace(temp, " ", "")
    If temp = "DEFAULT" Then
        fileName = "DEFAULT_" & Format(Now, "yyyymmdd_hhmm")
    ElseIf temp <> "" Then
        fileName = temp
    ElseIf temp = "" Then
        Exit Sub
    Else
        MsgBox "Exception"
    End If
    
    Call Module1.printOut(fileName)
    
exit1:
    
End Sub



Private Sub CommandButton11_Click()

     On Error GoTo exitNewCh
    
        
        If TextBox11 <> "" Then
            TextBox11 = functionModule.getPath_Rpt
            Worksheets("Main").Range("A36").value = TextBox11
        End If
        
exitNewCh:

End Sub

Private Sub CommandButton12_Click()

     On Error GoTo exitNewCh
    
        
        If TextBox12 <> "" Then
            TextBox12 = functionModule.getPath_Rpt
            Worksheets("Main").Range("A37").value = TextBox12
        End If

exitNewCh:

End Sub



Private Sub CommandButton8_Click()  'Reset
        
        functionModule.unProtectSheet ("Main")
        Worksheets("Main").Range("A30").value = 2
         Worksheets("Main").Range("A31").value = "B"
         Worksheets("Main").Range("A32").value = "D"
         Worksheets("Main").Range("A33").value = "E"
         
         Worksheets("Main").Range("B30").value = 2
         Worksheets("Main").Range("B31").value = "B"
         Worksheets("Main").Range("B32").value = "D"
         Worksheets("Main").Range("B33").value = "E"
         functionModule.protectSheet ("Main")
         
         functionModule.updateUserFormValue 0
End Sub
Private Sub CommandButton24_Click()
        functionModule.unProtectSheet ("Main")
        Worksheets("Main").Range("A30").value = 8
         Worksheets("Main").Range("A31").value = "B"
         Worksheets("Main").Range("A32").value = "D"
         Worksheets("Main").Range("A33").value = "E"
         
         functionModule.protectSheet ("Main")
         
         functionModule.updateUserFormValue 0
End Sub
Private Sub CommandButton27_Click()
    functionModule.unProtectSheet ("Main")
        
         
         Worksheets("Main").Range("B30").value = 8
         Worksheets("Main").Range("B31").value = "B"
         Worksheets("Main").Range("B32").value = "D"
         Worksheets("Main").Range("B33").value = "E"
         functionModule.protectSheet ("Main")
         
         functionModule.updateUserFormValue 0
End Sub
Private Sub CommandButton32_Click()
    functionModule.unProtectSheet ("Main")
        
         
         Worksheets("Main").Range("C30").value = "A"
         Worksheets("Main").Range("C31").value = "N"
         Worksheets("Main").Range("C32").value = "O"
         Worksheets("Main").Range("C33").value = "G"
         functionModule.protectSheet ("Main")
         
         functionModule.updateUserFormValue 0
End Sub
Private Sub CommandButton9_Click() 'Save
        
    Call subModule.saveSetData
        

End Sub
Private Sub CommandButton23_Click()

     Call subModule.saveNewSetData
            
End Sub

Private Sub CommandButton28_Click()
     Call subModule.saveOldSetData
        
End Sub

Private Sub CommandButton31_Click()
    Call subModule.saveChangeSetData
End Sub

Private Sub CommandButton25_Click()
    Me.CommandButton23.Visible = False
    Me.CommandButton24.Visible = False
    Me.CommandButton25.Visible = False
    
    Me.TextBox5.Enabled = False
    Me.TextBox6.Enabled = False
    Me.TextBox7.Enabled = False
    Me.TextBox19.Enabled = False
    
    
End Sub

Private Sub CommandButton26_Click()
     Me.CommandButton26.Visible = False
    Me.CommandButton27.Visible = False
    Me.CommandButton28.Visible = False
    
    Me.TextBox13.Enabled = False
    Me.TextBox14.Enabled = False
    Me.TextBox15.Enabled = False
    Me.TextBox21.Enabled = False
    

    
    
    
End Sub




Private Sub CommandButton33_Click()
    Me.CommandButton31.Visible = False
    Me.CommandButton32.Visible = False
    Me.CommandButton33.Visible = False
    
    Me.TextBox28.Enabled = False
    Me.TextBox24.Enabled = False
    Me.TextBox26.Enabled = False
    Me.TextBox30.Enabled = False
    Me.CheckBox2.SetFocus
    
End Sub

Private Sub Frame1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.CommandButton23.Visible = True
    Me.CommandButton24.Visible = True
    Me.CommandButton25.Visible = True
    
    Me.TextBox5.Enabled = True
    Me.TextBox6.Enabled = True
    Me.TextBox7.Enabled = True
    Me.TextBox19.Enabled = True
    
    
End Sub
Private Sub Frame4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.CommandButton26.Visible = True
    Me.CommandButton27.Visible = True
    Me.CommandButton28.Visible = True
    
    Me.TextBox13.Enabled = True
    Me.TextBox14.Enabled = True
    Me.TextBox15.Enabled = True
    Me.TextBox21.Enabled = True
    
    
End Sub
Private Sub Frame8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.CommandButton31.Visible = True
    Me.CommandButton32.Visible = True
    Me.CommandButton33.Visible = True
    
    Me.TextBox24.Enabled = True
    Me.TextBox26.Enabled = True
    Me.TextBox28.Enabled = True
    Me.TextBox30.Enabled = True
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame5_Click()
    'Call subModule.getSavePath
End Sub


Private Sub Label1_Click()

End Sub

Private Sub Label17_Click()
    Call subModule.getSavePath
End Sub





Private Sub TextBox31_Change()

End Sub

Private Sub UserForm_Activate()
     
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 180
    Me.Left = Application.Left + Application.Width - Me.Width - 50
    
    
     
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Call subModule.saveData("Main", UserForm1.CheckBox3.value, "N30")
        Call subModule.saveData("Main", UserForm1.CheckBox4.value, "N31")
        
        Call subModule.saveData("Main", UserForm1.CheckBox1.value, "O30")
        Call subModule.saveData("Main", Label17, "L30")
        Call subModule.saveData("Main", UserForm1.CheckBox2.value, "L34")
        
        'Call subModule.saveData_2(Replace(tempWorkBookName, ".xlsm", ".csv"), Replace(tempWorkBookName, ".xlsm", ""), UserForm1.CheckBox3.value, "A2")
        
        
        
    End If
End Sub
