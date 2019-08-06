Attribute VB_Name = "LinktoServer"
Sub linkToServer()

    Dim testPath As String
    
    On Error GoTo Exception1
    testPath = "\\10.242.184.38\Department\EE_Arthur\Public_access\04_Technical document\05_Tools\Excel\Data\testLink.txt"
    
    mf = FreeFile
    
    Open testPath For Input As #mf
    
    
    Do Until EOF(1)
    
        Line Input #1, textline
         If textline = "BU9_Arthur" Then GoTo EndPoint

    Loop
    
    
EndPoint:
    Call Module1.start
    Exit Sub
    
    
    
    
Exception1:
    MsgBox "Error: Can not Link to Server"
    Exit Sub





End Sub
