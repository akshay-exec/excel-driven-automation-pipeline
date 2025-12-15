' ==================== BAT Module Section ====================
Sub ShowMessage()
    MsgBox ">>> Work-in-progress Not Completed <<<", vbInformation, "From_exec....."
End Sub

Sub RunBatINV_______()
    Dim shell As Object
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Set")
    ws.Range("S5:T6").ClearContents
    
    LastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    If LastRow >= 7 Then
        ws.Range("T7:T" & LastRow).ClearContents
    End If
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat", 0, False
End Sub

Sub RunBatNTS_______()
    Dim shell As Object
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Set")
    ws.Range("S5:T6").ClearContents
    
    LastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    If LastRow >= 7 Then
        ws.Range("T7:T" & LastRow).ClearContents
    End If
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat", 0, False
End Sub

Sub RunBatMain()
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat"
End Sub
Sub ShowMessage_()
    MsgBox ">>> ========== Work-in-progress ========== <<<", vbInformation, "From_exec....."
End Sub
