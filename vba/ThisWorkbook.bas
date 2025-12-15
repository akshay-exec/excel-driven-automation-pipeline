Private Sub Workbook_Open()
    ' Sheets("Set").Activate
    Call GenerateSummary
    Call DeleteAllSheets
    
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Filing")

    With ws
        Union(.Range("J3:R4"), .Range("S3:T4"), .Range("T5:T5000")).ClearContents
    End With

    On Error Resume Next
    LastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0

    If LastRow >= 2 Then
        ws.Range("A2:H" & LastRow).Clear
    End If
End Sub