Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if the change happened in cell F5
    If Not Intersect(Target, Me.Range("F5")) Is Nothing Then
        ' Call the macro to check for data and update O5
        ' Call HighlightSheetNames
        Call CountCDOs
    End If
End Sub

Private Sub Worksheet_Activate()
    Call GenerateSummary
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim totalRow As Long
    Dim clickedRow As Long
    Dim rng1 As Range, rng2 As Range
    Dim combinedRange As Range
    Dim i As Long

    ' === Safety checks ===
    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge > 1 Then Exit Sub

    Set ws = Me
    
    ' === Set column widths for C, D, and E ===
    'With ws
    '    .Columns("C").ColumnWidth = 10 ' Approx. 70 pixels
    '    .Columns("D").ColumnWidth = 10 ' Approx. 75 pixels
    '    .Columns("E").ColumnWidth = 10 ' Approx. 75 pixels
    'End With

    ' === Get the last used row in column C, starting from row 7 ===
    LastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If LastRow < 7 Then LastRow = 7 ' Safety guard: don't go above row 7

    ' === Always update cell F5 on "Set" sheet if any cell in C7:E is selected ===
    If Not Intersect(Target, ws.Range("C7:E" & LastRow - 1)) Is Nothing Then
        ws.Range("S5:T6").ClearContents
        
        lastRRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
        If lastRRow < 7 Then lastRRow = 7
        ws.Range("T7:T" & lastRRow).ClearContents
        
        clickedRow = Target.Row
        ThisWorkbook.Sheets("Set").Range("F5").Value = ws.Cells(clickedRow, "C").Value
    End If

    ' === Find the "total" row in column C (case-insensitive match) ===
    totalRow = 0
    For i = 7 To LastRow
        If LCase(ws.Cells(i, "C").Value) Like "*total*" Then
            totalRow = i
            Exit For
        End If
    Next i
    
    ' === Build a combined range excluding the "total" row ===
    Set rng1 = Nothing
    Set rng2 = Nothing
    Set combinedRange = Nothing

    If totalRow > 0 Then
        If totalRow = 7 Then
            If LastRow > 7 Then Set combinedRange = ws.Range("C8:E" & LastRow)
        ElseIf totalRow = LastRow Then
            If LastRow > 7 Then Set combinedRange = ws.Range("C7:E" & LastRow - 1)
        Else
            If totalRow - 1 >= 7 Then Set rng1 = ws.Range("C7:E" & totalRow - 1)
            If totalRow + 1 <= LastRow Then Set rng2 = ws.Range("C" & totalRow + 1 & ":E" & LastRow)

            If Not rng1 Is Nothing And Not rng2 Is Nothing Then
                Set combinedRange = Union(rng1, rng2)
            ElseIf Not rng1 Is Nothing Then
                Set combinedRange = rng1
            ElseIf Not rng2 Is Nothing Then
                Set combinedRange = rng2
            End If
        End If
    Else
        If LastRow >= 7 Then Set combinedRange = ws.Range("C7:E" & LastRow)
    End If
    
    ' === Clear highlights in the combined range, if defined ===
    If Not combinedRange Is Nothing Then
        With combinedRange
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
            .Font.Size = 10
        End With
    End If

    ' === Highlight the selected row if it's inside the combined range ===
    If Not combinedRange Is Nothing Then
        If Not Intersect(Target, combinedRange) Is Nothing Then
            If Target.Row <> totalRow Then
                With ws.Range("C" & Target.Row & ":E" & Target.Row)
                    .Interior.Color = RGB(208, 206, 206) ' Highlight
                    .Font.Color = RGB(0, 0, 0)
                    .Font.Size = 11
                    .Font.Bold = True
                End With
            End If
        End If
    End If

    ' === Always format the "total" row if it exists ===
    If totalRow > 0 Then
        With ws.Range("C" & totalRow & ":E" & totalRow)
            .Interior.Color = RGB(132, 151, 176)
            .Font.Color = RGB(0, 0, 0)
            .Font.Size = 13
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(132, 151, 176)
        End With
    End If

End Sub

