' ========= GenerateSummary(Set page), HighlightSheetNames, count CDO, Download CDO, Filter ST, BD and DEL Module =========
Sub GenerateSummary()
    Dim ws As Worksheet
    Dim SummaryWs As Worksheet
    Dim outputRow As Long
    Dim LastRow As Long
    Dim targetSheetName As String

    Set SummaryWs = ThisWorkbook.Sheets("Set")
    targetSheetName = SummaryWs.Range("H1").Value ' Sheet to process for PDFs
    
    ' ==== Clear previous summary ====
    With SummaryWs.Range("C7:E100")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    ' ==== Clear Columns in summary ====
    With SummaryWs
        Union(.Range("F5:H6"), .Range("I5:R6"), .Range("S5:T6"), .Range("T7:T200")).ClearContents
    End With

    
    ' ==== Format summary description ====
    With SummaryWs.Range("C6:E6")
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
        .Borders.LineStyle = xlNone
    End With
    
    outputRow = 7

    ' ==== Loop through all sheets to fill summary table ====
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Set" And ws.Name <> "Filing" Then
            SummaryWs.Cells(outputRow, "C").Value = ws.Name
            SummaryWs.Cells(outputRow, "D").Value = ws.Range("H1").Value
            SummaryWs.Cells(outputRow, "E").Value = ws.Range("H2").Value

            outputRow = outputRow + 1
        End If
    Next ws

    LastRow = outputRow - 1

    ' ==== Format the summary table ====
    With SummaryWs.Range("C7:E" & LastRow)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 10
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(191, 191, 191)
    End With

    ' ==== Subtotal row ====
    SummaryWs.Cells(LastRow + 1, "C").Value = "Total:"
    SummaryWs.Cells(LastRow + 1, "D").Formula = "=SUBTOTAL(9, D7:D" & LastRow & ")"
    SummaryWs.Cells(LastRow + 1, "E").Formula = "=SUBTOTAL(9, E7:E" & LastRow & ")"

    With SummaryWs.Range("C" & LastRow + 1 & ":E" & LastRow + 1)
        .Interior.Color = RGB(132, 151, 176)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(132, 151, 176)
    End With
    
    ' ==== Find the last row in column C starting from C7 ====
    LastRow = SummaryWs.Cells(SummaryWs.Rows.Count, "C").End(xlUp).Row
    
    If LastRow > 0 Then LastRow = LastRow - 1
    
    ' ==== Define the range of values in column C from C7 to the last row ====
    uniqueRange = "C7:C" & LastRow
    
    ' ==== Remove any existing validation before applying new one ====
    SummaryWs.Range("F5").Validation.Delete
    
    ' ==== Apply data validation to cell F5 (list based on unique values in column C) ====
    With SummaryWs.Range("F5").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & uniqueRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub HighlightSheetNames()
    Dim selectedSheet As String
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim totalRow As Long
    Dim i As Long
    Dim rowRange As Range

    ' Get the selected sheet name from cell F5 on the "Set" sheet
    selectedSheet = ThisWorkbook.Sheets("Set").Range("F5").Value

    ' Set the target worksheet (the "Set" sheet)
    Set ws = ThisWorkbook.Sheets("Set")

    ' Get the last used row in column C starting from row 7
    LastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Find total row in column C (case-insensitive search for "total")
    totalRow = 0
    For i = 7 To LastRow
        If LCase(ws.Cells(i, "C").Value) Like "*total*" Then
            totalRow = i
            Exit For
        End If
    Next i

    ' === Step 1: Clear previous highlights excluding total row ===
    If totalRow > 0 Then
        If totalRow = 7 Then
            With ws.Range("C8:E" & LastRow)
                .Interior.ColorIndex = xlNone
                .Font.ColorIndex = xlAutomatic
                .Font.Bold = False
                .Font.Size = 10
            End With
        ElseIf totalRow = LastRow Then
            With ws.Range("C7:E" & LastRow - 1)
                .Interior.ColorIndex = xlNone
                .Font.ColorIndex = xlAutomatic
                .Font.Bold = False
                .Font.Size = 10
            End With
        Else
            With ws.Range("C7:E" & totalRow - 1)
                .Interior.ColorIndex = xlNone
                .Font.ColorIndex = xlAutomatic
                .Font.Bold = False
                .Font.Size = 10
            End With
            
            With ws.Range("C" & totalRow + 1 & ":E" & LastRow)
                .Interior.ColorIndex = xlNone
                .Font.ColorIndex = xlAutomatic
                .Font.Bold = False
                .Font.Size = 10
            End With
        End If
    Else
        With ws.Range("C7:E" & LastRow)
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
            .Font.Size = 10
        End With
    End If

    ' === Step 2: Loop through rows excluding total row and apply highlight only if match ===
    For i = 7 To LastRow
        If i <> totalRow Then
            If ws.Cells(i, "C").Value = selectedSheet Or _
               ws.Cells(i, "D").Value = selectedSheet Or _
               ws.Cells(i, "E").Value = selectedSheet Then

                ' Highlight matching row
                Set rowRange = ws.Range("C" & i & ":E" & i)
                rowRange.Interior.Color = RGB(208, 206, 206)   ' Highlight color
                rowRange.Font.Color = RGB(0, 0, 0)             ' Black font color
                rowRange.Font.Size = 11
                rowRange.Font.Bold = True
            End If
        End If
    Next i

    ' Format the total row if found
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

Sub CountCDOs()
    Dim selectedSheet As String
    Dim ws As Worksheet
    Dim targetWs As Worksheet
    Dim LastRow As Long
    Dim cdoCount As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Set")
    selectedSheet = ws.Range("F5").Value

    ' Check if the sheet selected exists
    On Error Resume Next
    Set targetWs = ThisWorkbook.Sheets(selectedSheet)
    On Error GoTo 0

    If targetWs Is Nothing Then
        ws.Range("I5").Value = ""
        Exit Sub
    End If

    ' Get the last row in column I (starting from row 4)
    LastRow = targetWs.Cells(targetWs.Rows.Count, "I").End(xlUp).Row

    ' Initialize the counter
    cdoCount = 0

    ' Loop through column I (starting from row 4)
    For i = 4 To LastRow
        If targetWs.Cells(i, "I").Value <> "" Then
            cdoCount = cdoCount + 1
        End If
    Next i

    ' Display the result in cell
    If cdoCount > 0 Then
        ws.Range("I5").Value = cdoCount & " CDO Found"
    Else
        ws.Range("I5").Value = "No CDO Found"
    End If
End Sub

Sub DownloadCDOFromSelectedSheet()
    Dim BaseURL As String
    Dim PDFNameToDownload As String
    Dim SavePDFName As String
    Dim SavePath1 As String
    Dim SavePath2 As String
    Dim FileURL As String
    Dim adoStream As Object
    Dim xmlhttp As Object
    Dim Folder1 As String
    Dim Folder2 As String
    Dim PageName As String
    Dim selectedSheet As String
    Dim ws As Worksheet
    Dim i As Integer
    Dim LastRow As Integer
    Dim FileFound As Boolean
    
    ' Set your base URL here
    BaseURL = "" 'set your url......................
    
    ' Get the selected sheet name from cell F5 in the "Set" sheet
    selectedSheet = ThisWorkbook.Sheets("Set").Range("F5").Value
    
    ' Check if a valid sheet is selected
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(selectedSheet)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Invalid sheet name in F5!"
        Exit Sub
    End If
    
    ' Find the last row with data in column I (PDF names) and column B (save names)
    LastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Define your folders here
    Folder1 = ""
    Folder2 = ""
    
    ' Check if Folder1 exists, if not, create it
    If Len(Dir(Folder1, vbDirectory)) = 0 Then
        MkDir Folder1
    End If
    
    ' Check if Folder2 exists, if not, create it
    If Len(Dir(Folder2, vbDirectory)) = 0 Then
        MkDir Folder2
    End If
    
    ' Initialize the flag that tracks whether any files were found
    FileFound = False
    
    ' Loop through rows starting from row 4 to LastRow
    For i = 4 To LastRow
        ' Get the PDF name from column I and the save name from column B
        PDFNameToDownload = ws.Range("I" & i).Value
        SavePDFName = ws.Range("B" & i).Value
        
        ' If either the PDF name or the save name is empty, skip this row
        If PDFNameToDownload = "" Or SavePDFName = "" Then
            GoTo SkipDownload
        End If
        
        ' Construct the full file URL
        FileURL = BaseURL & PDFNameToDownload
        
        ' Ensure the file name ends with .pdf
        If Not Right(SavePDFName, 4) = ".pdf" Then
            SavePDFName = SavePDFName & ".pdf"
        End If
        
        ' Set the save paths for both folders
        SavePath1 = Folder1 & SavePDFName  ' Save to Folder1
        SavePath2 = Folder2 & SavePDFName  ' Save to Folder2
        
        ' Create the xmlhttp object
        Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
        
        ' Set headers to simulate a direct download
        xmlhttp.Open "GET", FileURL, False
        xmlhttp.setRequestHeader "Accept", "application/pdf"
        xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        xmlhttp.setRequestHeader "Content-Type", "application/pdf"
        xmlhttp.Send
        
        If xmlhttp.Status = 200 Then
            ' Create ADO stream to save the file to Folder1
            Set adoStream = CreateObject("ADODB.Stream")
            adoStream.Type = 1 ' Binary
            adoStream.Open
            adoStream.Write xmlhttp.responseBody
            adoStream.SaveToFile SavePath1, 2 ' Overwrite if exists
            adoStream.Close
            Set adoStream = Nothing

            ' Now save the same file to Folder2
            Set adoStream = CreateObject("ADODB.Stream")
            adoStream.Type = 1 ' Binary
            adoStream.Open
            adoStream.Write xmlhttp.responseBody
            adoStream.SaveToFile SavePath2, 2 ' Overwrite if exists
            adoStream.Close
            Set adoStream = Nothing
            
            ' Set FileFound flag to True
            FileFound = True
        Else
            MsgBox "Failed to download file from URL: " & FileURL
        End If
        
SkipDownload:
    Next i
    
    ' After the loop, check if any file was found
    If Not FileFound Then
        ' Display the message "No file found" in cell O5 on the "Set" sheet
        With ThisWorkbook.Sheets("Set").Range("I5")
            .Value = "{404}"
            .Font.Color = RGB(38, 38, 38)
            ' .Interior.Color = RGB(132, 151, 176)
            .Font.Size = 16
            .Font.Bold = True
        End With
    Else
        ' Clear any previous "No file found" message
        ThisWorkbook.Sheets("Set").Range("I5").Value = "CDO Download completed"
    End If
    
End Sub

Sub Filteriiiiiiiiiiiiiii_ST()
    Dim ws As Worksheet
    Dim SummaryWs As Worksheet
    Dim outputRow As Long
    Dim LastRow As Long
    Dim targetSheetName As String

    ' Set references
    Set SummaryWs = ThisWorkbook.Sheets("Set")
    targetSheetName = SummaryWs.Range("H1").Value ' Sheet to process for PDFs
    
    ' Clear previous summary
    With SummaryWs.Range("C7:E1000")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    With SummaryWs.Range("C6:E6")
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
    End With
    
    outputRow = 7

    ' Loop through all sheets to fill summary table
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Left(ws.Name, 2)) = "st" Then
            ' Write summary info
            SummaryWs.Cells(outputRow, "C").Value = ws.Name
            SummaryWs.Cells(outputRow, "D").Value = ws.Range("H1").Value
            SummaryWs.Cells(outputRow, "E").Value = ws.Range("H2").Value
    
            outputRow = outputRow + 1
        End If
    Next ws

    LastRow = outputRow - 1

    ' Format the summary table
    With SummaryWs.Range("C7:E" & LastRow)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 10
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(191, 191, 191)
    End With

    ' Subtotal row
    SummaryWs.Cells(LastRow + 1, "C").Value = "Total:"
    SummaryWs.Cells(LastRow + 1, "D").Formula = "=SUBTOTAL(9, D7:D" & LastRow & ")"
    SummaryWs.Cells(LastRow + 1, "E").Formula = "=SUBTOTAL(9, E7:E" & LastRow & ")"

    With SummaryWs.Range("C" & LastRow + 1 & ":E" & LastRow + 1)
        .Interior.Color = RGB(132, 151, 176)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(132, 151, 176)
    End With
    
    ' Set the reference to the "Set" sheet
    Set SummaryWs = ThisWorkbook.Sheets("Set")
    
    ' Find the last row in column C starting from C7
    LastRow = SummaryWs.Cells(SummaryWs.Rows.Count, "C").End(xlUp).Row
    
    If LastRow > 0 Then LastRow = LastRow - 1
    
    ' Define the range of values in column C from C7 to the last row
    uniqueRange = "C7:C" & LastRow
    
    ' Remove any existing validation before applying new one
    SummaryWs.Range("F5").Validation.Delete
    
    ' Apply data validation to cell F5 (list based on unique values in column C)
    With SummaryWs.Range("F5").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & uniqueRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub Filteriiiiiiiiiiiiiii_BD()
    Dim ws As Worksheet
    Dim SummaryWs As Worksheet
    Dim outputRow As Long
    Dim LastRow As Long
    Dim targetSheetName As String

    ' Set references
    Set SummaryWs = ThisWorkbook.Sheets("Set")
    targetSheetName = SummaryWs.Range("H1").Value ' Sheet to process for PDFs
    
    ' Clear previous summary
    With SummaryWs.Range("C7:E1000")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    With SummaryWs.Range("C6:E6")
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
    End With
    
    outputRow = 7

    ' Loop through all sheets to fill summary table
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Left(ws.Name, 2)) = "bd" Then
            ' Write summary info
            SummaryWs.Cells(outputRow, "C").Value = ws.Name
            SummaryWs.Cells(outputRow, "D").Value = ws.Range("H1").Value
            SummaryWs.Cells(outputRow, "E").Value = ws.Range("H2").Value
    
            outputRow = outputRow + 1
        End If
    Next ws

    LastRow = outputRow - 1

    ' Format the summary table
    With SummaryWs.Range("C7:E" & LastRow)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 10
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(191, 191, 191)
    End With

    ' Subtotal row
    SummaryWs.Cells(LastRow + 1, "C").Value = "Total:"
    SummaryWs.Cells(LastRow + 1, "D").Formula = "=SUBTOTAL(9, D7:D" & LastRow & ")"
    SummaryWs.Cells(LastRow + 1, "E").Formula = "=SUBTOTAL(9, E7:E" & LastRow & ")"

    With SummaryWs.Range("C" & LastRow + 1 & ":E" & LastRow + 1)
        .Interior.Color = RGB(132, 151, 176)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(132, 151, 176)
    End With
    
    ' Set the reference to the "Set" sheet
    Set SummaryWs = ThisWorkbook.Sheets("Set")
    
    ' Find the last row in column C starting from C7
    LastRow = SummaryWs.Cells(SummaryWs.Rows.Count, "C").End(xlUp).Row
    
    If LastRow > 0 Then LastRow = LastRow - 1
    
    ' Define the range of values in column C from C7 to the last row
    uniqueRange = "C7:C" & LastRow
    
    ' Remove any existing validation before applying new one
    SummaryWs.Range("F5").Validation.Delete
    
    ' Apply data validation to cell F5 (list based on unique values in column C)
    With SummaryWs.Range("F5").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & uniqueRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub Filteriiiiiiiiiiiiiii_DEL()
    Dim ws As Worksheet
    Dim SummaryWs As Worksheet
    Dim outputRow As Long
    Dim LastRow As Long
    Dim targetSheetName As String

    ' Set references
    Set SummaryWs = ThisWorkbook.Sheets("Set")
    targetSheetName = SummaryWs.Range("H1").Value ' Sheet to process for PDFs
    
    ' Clear previous summary
    With SummaryWs.Range("C7:E1000")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
    
    With SummaryWs.Range("C6:E6")
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
    End With
    
    outputRow = 7

    ' Loop through all sheets to fill summary table
    For Each ws In ThisWorkbook.Worksheets
        If LCase(Left(ws.Name, 2)) = "de" Then
            ' Write summary info
            SummaryWs.Cells(outputRow, "C").Value = ws.Name
            SummaryWs.Cells(outputRow, "D").Value = ws.Range("H1").Value
            SummaryWs.Cells(outputRow, "E").Value = ws.Range("H2").Value
    
            outputRow = outputRow + 1
        End If
    Next ws

    LastRow = outputRow - 1

    ' Format the summary table
    With SummaryWs.Range("C7:E" & LastRow)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 10
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(191, 191, 191)
    End With

    ' Subtotal row
    SummaryWs.Cells(LastRow + 1, "C").Value = "Total:"
    SummaryWs.Cells(LastRow + 1, "D").Formula = "=SUBTOTAL(9, D7:D" & LastRow & ")"
    SummaryWs.Cells(LastRow + 1, "E").Formula = "=SUBTOTAL(9, E7:E" & LastRow & ")"

    With SummaryWs.Range("C" & LastRow + 1 & ":E" & LastRow + 1)
        .Interior.Color = RGB(132, 151, 176)
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 13
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(132, 151, 176)
    End With
    
    ' Set the reference to the "Set" sheet
    Set SummaryWs = ThisWorkbook.Sheets("Set")
    
    ' Find the last row in column C starting from C7
    LastRow = SummaryWs.Cells(SummaryWs.Rows.Count, "C").End(xlUp).Row
    
    If LastRow > 0 Then LastRow = LastRow - 1
    
    ' Define the range of values in column C from C7 to the last row
    uniqueRange = "C7:C" & LastRow
    
    ' Remove any existing validation before applying new one
    SummaryWs.Range("F5").Validation.Delete
    
    ' Apply data validation to cell F5 (list based on unique values in column C)
    With SummaryWs.Range("F5").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & uniqueRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub


