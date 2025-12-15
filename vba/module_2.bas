' ============== Import XLS, Print, Jump to page and Delete all pages Module ==============
Dim counter As Integer ' Declare the counter variable globally

Sub IncrementValue()

    counter = counter + 1
    Range("Q11").Value = counter
    
End Sub

Sub DecrementValue()

    If counter > 1 Then
        counter = counter - 1
    End If

    Range("Q11").Value = counter
End Sub

Sub ImportSheets_____()
    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim filePath As String
    Dim fd As FileDialog
    Dim ws As Worksheet
    Dim sheetName As String
    
    Set wbDest = ThisWorkbook
    
    ' Open file dialog to select workbook to import
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Open source workbook
    Set wbSource = Workbooks.Open(filePath)
    
    ' Loop through sheets in source workbook
    For Each ws In wbSource.Sheets
        sheetName = ws.Name
        
        ' Check if sheet name exists in destination workbook
        If Not SheetExists(sheetName, wbDest) Then
            ' Copy sheet only if it doesn't exist
            ws.Copy After:=wbDest.Sheets(wbDest.Sheets.Count)
        Else
            ' Skip sheets with duplicate names
            Debug.Print "Skipped duplicate sheet: " & sheetName
        End If
    Next ws
    
    wbSource.Close SaveChanges:=False
End Sub

' Helper function to check if a sheet exists in a workbook
Function SheetExists(shtName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(shtName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Sub PrintSheet()
    Dim wsSet As Worksheet
    Dim targetSheetName As String
    Dim ws As Worksheet
    Dim numCopies As Long
    
    Set wsSet = ThisWorkbook.Sheets("Set")
    targetSheetName = Trim(wsSet.Range("F5").Value)
    numCopies = Val(wsSet.Range("Q11").Value) ' Get number of copies (defaults to 0 if blank/invalid)
    
    If numCopies < 1 Then numCopies = 1 ' Ensure at least 1 copy
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & targetSheetName & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' === Autofit columns in range A1:L24 ===
    ws.Range("A:L").Columns.AutoFit
    
    With ws.PageSetup
        ' === Margins ===
        .TopMargin = Application.InchesToPoints(0.6)
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.7)
        .HeaderMargin = Application.InchesToPoints(0.15)
        .FooterMargin = Application.InchesToPoints(0.1)
        
        ' === HEADER SETTINGS ===
        ' Left: Logo (replace with actual image path)
        .LeftHeaderPicture.Filename = "image.jpeg"
        .LeftHeader = "&G" ' Show the image

        ' Center: Page X of Y
        .CenterHeader = vbLf & "Page &P of &N"

        ' Right: Current time
        .RightHeader = vbLf & Format(Time, "hh:mm AM/PM")
        
        ' Force Excel to paginate correctly
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        
        ' === FOOTER SETTINGS ===
        .LeftFooter = "For Security"
        .CenterFooter = "For Incharge"
        .RightFooter = "For Courier"
    End With

    ' Print specified number of copies
    ws.PrintOut Copies:=numCopies, Collate:=True
     'ws.PrintPreview
    
End Sub

Sub JumpToSheetFromCellValue()
    Dim targetSheetName As String
    Dim ws As Worksheet

    ' Get the value in cell C5 of the "Data" sheet
    targetSheetName = ThisWorkbook.Sheets("Set").Range("F5").Value

    ' Check if a sheet with that name exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ' Activate the sheet
        ws.Activate
    End If
End Sub

Sub DeleteAllSheets()
    Dim ws As Worksheet
    Dim name1, name2 As String
    name1 = "Set"
    name2 = "Filing"

    Application.DisplayAlerts = False ' Disable prompts
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> name1 And ws.Name <> name2 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True ' Re-enable prompts
End Sub

