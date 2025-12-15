' ============ Filing Module ============
Sub RunFiling()
    Call SortFilingByDescription
    Call filingBat__
End Sub

Sub SortFilingByDescription()
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Filing")

    LastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ws.Sort.SortFields.Clear

    ws.Sort.SortFields.Add2 Key:=ws.Range("F2:F" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range("A1:H" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub splitBat()
    Dim shell As Object
    'Dim ws As Worksheet
    'Dim lastRow As Long

    'Set ws = ThisWorkbook.Sheets("Filing")
    'With ws
    '    Union(.Range("J3:R4"), .Range("S3:T4"), .Range("T5:5000")).ClearContents
    'End With

    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat"
End Sub

Sub checkBat()
    Dim shell As Object
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Filing")
    With ws
        Union(.Range("J3:R4"), .Range("S3:T4"), .Range("T5:T5000")).ClearContents
    End With
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat", 0, False
End Sub

Sub filingBat__()
    Dim shell As Object
    Dim toPath As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Filing")
    With ws
        Union(.Range("J3:R4"), .Range("S3:T4"), .Range("T5:T5000")).ClearContents
    End With
    
    toPath = ""
    
    If Dir(toPath, vbDirectory) <> "" Then
        file = Dir(toPath & "*.*")
        Do While file <> ""
            Kill toPath & file
            file = Dir
        Loop
    End If
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat", 0, False
End Sub

Sub mergeBat__()
    Dim shell As Object
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Filing")
    With ws
        Union(.Range("J3:R4"), .Range("S3:T4"), .Range("T5:T5000")).ClearContents
    End With
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat", 0, False
End Sub

Sub runBoth()
    Call d_inv_Bat
    Call d_nts_Bat
End Sub
Sub d_inv_Bat()
    Dim shell As Object
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Filing")
    With ws
        Union(.Range("J3:R4"), .Range("S3:T4"), .Range("T5:T5000")).ClearContents
    End With
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat", 0, False
End Sub

Sub d_nts_Bat()
    Dim shell As Object
    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Filing")
    With ws
        Union(.Range("J3:R4"), .Range("S3:T4"), .Range("T5:T5000")).ClearContents
    End With
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run "batch.bat", 0, False
End Sub

Sub DownloadCDOFromFilingSheet()
    Dim BaseURL As String
    Dim PDFName As String
    Dim SavePDFName As String
    Dim FileURL As String
    Dim SavePath As String
    Dim LastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim xmlhttp As Object
    Dim adoStream As Object
    Dim FolderPath As String
    Dim FileDownloaded As Boolean
    
    ' Set base URL
    BaseURL = ""
    
    ' Set download folder
    FolderPath = ""
    
    ' Check if folder exists, if not, create it
    If Dir(FolderPath, vbDirectory) = "" Then MkDir FolderPath
    
    ' Set target worksheet
    Set ws = ThisWorkbook.Sheets("Filing")
    
    ' Find last row with data in column H
    LastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Reset download flag
    FileDownloaded = False
    
    ' Loop through each row from H2 down
    For i = 2 To LastRow
        PDFName = Trim(ws.Range("H" & i).Value)
        SavePDFName = Trim(ws.Range("A" & i).Value)
        
        If PDFName <> "" And SavePDFName <> "" Then
            If Not Right(SavePDFName, 4) = ".pdf" Then
                SavePDFName = SavePDFName & ".pdf"
            End If
            
            FileURL = BaseURL & PDFName
            SavePath = FolderPath & SavePDFName
            
            Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
            xmlhttp.Open "GET", FileURL, False
            xmlhttp.setRequestHeader "Accept", "application/pdf"
            xmlhttp.Send
            
            If xmlhttp.Status = 200 Then
                Set adoStream = CreateObject("ADODB.Stream")
                adoStream.Type = 1
                adoStream.Open
                adoStream.Write xmlhttp.responseBody
                adoStream.SaveToFile SavePath, 2
                adoStream.Close
                Set adoStream = Nothing
                FileDownloaded = True
            Else
                Debug.Print "Failed to download: " & FileURL
            End If
        End If
    Next i
    
    ' Show result message in merged cell J3
    With ws.Range("J3")
        .Font.Size = 16
        .Font.Bold = True
        '.Font.Color = RGB(0, 0, 0)
        
        If FileDownloaded Then
            .Value = "Download completed.."
        Else
            .Value = "{404} No CDO found"
        End If
    End With
    
End Sub

