Attribute VB_Name = "Module1"
Public SelectedFilePath As String

' ==========================================
' Step 1: Select File and Populate Sheet List
' ==========================================
Sub Step1_SelectFileAndGetSheets()
    Dim fd As FileDialog
    Dim wbTarget As Workbook
    Dim wsTool As Worksheet
    Dim i As Long
    Dim sheetList As String
    
    Set wsTool = ThisWorkbook.Sheets(1)
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Please select the source Excel file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsb; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            SelectedFilePath = .SelectedItems(1)
            wsTool.Range("C3").Value = SelectedFilePath
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbTarget = Workbooks.Open(Filename:=SelectedFilePath, ReadOnly:=True)
    
    sheetList = ""
    For i = 1 To wbTarget.Sheets.Count
        sheetList = sheetList & wbTarget.Sheets(i).Name & ","
    Next i
    
    If Len(sheetList) > 0 Then sheetList = Left(sheetList, Len(sheetList) - 1)
    
    wbTarget.Close SaveChanges:=False
    
    With wsTool.Range("C5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=sheetList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    If Len(sheetList) > 0 Then
        wsTool.Range("C5").Value = Split(sheetList, ",")(0)
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "File loaded successfully! Please select a sheet in cell C5.", vbInformation, "Success"
End Sub

' ==========================================
' Step 2: Check Column D (Numeric/Text) & Copy D to AO (38 Cols)
' ==========================================
Sub Step2_ProcessAndExport()
    Dim wsTool As Worksheet
    Dim targetFilePath As String
    Dim targetSheetName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim maxRows As Long
    Dim r As Long, c As Long
    Dim matchCount As Long
    Dim targetCell As Range
    Dim dataRange As Range
    Dim rawVal As String
    Dim cleanVal As String
    Dim savePath As String
    Dim folderPath As String
    Dim resultData() As Variant
    
    Set wsTool = ThisWorkbook.Sheets(1)
    targetFilePath = wsTool.Range("C3").Value
    targetSheetName = wsTool.Range("C5").Value
    
    If targetFilePath = "" Or Dir(targetFilePath) = "" Then
        MsgBox "Please select a valid Excel file in Step 1 first.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If targetSheetName = "" Then
        MsgBox "Please select a sheet in cell C5.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    Set wbSource = Workbooks.Open(Filename:=targetFilePath, ReadOnly:=True)
    
    On Error Resume Next
    Set wsSource = wbSource.Sheets(targetSheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Sheet not found: " & targetSheetName, vbCritical, "Error"
        wbSource.Close SaveChanges:=False
        GoTo CleanUp
    End If
    
    ' Unhide all rows and columns to handle hidden data
    wsSource.Rows.Hidden = False
    wsSource.Columns.Hidden = False
    
    ' Find max rows in sheet
    maxRows = wsSource.UsedRange.Rows.Count + wsSource.UsedRange.Row - 1
    If maxRows < 1 Then maxRows = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    
    If maxRows < 1 Then
        MsgBox "No data found in the selected sheet.", vbExclamation, "Warning"
        wbSource.Close SaveChanges:=False
        GoTo CleanUp
    End If

    ' Array size for Column D to AO -> 38 columns (D=4, AO=41)
    ReDim resultData(1 To maxRows, 1 To 38)
    matchCount = 0

    ' Sweep rows and validate Column D
    For r = 1 To maxRows
        Set targetCell = wsSource.Cells(r, 4) ' Column D
        
        ' Handle merged cells safely
        If targetCell.MergeCells Then
            rawVal = CStr(targetCell.MergeArea.Cells(1, 1).Value)
        Else
            rawVal = CStr(targetCell.Value)
        End If
        
        ' Clean value (Remove spaces, Chr 160, and single quotes)
        cleanVal = Trim(Replace(Replace(rawVal, Chr(160), ""), "'", ""))
        
        ' Check if cleaned value is Numeric (Handles both Number and Text-Number)
        If cleanVal <> "" And IsNumeric(cleanVal) Then
            matchCount = matchCount + 1
            ' Copy D to AO (Columns 4 to 41 -> 38 Columns)
            For c = 1 To 38
                Set dataRange = wsSource.Cells(r, c + 3) ' Column D offset
                If dataRange.MergeCells Then
                    resultData(matchCount, c) = dataRange.MergeArea.Cells(1, 1).Value
                Else
                    resultData(matchCount, c) = dataRange.Value
                End If
            Next c
        End If
    Next r

    wbSource.Close SaveChanges:=False

    If matchCount = 0 Then
        MsgBox "No numeric rows found in Column D.", vbInformation, "Notice"
        GoTo CleanUp
    End If

    ' Create Output File
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Sheets(1)
    
    ' Set Format to Text (@) to preserve leading zeros
    wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(matchCount, 38)).NumberFormat = "@"
    
    ' Output Data to New Sheet
    wsNew.Range("A1").Resize(matchCount, 38).Value = resultData
    wsNew.Columns("A:AL").AutoFit
    
    folderPath = Left(targetFilePath, InStrRev(targetFilePath, "\"))
    savePath = folderPath & "Cleaned_Upload_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=True

    MsgBox "Process completed successfully!" & vbCrLf & _
           "Total records matched: " & matchCount & " rows" & vbCrLf & _
           "Saved at: " & vbCrLf & savePath, vbInformation, "Success"

CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
    End With
End Sub


