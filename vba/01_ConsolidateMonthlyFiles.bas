Attribute VB_Name = "Module1"
Option Explicit

' Macro 1: merges every HR_MIS_*.xlsx (sheet "Data") in a chosen folder
' into one sheet called Master_Data in this workbook.

Public Sub ConsolidateMonthlyFiles()
    Dim folderPath As String, fName As String
    Dim wsMaster As Worksheet, wbSrc As Workbook, wsSrc As Worksheet
    Dim lastRow As Long, lastCol As Long, nextRow As Long
    Dim filesDone As Long, totalRows As Long, dataRows As Long
    Dim log As String

    folderPath = PickFolder()
    If folderPath = "" Then Exit Sub
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Create or clear the Master_Data sheet
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Worksheets("Master_Data")
    On Error GoTo Cleanup
    If wsMaster Is Nothing Then
        Set wsMaster = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMaster.Name = "Master_Data"
    Else
        wsMaster.Cells.Clear
    End If

    nextRow = 1
    fName = Dir(folderPath & "HR_MIS_*.xlsx")

    Do While fName <> ""
        Set wbSrc = Workbooks.Open(folderPath & fName, ReadOnly:=True)
        Set wsSrc = Nothing
        On Error Resume Next
        Set wsSrc = wbSrc.Worksheets("Data")
        On Error GoTo Cleanup

        If wsSrc Is Nothing Then
            log = log & fName & ": sheet 'Data' not found, skipped" & vbCrLf
        Else
            lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

            If nextRow = 1 Then
                ' first file: copy header + data, add Source_File header
                wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsMaster.Cells(1, 1)
                wsMaster.Cells(1, lastCol + 1).Value = "Source_File"
                dataRows = lastRow - 1
                nextRow = lastRow + 1
            Else
                ' later files: data rows only (no header)
                wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsMaster.Cells(nextRow, 1)
                dataRows = lastRow - 1
                nextRow = nextRow + dataRows
            End If

            wsMaster.Range(wsMaster.Cells(nextRow - dataRows, lastCol + 1), wsMaster.Cells(nextRow - 1, lastCol + 1)).Value = fName
            totalRows = totalRows + dataRows
            filesDone = filesDone + 1
            log = log & fName & ": " & dataRows & " rows" & vbCrLf
        End If

        wbSrc.Close SaveChanges:=False
        fName = Dir()
    Loop

    wsMaster.Rows(1).Font.Bold = True
    wsMaster.Columns.AutoFit
    Application.CutCopyMode = False

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Consolidation failed"
    ElseIf filesDone = 0 Then
        MsgBox "No HR_MIS_*.xlsx files found in " & folderPath, vbExclamation
    Else
        MsgBox filesDone & " files merged, " & totalRows & " rows in Master_Data." & vbCrLf & vbCrLf & log, vbInformation, "Consolidation complete"
    End If
End Sub

Private Function PickFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the data_raw folder with the monthly HR_MIS files"
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function
