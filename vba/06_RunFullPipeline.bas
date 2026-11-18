Attribute VB_Name = "Module6"
Option Explicit

' Macro 5: Summary sheet with month selector, conditional formatting, refresh buttons.
' RefreshReport     = rebuilds helper columns, pivots, formatting and Summary from Master_Data
' RunFullPipeline   = the whole monthly update: Macros 1, cleaning, 2, 3, then RefreshReport

Public Sub RunFullPipeline()
    If MsgBox("This rebuilds Master_Data from the monthly files and refreshes all reports." & vbCrLf & _
              "You will be asked to pick the data_raw folder and employee_master.csv." & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Full monthly update") <> vbYes Then Exit Sub
    ConsolidateMonthlyFiles
    If Not SheetExists("Master_Data") Then Exit Sub
    CleanMasterData
    ValidateMasterData
    ApplyDataFixes
    RefreshReport
End Sub

Public Sub RefreshReport()
    If Not PrepareHelperColumns() Then Exit Sub
    BuildPivotReports True
    ApplyFormatting
    BuildSummary
    BuildDashboard True
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Summary").Activate
    MsgBox "Reports refreshed.", vbInformation, "Done"
End Sub

Private Sub ApplyFormatting()
    Dim pt As PivotTable, ws As Worksheet, fc As Object

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("R3_Attrition")
    On Error GoTo 0
    If Not ws Is Nothing Then
        For Each pt In ws.PivotTables
            With pt.DataBodyRange
                .FormatConditions.Delete
                Set fc = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0.02")
                fc.Interior.Color = RGB(255, 199, 206)
                fc.Font.Color = RGB(156, 0, 6)
            End With
        Next pt
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("R4_Attendance")
    On Error GoTo 0
    If Not ws Is Nothing Then
        For Each pt In ws.PivotTables
            If pt.Name Like "*Att" Then
                With pt.DataBodyRange
                    .FormatConditions.Delete
                    Set fc = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0.9")
                    fc.Interior.Color = RGB(255, 235, 156)
                    fc.Font.Color = RGB(156, 87, 0)
                End With
            End If
        Next pt
    End If
End Sub

Private Sub BuildSummary()
    Dim ws As Worksheet, wm As Worksheet, lastRow As Long, i As Long, k As Long
    Dim cKey As Long, cOp As Long, cJn As Long, cEx As Long, cWk As Long
    Dim cPr As Long, cAb As Long, cUn As Long, cFl As Long
    Dim rKey As String, rOp As String, rJn As String, rEx As String
    Dim rWk As String, rPr As String, rAb As String, rUn As String, rFl As String
    Dim keys As Object, arr As Variant, firstTime As Boolean, k2 As Variant, cur As String
    Dim b As Object

    Set wm = ThisWorkbook.Worksheets("Master_Data")
    cKey = ColIdx(wm, "Month_Key"): cOp = ColIdx(wm, "Opening_HC"): cJn = ColIdx(wm, "Joiners")
    cEx = ColIdx(wm, "Exits"): cWk = ColIdx(wm, "Att_Working"): cPr = ColIdx(wm, "Att_Present")
    cAb = ColIdx(wm, "Att_Absent"): cUn = ColIdx(wm, "Att_Unauth"): cFl = ColIdx(wm, "DQ_Exclude")
    lastRow = wm.Cells(wm.Rows.Count, 1).End(xlUp).Row

    rKey = RangeText(cKey, lastRow): rOp = RangeText(cOp, lastRow): rJn = RangeText(cJn, lastRow)
    rEx = RangeText(cEx, lastRow): rWk = RangeText(cWk, lastRow): rPr = RangeText(cPr, lastRow)
    rAb = RangeText(cAb, lastRow): rUn = RangeText(cUn, lastRow)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Summary")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "Summary"
        firstTime = True
    End If

    If firstTime Then
        ws.Cells(1, 1).Value = "Monthly HR MIS Summary"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 1).Font.Size = 16
        ws.Cells(3, 1).Value = "Select month (yyyy-mm)"
        ws.Cells(3, 1).Font.Bold = True
        ws.Range("A5:A13").Value = Application.Transpose(Array("Opening headcount", "New joiners", "Employees exited", _
            "Closing headcount", "Attrition % (monthly)", "Attendance %", "Absenteeism %", _
            "Unauthorised absence %", "Rows excluded from attendance (DQ)"))
        ws.Cells(15, 1).Value = "Last refreshed"
        ws.Columns(1).ColumnWidth = 36
        ws.Columns(2).ColumnWidth = 14
        ws.Range("B3").Interior.Color = RGB(255, 242, 204)
        ws.Range("B3").Font.Bold = True
    End If

    ' month list (kept in hidden column H)
    Set keys = CreateObject("Scripting.Dictionary")
    arr = wm.Range(wm.Cells(2, cKey), wm.Cells(lastRow, cKey)).Value
    For i = 1 To UBound(arr, 1)
        If CStr(arr(i, 1)) <> "" Then
            If Not keys.Exists(CStr(arr(i, 1))) Then keys.Add CStr(arr(i, 1)), 1
        End If
    Next i
    ws.Range("H1:H100").ClearContents
    ws.Range("H1:H100").NumberFormat = "@"
    ws.Range("H1").Value = "Months"
    k = 1
    For Each k2 In keys.keys
        k = k + 1
        ws.Cells(k, 8).Value = CStr(k2)
    Next k2
    ws.Columns("H").Hidden = True

    ws.Range("B3").NumberFormat = "@"
    With ws.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$H$2:$H$" & k
    End With
    cur = CStr(ws.Range("B3").Value)
    If cur = "" Or Not keys.Exists(cur) Then ws.Range("B3").Value = CStr(ws.Cells(k, 8).Value)

    ' KPI formulas (SUMPRODUCT compares text exactly, so "2026-05" is never read as a date)
    ws.Range("B5").Formula = "=SUMPRODUCT((" & rKey & "=$B$3)*" & rOp & ")"
    ws.Range("B6").Formula = "=SUMPRODUCT((" & rKey & "=$B$3)*" & rJn & ")"
    ws.Range("B7").Formula = "=SUMPRODUCT((" & rKey & "=$B$3)*" & rEx & ")"
    ws.Range("B8").Formula = "=B5+B6-B7"
    ws.Range("B9").Formula = "=IF((B5+B8)=0,0,B7/((B5+B8)/2))"
    ws.Range("B10").Formula = "=IFERROR(SUMPRODUCT((" & rKey & "=$B$3)*" & rPr & ")/SUMPRODUCT((" & rKey & "=$B$3)*" & rWk & "),0)"
    ws.Range("B11").Formula = "=IFERROR(SUMPRODUCT((" & rKey & "=$B$3)*" & rAb & ")/SUMPRODUCT((" & rKey & "=$B$3)*" & rWk & "),0)"
    ws.Range("B12").Formula = "=IFERROR(SUMPRODUCT((" & rKey & "=$B$3)*" & rUn & ")/SUMPRODUCT((" & rKey & "=$B$3)*" & rWk & "),0)"
    If cFl > 0 Then
        rFl = RangeText(cFl, lastRow)
        ws.Range("B13").Formula = "=SUMPRODUCT((" & rKey & "=$B$3)*(" & rFl & "=""Y""))"
    Else
        ws.Range("B13").Value = 0
    End If
    ws.Range("B5:B8").NumberFormat = "#,##0"
    ws.Range("B13").NumberFormat = "#,##0"
    ws.Range("B9:B12").NumberFormat = "0.0%"
    ws.Range("B15").Value = Now
    ws.Range("B15").NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Range("A8:B8").Font.Bold = True

    ' conditional formatting on the KPI cells
    ws.Range("B9").FormatConditions.Delete
    With ws.Range("B9").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0.02")
        .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
    End With
    ws.Range("B10").FormatConditions.Delete
    With ws.Range("B10").FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0.9")
        .Interior.Color = RGB(255, 235, 156): .Font.Color = RGB(156, 87, 0)
    End With

    ' buttons (created once)
    If ws.Buttons.Count = 0 Then
        Set b = ws.Buttons.Add(ws.Range("D3").Left, ws.Range("D3").Top, 170, 28)
        b.Caption = "Rebuild reports"
        b.OnAction = "RefreshReport"
        Set b = ws.Buttons.Add(ws.Range("D5").Left, ws.Range("D5").Top, 170, 28)
        b.Caption = "Run full monthly update"
        b.OnAction = "RunFullPipeline"
    End If
End Sub

Private Function RangeText(c As Long, lastRow As Long) As String
    Dim L As String
    L = ColLetter(c)
    RangeText = "Master_Data!$" & L & "$2:$" & L & "$" & lastRow
End Function

Private Function ColLetter(c As Long) As String
    Dim a As String
    a = ThisWorkbook.Worksheets("Master_Data").Cells(1, c).Address(True, False)
    ColLetter = Left(a, InStr(a, "$") - 1)
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not s Is Nothing
End Function

Private Function ColIdx(ws As Worksheet, hdr As String) As Long
    Dim f As Range
    Set f = ws.Rows(1).Find(What:=hdr, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then ColIdx = 0 Else ColIdx = f.Column
End Function

