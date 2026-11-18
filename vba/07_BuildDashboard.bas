Attribute VB_Name = "Module7"
Option Explicit

' Macro 6: builds the Dashboard sheet (KPI cards + 6 charts) from Master_Data helper columns.
' Needs RefreshReport to have run (helper columns Month_Key, Opening_HC, ... must exist).
' Run RunDashboard from Developer > Macros.

Public Sub RunDashboard()
    BuildDashboard
End Sub

Public Sub BuildDashboard(Optional quiet As Boolean = False)
    Dim wm As Worksheet, wd As Worksheet, ws As Worksheet
    Dim lastRow As Long, n As Long, i As Long, k As Long
    Dim keys As Object, arr As Variant, kv As Variant
    Dim cKey As Long, cOp As Long, cJn As Long, cEx As Long, cWk As Long, cPr As Long
    Dim cAb As Long, cDept As Long, cLoc As Long, cLT As Long, cLV As Long
    Dim rKey As String, rOp As String, rJn As String, rEx As String, rWk As String
    Dim rPr As String, rAb As String, rDept As String, rLoc As String, rLT As String, rLV As String
    Dim sel As String, depts As Variant, locs As Variant, ltypes As Variant
    Dim ch As Chart, s1 As Series, s2 As Series, last As String, mRng As String

    On Error Resume Next
    Set wm = ThisWorkbook.Worksheets("Master_Data")
    On Error GoTo 0
    If wm Is Nothing Then MsgBox "Master_Data sheet not found.", vbExclamation: Exit Sub

    cKey = ColIdx(wm, "Month_Key"): cOp = ColIdx(wm, "Opening_HC"): cJn = ColIdx(wm, "Joiners")
    cEx = ColIdx(wm, "Exits"): cWk = ColIdx(wm, "Att_Working"): cPr = ColIdx(wm, "Att_Present")
    cAb = ColIdx(wm, "Att_Absent"): cDept = ColIdx(wm, "Department"): cLoc = ColIdx(wm, "Location")
    cLT = ColIdx(wm, "Leave_Type"): cLV = ColIdx(wm, "Leave_Days")
    If cKey = 0 Or cOp = 0 Or cJn = 0 Or cEx = 0 Or cWk = 0 Or cPr = 0 Or cAb = 0 Or cDept = 0 Or cLoc = 0 Or cLT = 0 Or cLV = 0 Then
        MsgBox "Helper columns are missing. Run RefreshReport first.", vbExclamation: Exit Sub
    End If

    lastRow = wm.Cells(wm.Rows.Count, 1).End(xlUp).Row
    rKey = RangeText(wm, cKey, lastRow): rOp = RangeText(wm, cOp, lastRow): rJn = RangeText(wm, cJn, lastRow)
    rEx = RangeText(wm, cEx, lastRow): rWk = RangeText(wm, cWk, lastRow): rPr = RangeText(wm, cPr, lastRow)
    rAb = RangeText(wm, cAb, lastRow): rDept = RangeText(wm, cDept, lastRow): rLoc = RangeText(wm, cLoc, lastRow)
    rLT = RangeText(wm, cLT, lastRow): rLV = RangeText(wm, cLV, lastRow)

    ' list of months
    Set keys = CreateObject("Scripting.Dictionary")
    arr = wm.Range(wm.Cells(2, cKey), wm.Cells(lastRow, cKey)).Value
    For i = 1 To UBound(arr, 1)
        If CStr(arr(i, 1)) <> "" Then
            If Not keys.Exists(CStr(arr(i, 1))) Then keys.Add CStr(arr(i, 1)), 1
        End If
    Next i
    n = keys.Count
    If n = 0 Then MsgBox "No months found in Master_Data.", vbExclamation: Exit Sub
    last = CStr(n + 1)

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    DeleteSheet "Dashboard"
    DeleteSheet "Dash_Data"
    Application.DisplayAlerts = True

    Set wd = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wd.Name = "Dash_Data"
    If SheetExists("Summary") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Summary"))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
    End If
    ws.Name = "Dashboard"
    sel = "Dashboard!$C$3"

    ' ----- Dash_Data: monthly trend table (A:I) -----
    wd.Range("A1:I1").Value = Array("Month", "Opening", "Joiners", "Exits", "Closing headcount", _
                                    "Attrition %", "Attendance %", "Absenteeism %", "Attendance target")
    wd.Range("A2:A" & last).NumberFormat = "@"
    i = 1
    For Each kv In keys.keys
        i = i + 1
        wd.Cells(i, 1).Value = CStr(kv)
        wd.Cells(i, 2).Formula = "=SUMPRODUCT((" & rKey & "=$A" & i & ")*" & rOp & ")"
        wd.Cells(i, 3).Formula = "=SUMPRODUCT((" & rKey & "=$A" & i & ")*" & rJn & ")"
        wd.Cells(i, 4).Formula = "=SUMPRODUCT((" & rKey & "=$A" & i & ")*" & rEx & ")"
        wd.Cells(i, 5).Formula = "=B" & i & "+C" & i & "-D" & i
        wd.Cells(i, 6).Formula = "=IF((B" & i & "+E" & i & ")=0,0,D" & i & "/((B" & i & "+E" & i & ")/2))"
        wd.Cells(i, 7).Formula = "=IFERROR(SUMPRODUCT((" & rKey & "=$A" & i & ")*" & rPr & ")/SUMPRODUCT((" & rKey & "=$A" & i & ")*" & rWk & "),0)"
        wd.Cells(i, 8).Formula = "=IFERROR(SUMPRODUCT((" & rKey & "=$A" & i & ")*" & rAb & ")/SUMPRODUCT((" & rKey & "=$A" & i & ")*" & rWk & "),0)"
        wd.Cells(i, 9).Value = 0.9
    Next kv
    wd.Range("F2:I" & last).NumberFormat = "0.0%"

    ' ----- Dash_Data: department table for the selected month (K:P) -----
    depts = Array("Admin", "Customer Support", "Finance", "HR", "IT Services", "Legal", "Operations", "Sales")
    wd.Range("K1:P1").Value = Array("Department", "Opening", "Joiners", "Exits", "Attrition %", "Attendance %")
    For k = 0 To UBound(depts)
        i = k + 2
        wd.Cells(i, 11).Value = depts(k)
        wd.Cells(i, 12).Formula = "=SUMPRODUCT((" & rKey & "=" & sel & ")*(" & rDept & "=$K" & i & ")*" & rOp & ")"
        wd.Cells(i, 13).Formula = "=SUMPRODUCT((" & rKey & "=" & sel & ")*(" & rDept & "=$K" & i & ")*" & rJn & ")"
        wd.Cells(i, 14).Formula = "=SUMPRODUCT((" & rKey & "=" & sel & ")*(" & rDept & "=$K" & i & ")*" & rEx & ")"
        wd.Cells(i, 15).Formula = "=IF((2*L" & i & "+M" & i & "-N" & i & ")=0,0,N" & i & "/((2*L" & i & "+M" & i & "-N" & i & ")/2))"
        wd.Cells(i, 16).Formula = "=IFERROR(SUMPRODUCT((" & rKey & "=" & sel & ")*(" & rDept & "=$K" & i & ")*" & rPr & ")/SUMPRODUCT((" & rKey & "=" & sel & ")*(" & rDept & "=$K" & i & ")*" & rWk & "),0)"
    Next k
    wd.Range("O2:P9").NumberFormat = "0.0%"

    ' ----- Dash_Data: location table (R:S) -----
    locs = Array("Bangalore", "Chennai", "Hyderabad", "Kolkata", "Mumbai", "Pune")
    wd.Range("R1:S1").Value = Array("Location", "Closing headcount")
    For k = 0 To UBound(locs)
        i = k + 2
        wd.Cells(i, 18).Value = locs(k)
        wd.Cells(i, 19).Formula = "=SUMPRODUCT((" & rKey & "=" & sel & ")*(" & rLoc & "=$R" & i & ")*(" & rOp & "+" & rJn & "-" & rEx & "))"
    Next k

    ' ----- Dash_Data: leave table (U:V) -----
    ltypes = Array("Casual Leave", "Earned Leave", "Sick Leave")
    wd.Range("U1:V1").Value = Array("Leave type", "Leave days taken")
    For k = 0 To UBound(ltypes)
        i = k + 2
        wd.Cells(i, 21).Value = ltypes(k)
        wd.Cells(i, 22).Formula = "=SUMPRODUCT((" & rKey & "=" & sel & ")*(" & rLT & "=$U" & i & ")*" & rLV & ")"
    Next k
    wd.Rows(1).Font.Bold = True
    wd.Columns("A:V").ColumnWidth = 16

    ' ----- Dashboard layout -----
    ws.Columns("A").ColumnWidth = 2
    ws.Range("B:M").ColumnWidth = 11.5
    ws.Range("B1").Value = "HR MIS Dashboard"
    ws.Range("B1").Font.Size = 20
    ws.Range("B1").Font.Bold = True
    ws.Range("B3").Value = "Select month"
    ws.Range("B3").Font.Bold = True
    ws.Range("C3").NumberFormat = "@"
    ws.Range("C3").Value = CStr(wd.Cells(n + 1, 1).Value)
    ws.Range("C3").Interior.Color = RGB(255, 242, 204)
    ws.Range("C3").Font.Bold = True
    ws.Range("C3").HorizontalAlignment = xlCenter
    With ws.Range("C3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Dash_Data!$A$2:$A$" & last
    End With
    ws.Range("E3").Value = "KPI cards and the department, location and leave charts follow the selected month."
    ws.Range("E3").Font.Italic = True
    ws.Range("E3").Font.Color = RGB(110, 110, 110)

    mRng = "Dash_Data!$A$2:$A$" & last
    MakeCard ws, 2, "Closing headcount", "=IFERROR(INDEX(Dash_Data!$E$2:$E$" & last & ",MATCH($C$3," & mRng & ",0)),0)", "#,##0", "opening + joiners - exits"
    MakeCard ws, 5, "Attrition % (monthly)", "=IFERROR(INDEX(Dash_Data!$F$2:$F$" & last & ",MATCH($C$3," & mRng & ",0)),0)", "0.0%", "red when above 2%"
    MakeCard ws, 8, "Attendance %", "=IFERROR(INDEX(Dash_Data!$G$2:$G$" & last & ",MATCH($C$3," & mRng & ",0)),0)", "0.0%", "amber when below 90%"
    MakeCard ws, 11, "Absenteeism % (incl. leave)", "=IFERROR(INDEX(Dash_Data!$H$2:$H$" & last & ",MATCH($C$3," & mRng & ",0)),0)", "0.0%", "leave + unauthorised absence"

    ws.Range("E6:G6").FormatConditions.Delete
    With ws.Range("E6:G6").FormatConditions.Add(Type:=xlExpression, Formula1:="=$E$6>0.02")
        .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
    End With
    ws.Range("H6:J6").FormatConditions.Delete
    With ws.Range("H6:J6").FormatConditions.Add(Type:=xlExpression, Formula1:="=$H$6<0.9")
        .Interior.Color = RGB(255, 235, 156): .Font.Color = RGB(156, 87, 0)
    End With

    ' ----- charts -----
    Set ch = NewChart(ws, "B10", "G24", "Closing headcount and attrition % by month", xlColumnClustered)
    Set s1 = AddSer(ch, "Closing headcount", wd.Range("E2:E" & last), wd.Range("A2:A" & last))
    s1.Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
    s1.HasDataLabels = True
    s1.DataLabels.NumberFormat = "#,##0"
    Set s2 = AddSer(ch, "Attrition %", wd.Range("F2:F" & last), wd.Range("A2:A" & last))
    s2.ChartType = xlLineMarkers
    s2.AxisGroup = xlSecondary
    s2.Format.Line.ForeColor.RGB = RGB(237, 125, 49)
    ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.0%"
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom

    Set ch = NewChart(ws, "H10", "M24", "Attendance % by month vs 90% target", xlLineMarkers)
    Set s1 = AddSer(ch, "Attendance %", wd.Range("G2:G" & last), wd.Range("A2:A" & last))
    s1.Format.Line.ForeColor.RGB = RGB(31, 78, 121)
    s1.HasDataLabels = True
    s1.DataLabels.NumberFormat = "0.0%"
    s1.DataLabels.Position = xlLabelPositionAbove
    Set s2 = AddSer(ch, "Target 90%", wd.Range("I2:I" & last), wd.Range("A2:A" & last))
    s2.ChartType = xlLine
    s2.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
    s2.Format.Line.DashStyle = msoLineDash
    ch.Axes(xlValue).MinimumScale = 0.85
    ch.Axes(xlValue).MaximumScale = 0.95
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom

    Set ch = NewChart(ws, "B26", "G40", "Attrition % by department (selected month)", xlBarClustered)
    Set s1 = AddSer(ch, "Attrition %", wd.Range("O2:O9"), wd.Range("K2:K9"))
    s1.Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
    s1.HasDataLabels = True
    s1.DataLabels.NumberFormat = "0.0%"
    ch.HasLegend = False
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"

    Set ch = NewChart(ws, "H26", "M40", "Closing headcount by location (selected month)", xlColumnClustered)
    Set s1 = AddSer(ch, "Closing headcount", wd.Range("S2:S7"), wd.Range("R2:R7"))
    s1.Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
    s1.HasDataLabels = True
    s1.DataLabels.NumberFormat = "#,##0"
    ch.HasLegend = False

    Set ch = NewChart(ws, "B42", "G56", "Leave days by type (selected month)", xlDoughnut)
    Set s1 = AddSer(ch, "Leave days", wd.Range("V2:V4"), wd.Range("U2:U4"))
    s1.HasDataLabels = True
    s1.DataLabels.ShowPercentage = True
    s1.DataLabels.ShowValue = False
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight

    Set ch = NewChart(ws, "H42", "M56", "Attendance % by department (selected month)", xlBarClustered)
    Set s1 = AddSer(ch, "Attendance %", wd.Range("P2:P9"), wd.Range("K2:K9"))
    s1.Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
    s1.HasDataLabels = True
    s1.DataLabels.NumberFormat = "0.0%"
    ch.HasLegend = False
    ch.Axes(xlValue).MinimumScale = 0.8
    ch.Axes(xlValue).MaximumScale = 1
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("C3").Select
    Application.ScreenUpdating = True
    If Not quiet Then MsgBox "Dashboard built. Use the yellow cell to change the month.", vbInformation, "Dashboard ready"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Dashboard build failed"
End Sub

' ---------- helpers ----------
Private Sub MakeCard(ws As Worksheet, c As Long, lbl As String, f As String, nf As String, cap As String)
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(5, c), ws.Cells(7, c + 2))
    rg.Interior.Color = RGB(242, 242, 242)
    rg.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(191, 191, 191)
    ws.Cells(5, c).Value = lbl
    ws.Cells(6, c).Formula = f
    ws.Cells(6, c).NumberFormat = nf
    ws.Cells(7, c).Value = cap
    ws.Range(ws.Cells(5, c), ws.Cells(5, c + 2)).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(ws.Cells(6, c), ws.Cells(6, c + 2)).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(ws.Cells(7, c), ws.Cells(7, c + 2)).HorizontalAlignment = xlCenterAcrossSelection
    ws.Cells(5, c).Font.Bold = True
    ws.Cells(5, c).Font.Color = RGB(89, 89, 89)
    ws.Cells(6, c).Font.Size = 26
    ws.Cells(6, c).Font.Bold = True
    ws.Cells(7, c).Font.Size = 9
    ws.Cells(7, c).Font.Color = RGB(128, 128, 128)
    ws.Rows(6).RowHeight = 40
End Sub

Private Function NewChart(ws As Worksheet, tl As String, br As String, ttl As String, ctype As Long) As Chart
    Dim rg As Range, co As ChartObject
    Set rg = ws.Range(tl & ":" & br)
    Set co = ws.ChartObjects.Add(rg.Left + 2, rg.Top + 2, rg.Width - 4, rg.Height - 4)
    With co.Chart
        .ChartType = ctype
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        .HasTitle = True
        .ChartTitle.Text = ttl
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    Set NewChart = co.Chart
End Function

Private Function AddSer(ch As Chart, nm As String, vals As Range, cats As Range) As Series
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.Values = vals
    s.XValues = cats
    Set AddSer = s
End Function

Private Function RangeText(ws As Worksheet, c As Long, lastRow As Long) As String
    Dim a As String, L As String
    a = ws.Cells(1, c).Address(True, False)
    L = Left(a, InStr(a, "$") - 1)
    RangeText = "Master_Data!$" & L & "$2:$" & L & "$" & lastRow
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not s Is Nothing
End Function

Private Sub DeleteSheet(nm As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
End Sub

Private Function ColIdx(ws As Worksheet, hdr As String) As Long
    Dim f As Range
    Set f = ws.Rows(1).Find(What:=hdr, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then ColIdx = 0 Else ColIdx = f.Column
End Function

