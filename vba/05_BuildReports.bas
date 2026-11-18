Attribute VB_Name = "Module5"
Option Explicit

' Macro 4: builds helper columns and report pivots R1 to R5 from Master_Data.
' Run BuildReports. (PrepareHelperColumns and BuildPivotReports can be called separately.)

Public Sub BuildReports()
    If Not PrepareHelperColumns() Then Exit Sub
    BuildPivotReports
End Sub

' ---------- Step 1: helper columns in Master_Data ----------
Public Function PrepareHelperColumns() As Boolean
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, n As Long, i As Long, k As Long
    Dim cMS As Long, cJoin As Long, cExit As Long, cWork As Long, cPres As Long
    Dim cLeave As Long, cUn As Long, cEx As Long
    Dim names As Variant, cols(0 To 7) As Long
    Dim ms As Variant, jn As Variant, ex As Variant, wk As Variant, pr As Variant
    Dim lv As Variant, un As Variant, fl As Variant
    Dim out() As Variant, colArr() As Variant
    Dim msd As Date, eom As Date, hasEx As Boolean, excl As Boolean, bad As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Master_Data")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Master_Data sheet not found.", vbExclamation: Exit Function

    cMS = ColIdx(ws, "Month_Start"): cJoin = ColIdx(ws, "Join_Date"): cExit = ColIdx(ws, "Exit_Date")
    cWork = ColIdx(ws, "Working_Days"): cPres = ColIdx(ws, "Days_Present")
    cLeave = ColIdx(ws, "Leave_Days"): cUn = ColIdx(ws, "Unauthorised_Absence")
    cEx = ColIdx(ws, "DQ_Exclude")
    If cMS * cJoin * cExit * cWork * cPres * cLeave * cUn = 0 Then
        MsgBox "A required column is missing. Run Macro 1, the cleaning macro and Macro 3 first.", vbCritical
        Exit Function
    End If
    hasEx = (cEx > 0)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    n = lastRow - 1
    If n < 2 Then MsgBox "Master_Data has no data.", vbExclamation: Exit Function

    Application.ScreenUpdating = False
    names = Array("Month_Key", "Opening_HC", "Joiners", "Exits", "Att_Working", "Att_Present", "Att_Absent", "Att_Unauth")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For k = 0 To 7
        cols(k) = ColIdx(ws, CStr(names(k)))
        If cols(k) = 0 Then
            lastCol = lastCol + 1
            cols(k) = lastCol
            ws.Cells(1, cols(k)).Value = names(k)
            ws.Cells(1, cols(k)).Font.Bold = True
        End If
    Next k

    ms = ws.Range(ws.Cells(2, cMS), ws.Cells(lastRow, cMS)).Value
    jn = ws.Range(ws.Cells(2, cJoin), ws.Cells(lastRow, cJoin)).Value
    ex = ws.Range(ws.Cells(2, cExit), ws.Cells(lastRow, cExit)).Value
    wk = ws.Range(ws.Cells(2, cWork), ws.Cells(lastRow, cWork)).Value
    pr = ws.Range(ws.Cells(2, cPres), ws.Cells(lastRow, cPres)).Value
    lv = ws.Range(ws.Cells(2, cLeave), ws.Cells(lastRow, cLeave)).Value
    un = ws.Range(ws.Cells(2, cUn), ws.Cells(lastRow, cUn)).Value
    If hasEx Then fl = ws.Range(ws.Cells(2, cEx), ws.Cells(lastRow, cEx)).Value

    ReDim out(1 To n, 1 To 8)
    For i = 1 To n
        For k = 2 To 8: out(i, k) = 0: Next k
        out(i, 1) = ""
        If VarType(ms(i, 1)) = vbDate Then
            msd = ms(i, 1)
            eom = DateSerial(Year(msd), Month(msd) + 1, 0)
            out(i, 1) = Format(msd, "yyyy-mm")
            If VarType(jn(i, 1)) = vbDate Then
                If jn(i, 1) < msd Then
                    out(i, 2) = 1
                ElseIf jn(i, 1) <= eom Then
                    out(i, 3) = 1
                End If
            End If
            If VarType(ex(i, 1)) = vbDate Then
                If ex(i, 1) >= msd And ex(i, 1) <= eom Then out(i, 4) = 1
            End If
            excl = False
            If hasEx Then excl = (UCase(Trim(CStr(fl(i, 1)))) = "Y")
            If Not excl Then
                out(i, 5) = Val(CStr(wk(i, 1)))
                out(i, 6) = Val(CStr(pr(i, 1)))
                out(i, 7) = Val(CStr(lv(i, 1))) + Val(CStr(un(i, 1)))
                out(i, 8) = Val(CStr(un(i, 1)))
            End If
        Else
            bad = bad + 1
        End If
    Next i

    For k = 0 To 7
        ReDim colArr(1 To n, 1 To 1)
        For i = 1 To n: colArr(i, 1) = out(i, k + 1): Next i
        If k = 0 Then ws.Range(ws.Cells(2, cols(k)), ws.Cells(lastRow, cols(k))).NumberFormat = "@"
        ws.Range(ws.Cells(2, cols(k)), ws.Cells(lastRow, cols(k))).Value = colArr
    Next k

    Application.ScreenUpdating = True
    If bad > 0 Then MsgBox bad & " rows have no valid Month_Start. Run the cleaning macro first.", vbExclamation
    PrepareHelperColumns = True
End Function

' ---------- Step 2: report sheets and pivots ----------
Public Sub BuildPivotReports(Optional quiet As Boolean = False)
    Dim ws As Worksheet, wsR As Worksheet, pc As PivotCache, pt As PivotTable
    Dim lastRow As Long, lastCol As Long, r As Long, nm As Variant, nPivots As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Master_Data")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Master_Data sheet not found.", vbExclamation: Exit Sub
    If ColIdx(ws, "Month_Key") = 0 Then MsgBox "Helper columns missing. Run BuildReports.", vbExclamation: Exit Sub

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each nm In Array("R1_Headcount", "R2_JoinersExits", "R3_Attrition", "R4_Attendance", "R5_Leave")
        DeleteSheet CStr(nm)
    Next nm
    Application.DisplayAlerts = True

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
             SourceData:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)))

    ' ----- R1 Headcount -----
    Set wsR = NewReportSheet("R1_Headcount", "R1 Headcount summary", "Closing headcount = opening + new joiners - exits")
    r = 4
    Title wsR, r, "Closing headcount by department"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R1_Dept")
    SetupPT pt, "Department", "Month_Key", Array(Array("Closing_HC", "Closing headcount", "#,##0")), True
    r = NextTitleRow(pt)
    Title wsR, r, "Closing headcount by location"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R1_Loc")
    SetupPT pt, "Location", "Month_Key", Array(Array("Closing_HC", "Closing headcount", "#,##0")), False
    nPivots = 2

    ' ----- R2 Joiners and exits -----
    Set wsR = NewReportSheet("R2_JoinersExits", "R2 Joiners and exits", "Closing headcount of one month equals the opening headcount of the next month")
    r = 4
    Title wsR, r, "Headcount movement by month"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R2_Month")
    SetupPT pt, "Month_Key", "", Array(Array("Opening_HC", "Opening headcount", "#,##0"), _
        Array("Joiners", "New joiners", "#,##0"), Array("Exits", "Employees exited", "#,##0"), _
        Array("Closing_HC", "Closing headcount", "#,##0")), False
    r = NextTitleRow(pt)
    Title wsR, r, "Joiners and exits by department (May to Oct 2026)"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R2_Dept")
    SetupPT pt, "Department", "", Array(Array("Joiners", "New joiners", "#,##0"), _
        Array("Exits", "Employees exited", "#,##0")), False
    nPivots = nPivots + 2

    ' ----- R3 Attrition -----
    Set wsR = NewReportSheet("R3_Attrition", "R3 Attrition report", "Attrition % = exits / average headcount. The Grand Total column is the average monthly attrition.")
    r = 4
    Title wsR, r, "Attrition % by department"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R3_Dept")
    SetupPT pt, "Department", "Month_Key", Array(Array("Attrition_Pct", "Attrition %", "0.0%")), False
    r = NextTitleRow(pt)
    Title wsR, r, "Attrition % by location"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R3_Loc")
    SetupPT pt, "Location", "Month_Key", Array(Array("Attrition_Pct", "Attrition %", "0.0%")), False
    r = NextTitleRow(pt)
    Title wsR, r, "Attrition % by salary band"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R3_Band")
    SetupPT pt, "Salary_Band", "Month_Key", Array(Array("Attrition_Pct", "Attrition %", "0.0%")), False
    nPivots = nPivots + 3

    ' ----- R4 Attendance -----
    Set wsR = NewReportSheet("R4_Attendance", "R4 Attendance and absenteeism", "Rows flagged DQ_Exclude (attendance errors) are left out of every figure on this sheet")
    r = 4
    Title wsR, r, "Attendance % by department"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R4_DeptAtt")
    SetupPT pt, "Department", "Month_Key", Array(Array("Attendance_Pct", "Attendance %", "0.0%")), False
    r = NextTitleRow(pt)
    Title wsR, r, "Attendance % by location"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R4_LocAtt")
    SetupPT pt, "Location", "Month_Key", Array(Array("Attendance_Pct", "Attendance %", "0.0%")), False
    r = NextTitleRow(pt)
    Title wsR, r, "Absenteeism by department (May to Oct 2026)"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R4_Absent")
    SetupPT pt, "Department", "", Array(Array("Absenteeism_Pct", "Absenteeism %", "0.0%"), _
        Array("Unauth_Pct", "Unauthorised absence %", "0.0%")), False
    nPivots = nPivots + 3

    ' ----- R5 Leave -----
    Set wsR = NewReportSheet("R5_Leave", "R5 Leave utilisation", "Leave days taken, by leave type")
    r = 4
    Title wsR, r, "Leave days by leave type and month"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R5_Type")
    SetupPT pt, "Leave_Type", "Month_Key", Array(Array("Leave_Days", "Leave days taken", "#,##0")), False
    HideBlank pt, "Leave_Type"
    r = NextTitleRow(pt)
    Title wsR, r, "Leave days by department and leave type"
    Set pt = NewPT(pc, wsR, r + 1, "pvt_R5_Dept")
    SetupPT pt, "Department", "Leave_Type", Array(Array("Leave_Days", "Leave days taken", "#,##0")), False
    HideBlank pt, "Leave_Type"
    nPivots = nPivots + 2

    ThisWorkbook.Worksheets("R1_Headcount").Activate
    Application.ScreenUpdating = True
    If Not quiet Then MsgBox "Built 5 report sheets with " & nPivots & " pivot tables.", vbInformation, "Reports ready"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Report build failed"
End Sub

' ---------- helpers ----------
Private Function NewPT(pc As PivotCache, ws As Worksheet, r As Long, nm As String) As PivotTable
    Set NewPT = pc.CreatePivotTable(TableDestination:=ws.Cells(r, 1), TableName:=nm)
End Function

Private Sub SetupPT(pt As PivotTable, rowFld As String, colFld As String, items As Variant, addCalc As Boolean)
    Dim it As Variant, df As PivotField
    pt.PivotFields(rowFld).Orientation = xlRowField
    If addCalc Then AddCalcFields pt
    If colFld <> "" Then pt.PivotFields(colFld).Orientation = xlColumnField
    For Each it In items
        Set df = pt.AddDataField(pt.PivotFields(CStr(it(0))), CStr(it(1)), xlSum)
        df.NumberFormat = CStr(it(2))
    Next it
    pt.TableStyle2 = "PivotStyleMedium2"
    pt.HasAutoFormat = False
End Sub

Private Sub AddCalcFields(pt As PivotTable)
    pt.CalculatedFields.Add "Closing_HC", "=Opening_HC+Joiners-Exits", True
    pt.CalculatedFields.Add "Attrition_Pct", "=Exits/(Opening_HC+(Joiners-Exits)/2)", True
    pt.CalculatedFields.Add "Attendance_Pct", "=Att_Present/Att_Working", True
    pt.CalculatedFields.Add "Absenteeism_Pct", "=Att_Absent/Att_Working", True
    pt.CalculatedFields.Add "Unauth_Pct", "=Att_Unauth/Att_Working", True
End Sub

Private Sub HideBlank(pt As PivotTable, fld As String)
    On Error Resume Next
    pt.PivotFields(fld).PivotItems("(blank)").Visible = False
    On Error GoTo 0
End Sub

Private Function NewReportSheet(nm As String, ttl As String, note As String) As Worksheet
    Dim w As Worksheet
    Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    w.Name = nm
    w.Cells(1, 1).Value = ttl
    w.Cells(1, 1).Font.Bold = True
    w.Cells(1, 1).Font.Size = 14
    w.Cells(2, 1).Value = note
    w.Cells(2, 1).Font.Italic = True
    w.Columns(1).ColumnWidth = 26
    w.Range("B:L").ColumnWidth = 14
    Set NewReportSheet = w
End Function

Private Sub Title(ws As Worksheet, r As Long, txt As String)
    ws.Cells(r, 1).Value = txt
    ws.Cells(r, 1).Font.Bold = True
End Sub

Private Function NextTitleRow(pt As PivotTable) As Long
    NextTitleRow = pt.TableRange2.Row + pt.TableRange2.Rows.Count + 2
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

