Attribute VB_Name = "Module3"
Option Explicit

' Macro 2: validates Master_Data against rules DQ1 to DQ5 and logs every problem.
' It does not change Master_Data.

Private mErr As Worksheet
Private mErrRow As Long
Private mCounts As Object
Private mRules As Variant

Public Sub ValidateMasterData()
    Dim ws As Worksheet, wsSum As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, r As Long, c As Long
    Dim cMonth As Long, cID As Long, cDept As Long, cLoc As Long
    Dim cJoin As Long, cExit As Long, cWork As Long, cPres As Long, cSrc As Long
    Dim data As Variant, seen As Object, months As Object, mKeys As Variant
    Dim m As String, id As String, src As String, expected As Variant
    Dim total As Long, rowTotal As Long, nMonths As Long, issues As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Master_Data")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Master_Data sheet not found. Run Macro 1 and the cleaning macro first.", vbExclamation: Exit Sub

    cMonth = ColIdx(ws, "Month"): cID = ColIdx(ws, "Employee_ID")
    cDept = ColIdx(ws, "Department"): cLoc = ColIdx(ws, "Location")
    cJoin = ColIdx(ws, "Join_Date"): cExit = ColIdx(ws, "Exit_Date")
    cWork = ColIdx(ws, "Working_Days"): cPres = ColIdx(ws, "Days_Present")
    cSrc = ColIdx(ws, "Source_File")
    If cMonth * cID * cDept * cLoc * cJoin * cExit * cWork * cPres * cSrc = 0 Then
        MsgBox "A required column header is missing in Master_Data.", vbCritical: Exit Sub
    End If

    mRules = Array("DQ1 Duplicate employee in month", "DQ2 Department blank", "DQ3 Location blank", _
                   "DQ4 Days_Present > Working_Days", "DQ5 Exit_Date before Join_Date")
    expected = Array(8, 6, 5, 4, 3)   ' expected errors per rule per month

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    DeleteSheet "Errors"
    DeleteSheet "DQ_Summary"
    Application.DisplayAlerts = True

    Set mErr = ThisWorkbook.Worksheets.Add(After:=ws)
    mErr.Name = "Errors"
    mErr.Range("A1:G1").Value = Array("Rule", "Month", "Employee_ID", "Master_Row", "Source_File", "Detail", "Status")
    mErrRow = 1

    Set mCounts = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set months = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value

    For i = 1 To UBound(data, 1)
        m = CStr(data(i, cMonth))
        id = Trim(CStr(data(i, cID)))
        src = CStr(data(i, cSrc))
        If Not months.Exists(m) Then months.Add m, 1

        ' DQ1: same employee twice in the same month
        If seen.Exists(m & "|" & id) Then
            LogIssue 1, m, id, i + 1, src, "Repeated Employee_ID in " & m
        Else
            seen.Add m & "|" & id, 1
        End If

        ' DQ2 / DQ3: blanks
        If Trim(CStr(data(i, cDept))) = "" Then LogIssue 2, m, id, i + 1, src, "Department is blank"
        If Trim(CStr(data(i, cLoc))) = "" Then LogIssue 3, m, id, i + 1, src, "Location is blank"

        ' DQ4: more days present than working days
        If IsNumeric(data(i, cPres)) And IsNumeric(data(i, cWork)) Then
            If Not IsEmpty(data(i, cPres)) And Not IsEmpty(data(i, cWork)) Then
                If CDbl(data(i, cPres)) > CDbl(data(i, cWork)) Then
                    LogIssue 4, m, id, i + 1, src, "Present " & data(i, cPres) & " > Working " & data(i, cWork)
                End If
            End If
        End If

        ' DQ5: exit before join (both must be real dates)
        If VarType(data(i, cJoin)) = vbDate And VarType(data(i, cExit)) = vbDate Then
            If data(i, cExit) < data(i, cJoin) Then
                LogIssue 5, m, id, i + 1, src, "Join " & Format(data(i, cJoin), "dd-mm-yyyy") & ", Exit " & Format(data(i, cExit), "dd-mm-yyyy")
            End If
        End If
    Next i

    issues = mErrRow - 1
    mErr.Rows(1).Font.Bold = True
    mErr.Columns.AutoFit

    ' Summary sheet: rule x month, with expected count and status
    Set wsSum = ThisWorkbook.Worksheets.Add(After:=mErr)
    wsSum.Name = "DQ_Summary"
    mKeys = months.keys
    nMonths = months.Count
    wsSum.Cells(1, 1).Value = "Rule"
    For c = 0 To nMonths - 1
        wsSum.Cells(1, c + 2).Value = mKeys(c)
    Next c
    wsSum.Cells(1, nMonths + 2).Value = "Total found"
    wsSum.Cells(1, nMonths + 3).Value = "Expected"
    wsSum.Cells(1, nMonths + 4).Value = "Status"

    For r = 1 To 5
        wsSum.Cells(r + 1, 1).Value = mRules(r - 1)
        rowTotal = 0
        For c = 0 To nMonths - 1
            total = 0
            If mCounts.Exists(r & "|" & mKeys(c)) Then total = mCounts(r & "|" & mKeys(c))
            wsSum.Cells(r + 1, c + 2).Value = total
            rowTotal = rowTotal + total
        Next c
        wsSum.Cells(r + 1, nMonths + 2).Value = rowTotal
        wsSum.Cells(r + 1, nMonths + 3).Value = expected(r - 1) * nMonths
        If rowTotal = expected(r - 1) * nMonths Then
            wsSum.Cells(r + 1, nMonths + 4).Value = "OK"
        Else
            wsSum.Cells(r + 1, nMonths + 4).Value = "Check"
        End If
    Next r
    wsSum.Rows(1).Font.Bold = True
    wsSum.Columns.AutoFit

    Application.ScreenUpdating = True
    wsSum.Activate
    MsgBox issues & " issues logged in the Errors sheet. See DQ_Summary for counts by rule and month.", vbInformation, "Validation complete"
End Sub

Private Sub LogIssue(ruleIdx As Long, m As String, id As String, rowNum As Long, src As String, detail As String)
    Dim k As String
    mErrRow = mErrRow + 1
    mErr.Cells(mErrRow, 1).Value = mRules(ruleIdx - 1)
    mErr.Cells(mErrRow, 2).Value = m
    mErr.Cells(mErrRow, 3).Value = id
    mErr.Cells(mErrRow, 4).Value = rowNum
    mErr.Cells(mErrRow, 5).Value = src
    mErr.Cells(mErrRow, 6).Value = detail
    mErr.Cells(mErrRow, 7).Value = "Open"
    k = ruleIdx & "|" & m
    mCounts(k) = mCounts(k) + 1
End Sub

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

