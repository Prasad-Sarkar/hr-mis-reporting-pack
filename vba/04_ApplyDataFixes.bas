Attribute VB_Name = "Module4"
Option Explicit

' Macro 3: applies fixes to Master_Data using the Errors sheet and employee_master.csv.
' DQ1 duplicate -> row deleted | DQ2/DQ3 -> filled from master | DQ4 -> flagged DQ_Exclude
' DQ5 -> Exit_Date taken from master. Errors sheet is renamed Fix_Log (audit trail).

Public Sub ApplyDataFixes()
    Dim ws As Worksheet, wsE As Worksheet
    Dim f As Variant, master As Object, v As Variant
    Dim cDept As Long, cLoc As Long, cExit As Long, cEx As Long, lastCol As Long
    Dim n As Long, r As Long, mr As Long, id As String, rule As String
    Dim nDup As Long, nDept As Long, nLoc As Long, nAtt As Long, nExit As Long, nMissing As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Master_Data")
    Set wsE = ThisWorkbook.Worksheets("Errors")
    On Error GoTo 0
    If ws Is Nothing Or wsE Is Nothing Then
        MsgBox "Master_Data or Errors sheet not found. Run Macro 2 first.", vbExclamation: Exit Sub
    End If
    If wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "The Errors sheet has no issues to fix.", vbInformation: Exit Sub
    End If

    f = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select employee_master.csv")
    If VarType(f) = vbBoolean Then Exit Sub
    Set master = LoadMaster(CStr(f))
    If master.Count = 0 Then
        MsgBox "Could not read any rows from the CSV. Check that you selected employee_master.csv.", vbCritical
        Exit Sub
    End If

    cDept = ColIdx(ws, "Department"): cLoc = ColIdx(ws, "Location"): cExit = ColIdx(ws, "Exit_Date")
    If cDept * cLoc * cExit = 0 Then MsgBox "A required column is missing in Master_Data.", vbCritical: Exit Sub

    ' Add DQ_Exclude column if missing
    cEx = ColIdx(ws, "DQ_Exclude")
    If cEx = 0 Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        cEx = lastCol + 1
        ws.Cells(1, cEx).Value = "DQ_Exclude"
        ws.Cells(1, cEx).Font.Bold = True
    End If

    Application.ScreenUpdating = False
    n = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row

    ' Pass 1: fix / flag (row numbers are still the original ones)
    For r = 2 To n
        rule = Left(CStr(wsE.Cells(r, 1).Value), 3)
        id = Trim(CStr(wsE.Cells(r, 3).Value))
        mr = CLng(wsE.Cells(r, 4).Value)
        Select Case rule
            Case "DQ1"
                wsE.Cells(r, 7).Value = "Duplicate row removed"
                nDup = nDup + 1
            Case "DQ2"
                If master.Exists(id) Then
                    v = master(id)
                    ws.Cells(mr, cDept).Value = v(0)
                    wsE.Cells(r, 7).Value = "Filled from employee master"
                    nDept = nDept + 1
                Else
                    wsE.Cells(r, 7).Value = "Employee not found in master"
                    nMissing = nMissing + 1
                End If
            Case "DQ3"
                If master.Exists(id) Then
                    v = master(id)
                    ws.Cells(mr, cLoc).Value = v(1)
                    wsE.Cells(r, 7).Value = "Filled from employee master"
                    nLoc = nLoc + 1
                Else
                    wsE.Cells(r, 7).Value = "Employee not found in master"
                    nMissing = nMissing + 1
                End If
            Case "DQ4"
                ws.Cells(mr, cEx).Value = "Y"
                wsE.Cells(r, 7).Value = "Excluded from attendance metrics"
                nAtt = nAtt + 1
            Case "DQ5"
                If master.Exists(id) Then
                    v = master(id)
                    If IsEmpty(v(2)) Then
                        ws.Cells(mr, cExit).ClearContents
                    Else
                        ws.Cells(mr, cExit).NumberFormat = "dd-mm-yyyy"
                        ws.Cells(mr, cExit).Value = v(2)
                    End If
                    wsE.Cells(r, 7).Value = "Exit_Date corrected from employee master"
                    nExit = nExit + 1
                Else
                    wsE.Cells(r, 7).Value = "Employee not found in master"
                    nMissing = nMissing + 1
                End If
        End Select
    Next r

    ' Pass 2: delete duplicate rows from the bottom up so row numbers stay valid
    For r = n To 2 Step -1
        If Left(CStr(wsE.Cells(r, 1).Value), 3) = "DQ1" Then
            ws.Rows(CLng(wsE.Cells(r, 4).Value)).Delete
        End If
    Next r

    ' Keep the log as Fix_Log
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Fix_Log").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    wsE.Name = "Fix_Log"
    wsE.Columns.AutoFit
    ws.Columns.AutoFit

    Application.ScreenUpdating = True
    MsgBox "Fixes applied:" & vbCrLf & _
           nDup & " duplicate rows removed" & vbCrLf & _
           nDept & " Departments filled" & vbCrLf & _
           nLoc & " Locations filled" & vbCrLf & _
           nAtt & " rows flagged DQ_Exclude" & vbCrLf & _
           nExit & " Exit_Dates corrected" & vbCrLf & _
           nMissing & " not found in master" & vbCrLf & vbCrLf & _
           "Master_Data now has " & (ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1) & " rows.", _
           vbInformation, "Fixes complete"
End Sub

' Reads employee_master.csv without Excel date conversion (locale-safe).
' Handles LF, CR and CRLF line endings.
' Returns a dictionary: Employee_ID -> Array(Department, Location, Exit_Date or Empty)
Private Function LoadMaster(path As String) As Object
    Dim d As Object, ff As Integer, s As String, lines() As String
    Dim i As Long, p As Variant, q As Variant, ex As Variant, t As String

    Set d = CreateObject("Scripting.Dictionary")
    ff = FreeFile
    Open path For Binary Access Read As #ff
    If LOF(ff) = 0 Then
        Close #ff
        Set LoadMaster = d
        Exit Function
    End If
    s = Space$(LOF(ff))
    Get #ff, , s
    Close #ff

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    lines = Split(s, vbLf)

    For i = 1 To UBound(lines)            ' row 0 is the header
        If Len(Trim(lines(i))) > 0 Then
            p = Split(lines(i), ",")
            If UBound(p) >= 9 Then
                ex = Empty
                t = Trim(CStr(p(9)))
                If t <> "" Then
                    q = Split(t, "-")
                    If UBound(q) = 2 Then ex = DateSerial(CLng(q(2)), CLng(q(1)), CLng(q(0)))
                End If
                d(Trim(CStr(p(0)))) = Array(Trim(CStr(p(3))), Trim(CStr(p(5))), ex)
            End If
        End If
    Next i
    Set LoadMaster = d
End Function

Private Function ColIdx(ws As Worksheet, hdr As String) As Long
    Dim f As Range
    Set f = ws.Rows(1).Find(What:=hdr, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then ColIdx = 0 Else ColIdx = f.Column
End Function

