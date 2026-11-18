Attribute VB_Name = "Module2"
Option Explicit

' Macro 1b: cleans Master_Data (dates, department spelling, helper columns)

Public Sub CleanMasterData()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Dim cMonth As Long, cID As Long, cDept As Long, cLoc As Long, cJoin As Long, cExit As Long
    Dim cKey As Long, cMS As Long
    Dim map As Object, nm As Variant
    Dim arr As Variant, ids As Variant, mons As Variant, keys() As Variant, ms() As Variant
    Dim i As Long, s As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Master_Data")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Run Macro 1 first: Master_Data sheet not found.", vbExclamation: Exit Sub

    cMonth = ColIndex(ws, "Month"): cID = ColIndex(ws, "Employee_ID")
    cDept = ColIndex(ws, "Department"): cLoc = ColIndex(ws, "Location")
    cJoin = ColIndex(ws, "Join_Date"): cExit = ColIndex(ws, "Exit_Date")
    If cMonth * cID * cDept * cLoc * cJoin * cExit = 0 Then
        MsgBox "A required column header was not found in Master_Data.", vbCritical: Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False

    ' 1. Department: standard spelling
    Set map = CreateObject("Scripting.Dictionary")
    For Each nm In Array("IT Services", "Operations", "HR", "Finance", "Sales", "Customer Support", "Admin", "Legal")
        map(UCase(nm)) = nm
    Next nm
    arr = ws.Range(ws.Cells(2, cDept), ws.Cells(lastRow, cDept)).Value
    For i = 1 To UBound(arr, 1)
        s = Trim(CStr(arr(i, 1)))
        If s = "" Then
            arr(i, 1) = Empty
        ElseIf map.Exists(UCase(s)) Then
            arr(i, 1) = map(UCase(s))
        Else
            arr(i, 1) = s
        End If
    Next i
    ws.Range(ws.Cells(2, cDept), ws.Cells(lastRow, cDept)).Value = arr

    ' 2. Location: trim spaces, keep blanks blank
    arr = ws.Range(ws.Cells(2, cLoc), ws.Cells(lastRow, cLoc)).Value
    For i = 1 To UBound(arr, 1)
        s = Trim(CStr(arr(i, 1)))
        If s = "" Then arr(i, 1) = Empty Else arr(i, 1) = s
    Next i
    ws.Range(ws.Cells(2, cLoc), ws.Cells(lastRow, cLoc)).Value = arr

    ' 3. Dates: text dd-mm-yyyy to real dates
    ConvertDateColumn ws, cJoin, lastRow
    ConvertDateColumn ws, cExit, lastRow

    ' 4. Helper columns: Row_Key and Month_Start
    cKey = ColIndex(ws, "Row_Key")
    cMS = ColIndex(ws, "Month_Start")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If cKey = 0 Then lastCol = lastCol + 1: cKey = lastCol
    If cMS = 0 Then lastCol = lastCol + 1: cMS = lastCol
    ws.Cells(1, cKey).Value = "Row_Key"
    ws.Cells(1, cMS).Value = "Month_Start"

    mons = ws.Range(ws.Cells(2, cMonth), ws.Cells(lastRow, cMonth)).Value
    ids = ws.Range(ws.Cells(2, cID), ws.Cells(lastRow, cID)).Value
    ReDim keys(1 To UBound(mons, 1), 1 To 1)
    ReDim ms(1 To UBound(mons, 1), 1 To 1)
    For i = 1 To UBound(mons, 1)
        keys(i, 1) = CStr(mons(i, 1)) & "|" & Trim(CStr(ids(i, 1)))
        ms(i, 1) = MonthStart(mons(i, 1))
    Next i
    ws.Range(ws.Cells(2, cKey), ws.Cells(lastRow, cKey)).Value = keys
    ws.Range(ws.Cells(2, cMS), ws.Cells(lastRow, cMS)).NumberFormat = "mmm-yyyy"
    ws.Range(ws.Cells(2, cMS), ws.Cells(lastRow, cMS)).Value = ms

    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "Master_Data cleaned: " & (lastRow - 1) & " rows.", vbInformation
End Sub

Private Sub ConvertDateColumn(ws As Worksheet, col As Long, lastRow As Long)
    Dim rng As Range, arr As Variant, i As Long, p As Variant, s As String
    Set rng = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
    arr = rng.Value
    For i = 1 To UBound(arr, 1)
        If VarType(arr(i, 1)) = vbDate Then
            ' already a real date
        Else
            s = Trim(CStr(arr(i, 1)))
            If s = "" Then
                arr(i, 1) = Empty
            Else
                p = Split(s, "-")
                If UBound(p) = 2 Then
                    arr(i, 1) = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0)))
                Else
                    arr(i, 1) = Empty
                End If
            End If
        End If
    Next i
    rng.NumberFormat = "dd-mm-yyyy"
    rng.Value = arr
End Sub

Private Function MonthStart(v As Variant) As Variant
    Dim s As String, pos As Long
    If VarType(v) = vbDate Then
        MonthStart = DateSerial(Year(v), Month(v), 1)
        Exit Function
    End If
    s = Trim(CStr(v))
    pos = InStr(1, "JanFebMarAprMayJunJulAugSepOctNovDec", Left(s, 3), vbTextCompare)
    If pos > 0 And Len(s) >= 4 Then
        MonthStart = DateSerial(CLng(Right(s, 4)), (pos + 2) \ 3, 1)
    Else
        MonthStart = Empty
    End If
End Function

Private Function ColIndex(ws As Worksheet, hdr As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(What:=hdr, LookIn:=xlValues, LookAt:=xlWhole)
    If c Is Nothing Then ColIndex = 0 Else ColIndex = c.Column
End Function
