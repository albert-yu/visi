Attribute VB_Name = "Module1"
' Builds a pivot table on Sheet1 of ThisWorkbook from a config encoded as
' plain strings, so it can be invoked from AppleScript's `run VB macro`
' (which Mac Excel's AppleScript dictionary supports for *running* an
' existing macro, even though it has no working path to *create* a
' PivotCache directly via AppleScript -- see fuzz/README.md's "Pivot Table
' Fuzzing" section for how this file fits into fuzz/fuzz_pivot.py).
'
' This is the reference implementation `fuzz_pivot.py`'s VisiPivotDriver and
' win32com driver are meant to match: same PivotCaches.Create /
' CreatePivotTable / PivotFields object model, just reached through VBA
' instead of Python's win32com bindings.
'
' Argument encoding (all strings, since `run VB macro`'s arg1..arg14 are
' most reliably passed/coerced as text across the AppleScript boundary):
'   rowFieldsCSV / colFieldsCSV : "Name:1;Name2:0" (name:subtotalOnOff), or "" for none
'   valueFieldsCSV              : "Name:sum;Name2:count-numbers" (name:aggKeyword), required
'   filterSpec                  : "" for no filter field, else "Name|val1,val2" (values may be empty -> filters out everything)
'   destCell                    : e.g. "H1"
'   grandRowStr / grandColStr   : "1" or "0"
'   sourceIsTableStr            : "1" if sourceRef names an Excel Table (ListObject), "0" if it's a plain range
'   sourceRef                   : table name, or a range like "A1:F31"
Sub BuildFuzzPivot(rowFieldsCSV As String, colFieldsCSV As String, valueFieldsCSV As String, _
                    filterSpec As String, destCell As String, grandRowStr As String, grandColStr As String, _
                    sourceIsTableStr As String, sourceRef As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim srcRange As Range
    If sourceIsTableStr = "1" Then
        Set srcRange = ws.ListObjects(sourceRef).Range
    Else
        Set srcRange = ws.Range(sourceRef)
    End If

    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcRange)

    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range(destCell), TableName:="FuzzPivot")

    ApplyAxisFields pt, rowFieldsCSV, xlRowField
    ApplyAxisFields pt, colFieldsCSV, xlColumnField
    ApplyValueFields pt, valueFieldsCSV
    ApplyFilterField pt, filterSpec

    pt.MergeLabels = False
    pt.HasAutoFormat = False
    pt.ColumnGrand = (grandColStr = "1")
    pt.RowGrand = (grandRowStr = "1")
    pt.RefreshTable

    ThisWorkbook.Save
End Sub

Private Sub ApplyAxisFields(pt As PivotTable, fieldsCSV As String, orientation As XlPivotFieldOrientation)
    If Len(fieldsCSV) = 0 Then Exit Sub
    Dim parts() As String, i As Integer
    parts = Split(fieldsCSV, ";")
    For i = 0 To UBound(parts)
        Dim nv() As String
        nv = Split(parts(i), ":")
        Dim pf As PivotField
        Set pf = pt.PivotFields(nv(0))
        pf.Orientation = orientation
        ' Per-field, not the table-wide PivotTable.RowAxisLayout /
        ' ColumnAxisLayout / SubtotalLocation methods: those hang Mac Excel
        ' outright (not a catchable VBA error -- confirmed by wrapping each
        ' one in its own On Error Resume Next block and still hanging) when
        ' invoked this way, so don't reach for them here even though this
        ' per-field LayoutForm setting is known NOT to actually change the
        ' exported <pivotField>'s `compact` attribute (Excel still shows
        ' "Row Labels"/"Column Labels" captions instead of the field name --
        ' a real, still-open gap, see fuzz/README.md).
        pf.LayoutForm = xlTabular
        pf.LayoutSubtotalLocation = xlAtBottom
        pf.RepeatLabels = False
        If nv(1) = "0" Then
            pf.Subtotals(1) = False ' index 1 = Automatic, the single generic subtotal
        End If
    Next i
End Sub

Private Sub ApplyValueFields(pt As PivotTable, fieldsCSV As String)
    If Len(fieldsCSV) = 0 Then Exit Sub
    Dim parts() As String, i As Integer
    parts = Split(fieldsCSV, ";")
    For i = 0 To UBound(parts)
        Dim nv() As String
        nv = Split(parts(i), ":")
        Dim pf As PivotField
        Set pf = pt.PivotFields(nv(0))
        pf.Orientation = xlDataField
        Select Case nv(1)
            Case "sum": pf.Function = xlSum
            Case "count": pf.Function = xlCount
            Case "count-numbers": pf.Function = xlCountNums
            Case "average": pf.Function = xlAverage
            Case "max": pf.Function = xlMax
            Case "min": pf.Function = xlMin
        End Select
    Next i
End Sub

Private Sub ApplyFilterField(pt As PivotTable, filterSpec As String)
    If Len(filterSpec) = 0 Then Exit Sub
    Dim colName As String, valuesPart As String
    Dim barPos As Integer
    barPos = InStr(filterSpec, "|")
    colName = Left(filterSpec, barPos - 1)
    valuesPart = Mid(filterSpec, barPos + 1)

    Dim values() As String
    values = Split(valuesPart, ",")

    Dim pf As PivotField
    Set pf = pt.PivotFields(colName)
    pf.Orientation = xlPageField

    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        Dim found As Boolean
        found = False
        Dim k As Integer
        If Len(valuesPart) > 0 Then
            For k = 0 To UBound(values)
                If pi.Name = values(k) Then found = True
            Next k
        End If
        pi.Visible = found
    Next pi
End Sub
