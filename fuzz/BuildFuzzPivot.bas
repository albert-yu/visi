Attribute VB_Name = "Module1"





















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



    pt.ColumnGrand = (grandRowStr = "1")
    pt.RowGrand = (grandColStr = "1")
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









        pf.LayoutForm = xlTabular
        pf.LayoutSubtotalLocation = xlAtBottom
        pf.RepeatLabels = False
        If nv(1) = "0" Then
            pf.Subtotals(1) = False
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
        Dim fn As XlConsolidationFunction
        Select Case nv(1)
            Case "sum": fn = xlSum
            Case "count": fn = xlCount
            Case "count-numbers": fn = xlCountNums
            Case "average": fn = xlAverage
            Case "max": fn = xlMax
            Case "min": fn = xlMin
        End Select






        pt.AddDataField pf, , fn
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









    If Len(valuesPart) = 0 Then Exit Sub

    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        Dim found As Boolean
        found = False
        Dim k As Integer
        For k = 0 To UBound(values)
            If pi.Name = values(k) Then found = True
        Next k
        pi.Visible = found
    Next pi
End Sub
