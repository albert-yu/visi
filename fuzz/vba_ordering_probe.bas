Attribute VB_Name = "Ordering"
' Probe for the error-ordering and Null-handling questions left open after
' Phase 1 of docs/vba-macro-support.md. See docs/vba-error-ordering.md for
' the results and what they imply.
'
' Build and run it the way fuzz/vba_probe.py builds its workbooks:
'
'     import openpyxl, visi_core
'     openpyxl.Workbook().save("base.xlsx")
'     wb = visi_core.Workbook.load("base.xlsx")
'     wb.add_macro("Ordering", open("fuzz/vba_ordering_probe.bas").read())
'     wb.save("ordering.xlsm")
'
' then `run VB macro "P"` against it through AppleScript.
'
' METHODOLOGY NOTE, learned the hard way: the Null section uses IsNull() and
' NOT CStr(). CStr(Null) is itself error 94, so a probe that stringifies the
' result cannot tell "this function raised 94" from "this function returned
' Null and CStr raised 94". A first pass did stringify, reported that every
' intrinsic rejects Null, and was wrong -- most of them propagate it.
Public Function P() As String
    Dim s As String
    Dim i As Long
    For i = 1 To 30
        s = s & i & ":" & One(i) & ";"
    Next i
    P = s
End Function
Private Function One(ByVal i As Long) As String
    On Error GoTo Failed
    Dim v, a, b, n
    n = Null
    Select Case i
        ' --- is 0 / 0 a different error from x / 0 ?
        Case 1: a = 1: b = 0: v = (a / b)
        Case 2: a = 0: b = 0: v = (a / b)
        Case 3: a = False: b = 0: v = (a / b)
        Case 4: a = -1: b = 0: v = (a / b)
        Case 5: a = 1.5: b = 0: v = (a / b)
        Case 6: a = 0: b = 0: v = (a \ b)
        Case 7: a = 0: b = 0: v = (a Mod b)
        Case 8: a = "": b = 0: v = (a / b)
        ' --- which operand is coerced first?
        Case 9: a = "xxxx": b = 0: v = (a / b)
        Case 10: a = 0: b = "xxxx": v = (a / b)
        Case 11: a = "xxxx": b = 0: v = (a - b)
        Case 12: a = 0: b = "xxxx": v = (a - b)
        ' --- does a constant ^ overflow, where a runtime one gives INF?
        Case 13: v = (3.75 ^ 32767)
        Case 14: a = 3.75: v = (a ^ 32767)
        Case 15: v = (255 ^ 255)
        Case 16: a = 255: v = (a ^ 255)
        Case 17: v = ((3.75 ^ 32767) & "x")
        ' --- which intrinsics propagate Null and which reject it?
        '     IsNull, not CStr -- see the methodology note at the top.
        Case 18: v = IsNull(Sgn(n))
        Case 19: v = IsNull(Abs(n))
        Case 20: v = IsNull(Int(n))
        Case 21: v = IsNull(Fix(n))
        Case 22: v = IsNull(Len(n))
        Case 23: v = IsNull(UCase(n))
        Case 24: v = IsNull(Val(n))
        Case 25: v = IsNull(Sqr(n))
        Case 26: v = IsNull(Trim(n))
        Case 27: v = IsNull(Left(n, 1))
        Case 28: v = IsNull(Mid(n, 1, 1))
        Case 29: v = IsNull(InStr(n, "a"))
        Case 30: v = IsNull(Replace(n, "a", "b"))
    End Select
    One = TypeName(v) & "|" & CStr(v)
    Exit Function
Failed:
    One = "ERR|" & CStr(Err.Number)
End Function
