Attribute VB_Name = "Variant"
Public Function P() As String
    Dim s As String
    Dim i As Long
    For i = 1 To 64
        s = s & i & ":" & One(i) & ";"
    Next i
    P = s
End Function

Private Function One(ByVal i As Long) As String
    On Error GoTo Failed
    Dim v As Variant
    Select Case i
        ' --- literal typing
        Case 1: v = 1
        Case 2: v = 100000
        Case 3: v = 1.5
        Case 4: v = "a"
        Case 5: v = True
        Case 6: v = 1&
        Case 7: v = 1%
        Case 8: v = 1!
        Case 9: v = 1@
        Case 10: v = 1#
        Case 11: v = &HFF
        Case 12: v = &HFFFF&
        Case 13: v = #1/1/2000#
        Case 14: v = Empty
        Case 15: v = 32767
        Case 16: v = 32768
        Case 17: v = 2147483647
        Case 18: v = 2147483648#
        Case 19: v = 1E5
        Case 20: v = 3000000000
        ' --- arithmetic result types
        Case 21: v = 1 + 1
        Case 22: v = 32767 + 1
        Case 23: v = 2147483647 + 1
        Case 24: v = 1 / 2
        Case 25: v = 4 / 2
        Case 26: v = 7 \ 2
        Case 27: v = -7 \ 2
        Case 28: v = 7.6 \ 2
        Case 29: v = 7 Mod 2
        Case 30: v = -7 Mod 2
        Case 31: v = 7.6 Mod 2
        Case 32: v = 2 ^ 2
        Case 33: v = 1.5 + 1
        Case 34: v = 100000 * 100000
        ' --- string coercion
        Case 35: v = "1" + 1
        Case 36: v = "1" + "2"
        Case 37: v = "abc" + 1
        Case 38: v = 1 & 2
        Case 39: v = "1.5" * 2
        Case 40: v = "  3  " + 1
        ' --- boolean
        Case 41: v = True + 1
        Case 42: v = CInt(True)
        Case 43: v = True And False
        Case 44: v = 1 = 1
        Case 45: v = 5 And 3
        Case 46: v = Not 5
        Case 47: v = True + True
        ' --- empty and null
        Case 48: v = Empty + 1
        Case 49: v = Empty & "a"
        Case 50: v = Null + 1
        Case 51: v = Null & "a"
        Case 52: v = IsNull(Null + 1)
        Case 53: v = Empty = 0
        Case 54: v = Empty = ""
        ' --- rounding and conversion
        Case 55: v = CLng(0.5)
        Case 56: v = CLng(1.5)
        Case 57: v = CLng(2.5)
        Case 58: v = CLng(-0.5)
        Case 59: v = CLng(-1.5)
        Case 60: v = Int(-1.5)
        Case 61: v = Fix(-1.5)
        Case 62: v = CInt(32768)
        Case 63: v = CStr(1.5)
        Case 64: v = CDbl("1e3")
    End Select
    One = TypeName(v) & "|" & CStr(v)
    Exit Function
Failed:
    One = "ERR|" & CStr(Err.Number)
End Function
