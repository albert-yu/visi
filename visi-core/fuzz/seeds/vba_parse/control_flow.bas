Sub S()
    On Error GoTo Failed
    For i = 1 To 10 Step 2
        If i Mod 2 = 0 Then
            Select Case i
            Case 1, 2
            Case 3 To 5
            Case Is >= 6
            Case Else
            End Select
        ElseIf i > 5 Then
            Do While x
            Loop
        Else
            With obj
                .a = .b(1)
            End With
        End If
    Next i
    Exit Sub
Failed:
    Resume Next
End Sub
