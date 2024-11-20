Attribute VB_Name = "Functions"
' Make inputs only take integers (0-9).
Public Sub InputIntegers(ByVal KeyAscii As MSForms.ReturnInteger, ByVal TextBoxContent As String)
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 49 To 57, 61 To 69 ' Numbers 1-9 and Numpad numbers 1-9.
        Case 48, 60 ' Numbers 0 and Numpad number 0.
            If Len(TextBoxContent) = 0 Then
                KeyAscii = 0 ' Disallow leading zero.
            End If
        Case Else
            KeyAscii = 0 ' Disallow all other characters.
    End Select
End Sub

' Avoid copy and pasting (ctrl+v) in text boxes.
Public Sub AvoidCopyPaste(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+v (paste)
        KeyCode = 0
    End If
End Sub
