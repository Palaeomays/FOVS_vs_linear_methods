VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AssistantCounting 
   Caption         =   "Counting assistant"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8520.001
   OleObjectBlob   =   "AssistantCounting.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AssistantCounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Variables
'

    ' Input variables
    Dim Now1 As Long
    Dim Now2 As Long
    Dim Now3 As Long
    Dim Now4 As Long
    Dim Now5 As Long
    Dim Now6 As Long
    Dim Now7 As Long
    Dim Now8 As Long
    Dim Now9 As Long
    Dim Now0 As Long
    
    ' Output variables
    Dim Total1 As Long
    Dim Total2 As Long
    Dim Total3 As Long
    Dim Total4 As Long
    Dim Total5 As Long
    Dim Total6 As Long
    Dim Total7 As Long
    Dim Total8 As Long
    Dim Total9 As Long
    Dim Total0 As Long
    
    Dim Y3_1 As Double
    Dim Y3_2 As Double
    Dim Y3_3 As Double
    Dim Y3_4 As Double
    Dim Y3_5 As Double
    Dim Y3_6 As Double
    Dim Y3_7 As Double
    Dim Y3_8 As Double
    Dim Y3_9 As Double
    Dim Y3_0 As Double
    
    Dim SD1 As Double
    Dim SD2 As Double
    Dim SD3 As Double
    Dim SD4 As Double
    Dim SD5 As Double
    Dim SD6 As Double
    Dim SD7 As Double
    Dim SD8 As Double
    Dim SD9 As Double
    Dim SD0 As Double
    
    ' Create a new worksheet named "Counting (Exhaustive)"
    Dim SavedVariablesCountingRunning As Worksheet
    Dim SavedVariablesCountingRunningExists As Boolean
    
    ' Create a new worksheet named "Counting (Summary)"
    Dim SavedVariablesCountingEnd As Worksheet
    Dim SavedVariablesCountingEndExists As Boolean
    
    ' Extra variables
    Dim CurrentFOV As Long


Private Sub ToggleButton_Hotkeys_Change()

End Sub

'
' Startup
'

Private Sub UserForm_Initialize()
    If Len(SampleName) > 1 Then
        txt_SampleName.Text = SampleName
    Else
    End If
    
    If OriginStarter Or OriginLinear Or OriginFOVSTarget Or OriginFOVSMarker Or OriginCountingEffort Or OriginCalibrationFOV Then
        MsgBox "The counting assistant is designed to enable up to nine target and one marker specimen categories to be counted concurrently." & vbNewLine & vbNewLine & "First: label the specimen categories you wish to count." & vbNewLine & "Second: perform a count of the first field of view." & vbNewLine & "Third: Press the 'next FOV' button when you transition to a new field of view." & vbNewLine & vbNewLine & "The hotkeys enable the rapid counting of multiple specimen categories using the 0–9 keys on a keyboard or numpad." & vbNewLine & vbNewLine & "The optional timer enables the automatic calculation of data collection effort and the determination of the most efficient count method." & vbNewLine & vbNewLine & "IMPORTANT: for valid statistics, please ensure that you include all specimen categories before counting.", vbInformation
    
        ' Check if certain sheets are present. Iterate through all worksheets in the workbook.
        
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = "Counting (Summary)" Then
                ' Set the flag to True if the worksheet exists
                SavedVariablesCountingEndExists = True
                Exit For
            End If
        Next ws
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = "Counting (Exhaustive)" Then
                ' Set the flag to True if the worksheet exists
                SavedVariablesCountingRunningExists = True
                Exit For
            End If
        Next ws
    End If
End Sub

'
' Shutdown
'

Private Sub UserForm_Terminate()
    OriginStarter = False
    OriginLinear = False
    OriginFOVSTarget = False
    OriginFOVSMarker = False
    OriginCountingEffort = False
    OriginCalibrationFOV = False
    
    If SavedVariablesCountingRunningExists Or SavedVariablesCountingEndExists Then
        response = MsgBox("Would you like to retain the counting worksheets 'Counting (Exhaustive)' and 'Counting (Summary)'?", vbQuestion + vbYesNo, "Save Counting?")
        ' Check user response
        If response = vbYes Then
        ' Do nothing.
            Unload Me
        Else
            Application.DisplayAlerts = False ' Suppress the confirmation dialog
           
             ' Delete sheet if it exists
            For Each ws In ThisWorkbook.Worksheets
                If SavedVariablesCountingRunningExists And ws.Name = "Counting (Exhaustive)" Then
                    ws.Delete
                    SavedVariablesCountingRunningExists = False
                End If
            Next ws
            For Each ws In ThisWorkbook.Worksheets
                If SavedVariablesCountingEndExists And ws.Name = "Counting (Summary)" Then
                    ws.Delete
                    SavedVariablesCountingEndExists = False
                End If
            Next ws
          
            Application.DisplayAlerts = True ' Re-enable confirmation dialogs
            On Error GoTo 0 ' Reset error handling to default
    
            CountingSaved = False
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

'
' Inputs
'

Private Sub txt_Now1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name1.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("Please provide a name of this specimen category first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name2.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name3.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name4.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name5.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name6.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name7.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name8.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name9.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_Now0_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
    Select Case KeyAscii
        Case 8 ' Backspace.
        Case 48 To 57, 96 To 105 ' Numbers 0-9 and Numpad numbers 0-9.
        If Len(txt_Name0.Text) > 0 Then
        ' Allow input if the textbox is not empty.
        Else
            KeyAscii = 0 ' Disallow input if the textbox is empty.
            response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
        End If
        Case Else
            KeyAscii = 0 ' Disallow other characters.
    End Select
End Sub

Private Sub txt_TimeTotal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_TimeTotal.Text) > 0 Then
                ' Allow input if the textbox is not empty
                ' Do nothing, allow input
            Else
                ' Disallow input if the textbox is empty
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0 ' Disallow other characters
    End Select
End Sub

' Avoid pasting words and numbers.

Private Sub txt_Now1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_Now0_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_TimeTotal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

'
' Hotkeys
'
Private Sub ToggleButton_Hotkeys_Click()
    If ToggleButton_Hotkeys.Value = True Then
        ToggleButton_Hotkeys.Caption = "Deactivate hotkeys"
        ToggleButton_Hotkeys.BackColor = RGB(212, 236, 214) ' Greenish color
    Else
        ToggleButton_Hotkeys.Caption = "Activate hotkeys"
        ToggleButton_Hotkeys.BackColor = &H8000000F  ' Grey color
    End If
End Sub


' #1
Private Sub ToggleButton_Hotkeys_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If ToggleButton_Hotkeys.Value = True Then
        If KeyCode = 49 Or KeyCode = 97 Then ' Check if the pressed key is the number "1"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now1
                If Len(txt_Name1.Text) > 0 Then
                    If txt_Now1.Value = "" Then
                        txt_Now1.Value = 0
                    ElseIf txt_Now1.Value <> 0 Then
                        txt_Now1.Value = txt_Now1.Value - 1
                    ElseIf txt_Now1.Value < 0 Then
                        txt_Now1.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name1.Text) > 0 Then
                    If txt_Now1.Value = "" Then
                        txt_Now1.Value = 1
                    Else
                        txt_Now1.Value = txt_Now1.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 50 Or KeyCode = 98 Then ' Check if the pressed key is the number "2"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now2
                If Len(txt_Name2.Text) > 0 Then
                    If txt_Now2.Value = "" Then
                        txt_Now2.Value = 0
                    ElseIf txt_Now2.Value <> 0 Then
                        txt_Now2.Value = txt_Now2.Value - 1
                    ElseIf txt_Now2.Value < 0 Then
                        txt_Now2.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name2.Text) > 0 Then
                    If txt_Now2.Value = "" Then
                        txt_Now2.Value = 1
                    Else
                        txt_Now2.Value = txt_Now2.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 51 Or KeyCode = 99 Then ' Check if the pressed key is the number "3"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now3
                If Len(txt_Name3.Text) > 0 Then
                    If txt_Now3.Value = "" Then
                        txt_Now3.Value = 0
                    ElseIf txt_Now3.Value <> 0 Then
                        txt_Now3.Value = txt_Now3.Value - 1
                    ElseIf txt_Now3.Value < 0 Then
                        txt_Now3.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name3.Text) > 0 Then
                    If txt_Now3.Value = "" Then
                        txt_Now3.Value = 1
                    Else
                        txt_Now3.Value = txt_Now3.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 52 Or KeyCode = 100 Then ' Check if the pressed key is the number "4"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now4
                If Len(txt_Name4.Text) > 0 Then
                    If txt_Now4.Value = "" Then
                        txt_Now4.Value = 0
                    ElseIf txt_Now4.Value <> 0 Then
                        txt_Now4.Value = txt_Now4.Value - 1
                    ElseIf txt_Now4.Value < 0 Then
                        txt_Now4.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name4.Text) > 0 Then
                    If txt_Now4.Value = "" Then
                        txt_Now4.Value = 1
                    Else
                        txt_Now4.Value = txt_Now4.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 53 Or KeyCode = 101 Then ' Check if the pressed key is the number "5"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now5
                If Len(txt_Name5.Text) > 0 Then
                    If txt_Now5.Value = "" Then
                        txt_Now5.Value = 0
                    ElseIf txt_Now5.Value <> 0 Then
                        txt_Now5.Value = txt_Now5.Value - 1
                    ElseIf txt_Now5.Value < 0 Then
                        txt_Now5.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name5.Text) > 0 Then
                    If txt_Now5.Value = "" Then
                        txt_Now5.Value = 1
                    Else
                        txt_Now5.Value = txt_Now5.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 54 Or KeyCode = 102 Then ' Check if the pressed key is the number "6"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now6
                If Len(txt_Name6.Text) > 0 Then
                    If txt_Now6.Value = "" Then
                        txt_Now6.Value = 0
                    ElseIf txt_Now6.Value <> 0 Then
                        txt_Now6.Value = txt_Now6.Value - 1
                    ElseIf txt_Now6.Value < 0 Then
                        txt_Now6.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name6.Text) > 0 Then
                    If txt_Now6.Value = "" Then
                        txt_Now6.Value = 1
                    Else
                        txt_Now6.Value = txt_Now6.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 55 Or KeyCode = 103 Then ' Check if the pressed key is the number "7"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now7
                If Len(txt_Name7.Text) > 0 Then
                    If txt_Now7.Value = "" Then
                        txt_Now7.Value = 0
                    ElseIf txt_Now7.Value <> 0 Then
                        txt_Now7.Value = txt_Now7.Value - 1
                    ElseIf txt_Now7.Value < 0 Then
                        txt_Now7.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name7.Text) > 0 Then
                    If txt_Now7.Value = "" Then
                        txt_Now7.Value = 1
                    Else
                        txt_Now7.Value = txt_Now7.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 56 Or KeyCode = 104 Then ' Check if the pressed key is the number "8"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now8
                If Len(txt_Name8.Text) > 0 Then
                    If txt_Now8.Value = "" Then
                        txt_Now8.Value = 0
                    ElseIf txt_Now8.Value <> 0 Then
                        txt_Now8.Value = txt_Now8.Value - 1
                    ElseIf txt_Now8.Value < 0 Then
                        txt_Now8.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name8.Text) > 0 Then
                    If txt_Now8.Value = "" Then
                        txt_Now8.Value = 1
                    Else
                        txt_Now8.Value = txt_Now8.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 57 Or KeyCode = 105 Then ' Check if the pressed key is the number "9"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now9
                If Len(txt_Name9.Text) > 0 Then
                    If txt_Now9.Value = "" Then
                        txt_Now9.Value = 0
                    ElseIf txt_Now9.Value <> 0 Then
                        txt_Now9.Value = txt_Now9.Value - 1
                    ElseIf txt_Now9.Value < 0 Then
                        txt_Now9.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name9.Text) > 0 Then
                    If txt_Now9.Value = "" Then
                        txt_Now9.Value = 1
                    Else
                        txt_Now9.Value = txt_Now9.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 48 Or KeyCode = 96 Then ' Check if the pressed key is the number "0"
            If Shift = 1 Then ' Check if Shift is also pressed
                ' Subtract 1 from the value of txt_Now0
                If Len(txt_Name0.Text) > 0 Then
                    If txt_Now0.Value = "" Then
                        txt_Now0.Value = 0
                    ElseIf txt_Now0.Value <> 0 Then
                        txt_Now0.Value = txt_Now0.Value - 1
                    ElseIf txt_Now0.Value < 0 Then
                        txt_Now0.Value = 0
                    Else
                    End If
                Else
                End If
            Else
                If Len(txt_Name0.Text) > 0 Then
                    If txt_Now0.Value = "" Then
                        txt_Now0.Value = 1
                    Else
                        txt_Now0.Value = txt_Now0.Value + 1
                    End If
                Else
                End If
            End If
        ElseIf KeyCode = 32 Then ' Check if the pressed key is the Space bar.
            ToggleButton_Hotkeys.Value = False
            CommandButton_Next_Click
            ToggleButton_Hotkeys_Click
        End If
    End If
End Sub

Private Sub ToggleButton_Hotkeys_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Set the value of the toggle button to False when it loses focus
    ToggleButton_Hotkeys.Value = False
    ToggleButton_Hotkeys.Caption = "Activate shortcuts"
    ToggleButton_Hotkeys.BackColor = &H8000000F  ' Grey color
End Sub



'
' SpinButtons
'

Private Sub SpinButton_1_SpinUp()
    If Len(txt_Name1.Text) > 0 Then
        If txt_Now1.Value = "" Then
            txt_Now1.Value = 1
        Else
            txt_Now1.Value = txt_Now1.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_1_SpinDown()
    If Len(txt_Name1.Text) > 0 Then
        If txt_Now1.Value = "" Then
            txt_Now1.Value = 0
        ElseIf txt_Now1.Value <> 0 Then
            txt_Now1.Value = txt_Now1.Value - 1
        ElseIf txt_Now1.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now1.Value = 0
        Else
            txt_Now1.Value = txt_Now1.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_2_SpinUp()
    If Len(txt_Name2.Text) > 0 Then
        If txt_Now2.Value = "" Then
            txt_Now2.Value = 1
        Else
            txt_Now2.Value = txt_Now2.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_2_SpinDown()
    If Len(txt_Name2.Text) > 0 Then
        If txt_Now2.Value = "" Then
            txt_Now2.Value = 0
        ElseIf txt_Now2.Value <> 0 Then
            txt_Now2.Value = txt_Now2.Value - 1
        ElseIf txt_Now1.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now2.Value = 0
        Else
            txt_Now2.Value = txt_Now2.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_3_SpinUp()
    If Len(txt_Name3.Text) > 0 Then
        If txt_Now3.Value = "" Then
            txt_Now3.Value = 1
        Else
            txt_Now3.Value = txt_Now3.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_3_SpinDown()
    If Len(txt_Name3.Text) > 0 Then
        If txt_Now3.Value = "" Then
            txt_Now3.Value = 0
        ElseIf txt_Now3.Value <> 0 Then
            txt_Now3.Value = txt_Now3.Value - 1
        ElseIf txt_Now3.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now3.Value = 0
        Else
            txt_Now3.Value = txt_Now3.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_4_SpinUp()
    If Len(txt_Name4.Text) > 0 Then
        If txt_Now4.Value = "" Then
            txt_Now4.Value = 1
        Else
            txt_Now4.Value = txt_Now4.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_4_SpinDown()
    If Len(txt_Name4.Text) > 0 Then
        If txt_Now4.Value = "" Then
            txt_Now4.Value = 0
        ElseIf txt_Now4.Value <> 0 Then
            txt_Now4.Value = txt_Now4.Value - 1
        ElseIf txt_Now4.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now4.Value = 0
        Else
            txt_Now4.Value = txt_Now4.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_5_SpinUp()
    If Len(txt_Name5.Text) > 0 Then
        If txt_Now5.Value = "" Then
            txt_Now5.Value = 1
        Else
            txt_Now5.Value = txt_Now5.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_5_SpinDown()
    If Len(txt_Name5.Text) > 0 Then
        If txt_Now5.Value = "" Then
            txt_Now5.Value = 0
        ElseIf txt_Now5.Value <> 0 Then
            txt_Now5.Value = txt_Now5.Value - 1
        ElseIf txt_Now5.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now5.Value = 0
        Else
            txt_Now5.Value = txt_Now5.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_6_SpinUp()
    If Len(txt_Name6.Text) > 0 Then
        If txt_Now6.Value = "" Then
            txt_Now6.Value = 1
        Else
            txt_Now6.Value = txt_Now6.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_6_SpinDown()
    If Len(txt_Name6.Text) > 0 Then
        If txt_Now6.Value = "" Then
            txt_Now6.Value = 0
        ElseIf txt_Now6.Value <> 0 Then
            txt_Now6.Value = txt_Now6.Value - 1
        ElseIf txt_Now6.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now6.Value = 0
        Else
            txt_Now6.Value = txt_Now6.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_7_SpinUp()
    If Len(txt_Name7.Text) > 0 Then
        If txt_Now7.Value = "" Then
            txt_Now7.Value = 1
        Else
            txt_Now7.Value = txt_Now7.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_7_SpinDown()
    If Len(txt_Name7.Text) > 0 Then
        If txt_Now7.Value = "" Then
            txt_Now7.Value = 0
        ElseIf txt_Now7.Value <> 0 Then
            txt_Now7.Value = txt_Now7.Value - 1
        ElseIf txt_Now7.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now7.Value = 0
        Else
            txt_Now7.Value = txt_Now7.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_8_SpinUp()
    If Len(txt_Name8.Text) > 0 Then
        If txt_Now8.Value = "" Then
            txt_Now8.Value = 1
        Else
            txt_Now8.Value = txt_Now8.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_8_SpinDown()
    If Len(txt_Name8.Text) > 0 Then
        If txt_Now8.Value = "" Then
            txt_Now8.Value = 0
        ElseIf txt_Now8.Value <> 0 Then
            txt_Now8.Value = txt_Now8.Value - 1
        ElseIf txt_Now8.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now8.Value = 0
        Else
            txt_Now8.Value = txt_Now8.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_9_SpinUp()
    If Len(txt_Name9.Text) > 0 Then
        If txt_Now9.Value = "" Then
            txt_Now9.Value = 1
        Else
            txt_Now9.Value = txt_Now9.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_9_SpinDown()
    If Len(txt_Name9.Text) > 0 Then
        If txt_Now9.Value = "" Then
            txt_Now9.Value = 0
        ElseIf txt_Now9.Value <> 0 Then
            txt_Now9.Value = txt_Now9.Value - 1
        ElseIf txt_Now9.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now9.Value = 0
        Else
            txt_Now9.Value = txt_Now9.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_0_SpinUp()
    If Len(txt_Name0.Text) > 0 Then
        If txt_Now0.Value = "" Then
            txt_Now0.Value = 1
        Else
            txt_Now0.Value = txt_Now0.Value + 1
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

Private Sub SpinButton_0_SpinDown()
    If Len(txt_Name0.Text) > 0 Then
        If txt_Now0.Value = "" Then
            txt_Now0.Value = 0
        ElseIf txt_Now0.Value <> 0 Then
            txt_Now0.Value = txt_Now0.Value - 1
        ElseIf txt_Now0.Value < 0 Then ' Check if the resulting value is less than 0, set it to 0 instead.
            txt_Now0.Value = 0
        Else
            txt_Now0.Value = txt_Now0.Value
        End If
    Else
    response = MsgBox("To allow counting please name the variable first.", vbInformation, "Unlabeled Variables")
    End If
End Sub

'
' Outputs
'

Private Sub CommandButton_Next_Click()
  
    If Len(txt_Name1.Value) > 0 And Not IsNumeric(txt_Now1.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #1.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name2.Value) > 0 And Not IsNumeric(txt_Now2.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #2.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name3.Value) > 0 And Not IsNumeric(txt_Now3.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #3.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name4.Value) > 0 And Not IsNumeric(txt_Now4.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #4.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name5.Value) > 0 And Not IsNumeric(txt_Now5.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #5.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name6.Value) > 0 And Not IsNumeric(txt_Now6.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #6.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name7.Value) > 0 And Not IsNumeric(txt_Now7.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #7.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name8.Value) > 0 And Not IsNumeric(txt_Now8.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #8.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name9.Value) > 0 And Not IsNumeric(txt_Now9.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #7.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Len(txt_Name0.Value) > 0 And Not IsNumeric(txt_Now0.Value) Then ' Validate input field to see if not empty, if not, then ask for count number.
        MsgBox "Please enter the amount of counted specimens in #0.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
'    If Not IsNumeric(txt_TimeTotal.Value) Then
'        MsgBox "If you wish to measure your sampling effort or test for the most efficient method, please insert the time taken for your count", "Input Required"
'        Exit Sub
'    End If
    
    ' Store input values in memory as longs.
    
    If Len(txt_SampleName.Text) > 1 Then
        SampleName = txt_SampleName.Text
    Else
    End If

    If Not txt_Now1.Value = "" Then
        Now1 = CLng(txt_Now1.Value)
    Else
    End If
    
    If Not txt_Now2.Value = "" Then
        Now2 = CLng(txt_Now2.Value)
    Else
    End If
    
    If Not txt_Now3.Value = "" Then
        Now3 = CLng(txt_Now3.Value)
    Else
    End If
    
    If Not txt_Now4.Value = "" Then
        Now4 = CLng(txt_Now4.Value)
    Else
    End If
    
    If Not txt_Now5.Value = "" Then
        Now5 = CLng(txt_Now5.Value)
    Else
    End If
    
    If Not txt_Now6.Value = "" Then
        Now6 = CLng(txt_Now6.Value)
    Else
    End If
    
    If Not txt_Now7.Value = "" Then
        Now7 = CLng(txt_Now7.Value)
    Else
    End If
    
    If Not txt_Now8.Value = "" Then
        Now8 = CLng(txt_Now8.Value)
    Else
    End If
    
    If Not txt_Now9.Value = "" Then
        Now9 = CLng(txt_Now9.Value)
    Else
    End If
    
    If Not txt_Now0.Value = "" Then
        Now0 = CLng(txt_Now0.Value)
    Else
    End If
    
    CurrentFOV = CLng(txt_CurrentFOV.Value)
    
    If Not txt_NumFOV.Value = "" Then
        NumFOV = CLng(txt_NumFOV.Value)
    Else
    End If
    
    ' Save values in spreadsheet

    ' Check if the sheet "Counting (Exhaustive)" already exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Counting (Exhaustive)" Then
            SavedVariablesCountingRunningExists = True
            Set SavedVariablesCountingRunning = ws
            Exit For
        End If
    Next ws

    ' If the sheet doesn't exist, create a new one
    If Not SavedVariablesCountingRunningExists Then
        Set SavedVariablesCountingRunning = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Calculator"))
        SavedVariablesCountingRunning.Name = "Counting (Exhaustive)"
    End If

    ' Check if the SavedVariablesCountingRunning object is initialized
    ' Check if the headers have already been filled
    If SavedVariablesCountingRunning.Cells(1, 1).Value = "" Then ' Headers have not been filled yet, so fill them
        With SavedVariablesCountingRunning
            .Cells(1, 1).Value = "Date and time"
            .Cells(1, 2).Value = "Sample name"
            .Cells(1, 3).Value = "Field of view number"
            .Cells(1, 4).Value = "Specimen name of #0"
            .Cells(1, 5).Value = "Number of #0"
            .Cells(1, 6).Value = "Specimen name of #1"
            .Cells(1, 7).Value = "Number of #1"
            .Cells(1, 8).Value = "Specimen name of #2"
            .Cells(1, 9).Value = "Number of #2"
            .Cells(1, 10).Value = "Specimen name of #3"
            .Cells(1, 11).Value = "Number of #3"
            .Cells(1, 12).Value = "Specimen name of #4"
            .Cells(1, 13).Value = "Number of #4"
            .Cells(1, 14).Value = "Specimen name of #5"
            .Cells(1, 15).Value = "Number of #5"
            .Cells(1, 16).Value = "Specimen name of #6"
            .Cells(1, 17).Value = "Number of #6"
            .Cells(1, 18).Value = "Specimen name of #7"
            .Cells(1, 19).Value = "Number of #7"
            .Cells(1, 20).Value = "Specimen name of #8"
            .Cells(1, 21).Value = "Number of #8"
            .Cells(1, 22).Value = "Specimen name of #9"
            .Cells(1, 23).Value = "Number of #9"
            .Cells(1, 24).Value = "Sum of target specimens"
        End With
    End If
    
    ' Find the next available row in column A
    Dim nextRow As Long
    nextRow = SavedVariablesCountingRunning.Cells(SavedVariablesCountingRunning.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Write the current date and time to the first column (Column A) in the next available row
    SavedVariablesCountingRunning.Cells(nextRow, "A").Value = Now
    
    ' Write values from the userform to specific cells in the next available row
    SavedVariablesCountingRunning.Cells(nextRow, "B").Value = txt_SampleName.Text
    SavedVariablesCountingRunning.Cells(nextRow, "C").Value = txt_CurrentFOV.Value
    SavedVariablesCountingRunning.Cells(nextRow, "D").Value = txt_Name0.Text
    SavedVariablesCountingRunning.Cells(nextRow, "E").Value = Now0
    SavedVariablesCountingRunning.Cells(nextRow, "F").Value = txt_Name1.Text
    SavedVariablesCountingRunning.Cells(nextRow, "G").Value = Now1
    SavedVariablesCountingRunning.Cells(nextRow, "H").Value = txt_Name2.Text
    SavedVariablesCountingRunning.Cells(nextRow, "I").Value = Now2
    SavedVariablesCountingRunning.Cells(nextRow, "J").Value = txt_Name3.Text
    SavedVariablesCountingRunning.Cells(nextRow, "K").Value = Now3
    SavedVariablesCountingRunning.Cells(nextRow, "L").Value = txt_Name4.Text
    SavedVariablesCountingRunning.Cells(nextRow, "M").Value = Now4
    SavedVariablesCountingRunning.Cells(nextRow, "N").Value = txt_Name5.Text
    SavedVariablesCountingRunning.Cells(nextRow, "O").Value = Now5
    SavedVariablesCountingRunning.Cells(nextRow, "P").Value = txt_Name6.Text
    SavedVariablesCountingRunning.Cells(nextRow, "Q").Value = Now6
    SavedVariablesCountingRunning.Cells(nextRow, "R").Value = txt_Name7.Text
    SavedVariablesCountingRunning.Cells(nextRow, "S").Value = Now7
    SavedVariablesCountingRunning.Cells(nextRow, "T").Value = txt_Name8.Text
    SavedVariablesCountingRunning.Cells(nextRow, "U").Value = Now8
    SavedVariablesCountingRunning.Cells(nextRow, "V").Value = txt_Name9.Text
    SavedVariablesCountingRunning.Cells(nextRow, "W").Value = Now9
    SavedVariablesCountingRunning.Cells(nextRow, "X").Value = Now1 + Now2 + Now3 + Now4 + Now5 + Now6 + Now7 + Now8 + Now9
    
    ' Reset current counts to 0.
    
    If Len(txt_Name1.Text) > 0 Then
        txt_Now1.Text = 0
    Else
    End If
    
    If Len(txt_Name2.Text) > 0 Then
        txt_Now2.Text = 0
    Else
    End If
    
    If Len(txt_Name3.Text) > 0 Then
        txt_Now3.Text = 0
    Else
    End If
    
    If Len(txt_Name4.Text) > 0 Then
        txt_Now4.Text = 0
    Else
    End If
    
    If Len(txt_Name5.Text) > 0 Then
        txt_Now5.Text = 0
    Else
    End If
    
    If Len(txt_Name6.Text) > 0 Then
        txt_Now6.Text = 0
    Else
    End If
    
    If Len(txt_Name7.Text) > 0 Then
        txt_Now7.Text = 0
    Else
    End If
    
    If Len(txt_Name8.Text) > 0 Then
        txt_Now8.Text = 0
    Else
    End If
    
    If Len(txt_Name9.Text) > 0 Then
        txt_Now9.Text = 0
    Else
    End If
    
    If Len(txt_Name0.Text) > 0 Then
        txt_Now0.Text = 0
    Else
    End If
    
    ' Increase current field-of-view by 1.
    txt_CurrentFOV.Value = txt_CurrentFOV.Value + 1
    
    ' Increase total field-of-views by 1.
    If txt_NumFOV.Value = "" Then
        txt_NumFOV.Value = 1
    Else
        txt_NumFOV.Value = txt_NumFOV.Value + 1
    End If
    
    '
    '
    '
    
    ' Check if the sheet "Counting (Summary)" already exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Counting (Summary)" Then
            SavedVariablesCountingEndExists = True
            Set SavedVariablesCountingEnd = ws
            Exit For
        End If
    Next ws

    ' If the sheet doesn't exist, create a new one
    If Not SavedVariablesCountingEndExists Then
        Set SavedVariablesCountingEnd = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Counting (Exhaustive)"))
        SavedVariablesCountingEnd.Name = "Counting (Summary)"
    End If

    ' Check if the SavedVariablesCountingEnd object is initialized
    ' Check if the headers have already been filled
    If SavedVariablesCountingEnd.Cells(1, 1).Value = "" Then ' Headers have not been filled yet, so fill them
        With SavedVariablesCountingEnd
            .Cells(1, 1).Value = "Date and time"
            .Cells(1, 2).Value = "Sample name"
            .Cells(1, 3).Value = "Field of view number"
            .Cells(1, 4).Value = "Specimen name of #0"
            .Cells(1, 5).Value = "Total number of #0"
            .Cells(1, 6).Value = "Mean number per FOV [Y3] of #0"
            .Cells(1, 7).Value = "Standard deviation [S3] of #0"
            .Cells(1, 8).Value = "Specimen name of #1"
            .Cells(1, 9).Value = "Total number of #1"
            .Cells(1, 10).Value = "Mean number per FOV [Y3] of #1"
            .Cells(1, 11).Value = "Standard deviation [S3] of #1"
            .Cells(1, 12).Value = "Specimen name #2"
            .Cells(1, 13).Value = "Total number of #2"
            .Cells(1, 14).Value = "Mean number per FOV [Y3] of #2"
            .Cells(1, 15).Value = "Standard deviation [S3] of #2"
            .Cells(1, 16).Value = "Specimen name #3"
            .Cells(1, 17).Value = "Total number of #3"
            .Cells(1, 18).Value = "Mean number per FOV [Y3] of #3"
            .Cells(1, 19).Value = "Standard deviation [S3] of #3"
            .Cells(1, 20).Value = "Specimen name #4"
            .Cells(1, 21).Value = "Total number of #4"
            .Cells(1, 22).Value = "Mean number per FOV [Y3] of #4"
            .Cells(1, 23).Value = "Standard deviation [S3] of #4"
            .Cells(1, 24).Value = "Specimen Name #5"
            .Cells(1, 25).Value = "Total number of #5"
            .Cells(1, 26).Value = "Mean number per FOV [Y3] of #5"
            .Cells(1, 27).Value = "Standard deviation [S3] of #5"
            .Cells(1, 28).Value = "Specimen Name #6"
            .Cells(1, 29).Value = "Total number of #6"
            .Cells(1, 30).Value = "Mean number per FOV [Y3] of #6"
            .Cells(1, 31).Value = "Standard deviation [S3] of #6"
            .Cells(1, 32).Value = "Specimen Name #7"
            .Cells(1, 33).Value = "Total number of #7"
            .Cells(1, 34).Value = "Mean number per FOV [Y3] of #7"
            .Cells(1, 35).Value = "Standard deviation [S3] of #7"
            .Cells(1, 36).Value = "Specimen Name #8"
            .Cells(1, 37).Value = "Total number of #8"
            .Cells(1, 38).Value = "Mean number per FOV [Y3] of #8"
            .Cells(1, 39).Value = "Standard deviation [S3] of #8"
            .Cells(1, 40).Value = "Specimen Name #9"
            .Cells(1, 41).Value = "Total number of #9"
            .Cells(1, 42).Value = "Mean number per FOV [Y3] of #9"
            .Cells(1, 43).Value = "Standard deviation [S3] of #9"
            .Cells(1, 44).Value = "Total of sum of target specimens"
            .Cells(1, 45).Value = "Standard deviation [S3] of sum of target specimens"
        End With
    End If
    
    ' Only update row, TODO: for now.
    Dim dataRow As Long
    dataRow = 2
    
    ' Write the current date and time to the first column (Column A) in the next available row
    SavedVariablesCountingEnd.Cells(dataRow, "A").Value = Now
    
    ' Write values from the userform to specific cells in the next available row
    SavedVariablesCountingEnd.Cells(dataRow, "B").Value = txt_SampleName.Text
    SavedVariablesCountingEnd.Cells(dataRow, "C").Value = txt_NumFOV.Value
    
    If Not txt_Now0.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "D").Value = txt_Name0.Text
    SavedVariablesCountingEnd.Cells(dataRow, "E").Formula = "=SUM('Counting (Exhaustive)'!E2:E1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "F").Formula = "=SUM('Counting (Exhaustive)'!E2:E1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "G").Formula = "=STDEV.S('Counting (Exhaustive)'!E2:E1048576)"
    Else
    End If

    If Not txt_Now1.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "H").Value = txt_Name1.Text
    SavedVariablesCountingEnd.Cells(dataRow, "I").Formula = "=SUM('Counting (Exhaustive)'!G2:G1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "J").Formula = "=SUM('Counting (Exhaustive)'!G2:G1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "K").Formula = "=STDEV.S('Counting (Exhaustive)'!G2:G1048576)"
    Else
    End If
    
    If Not txt_Now2.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "L").Value = txt_Name2.Text
    SavedVariablesCountingEnd.Cells(dataRow, "M").Formula = "=SUM('Counting (Exhaustive)'!I2:I1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "N").Formula = "=SUM('Counting (Exhaustive)'!I2:I1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "O").Formula = "=STDEV.S('Counting (Exhaustive)'!I2:I1048576)"
    Else
    End If
    
    If Not txt_Now3.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "P").Value = txt_Name3.Value
    SavedVariablesCountingEnd.Cells(dataRow, "Q").Formula = "=SUM('Counting (Exhaustive)'!K2:K1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "R").Formula = "=SUM('Counting (Exhaustive)'!K2:K1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "S").Formula = "=STDEV.S('Counting (Exhaustive)'!K2:K1048576)"
    Else
    End If
    
    If Not txt_Now4.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "T").Value = txt_Name4.Text
    SavedVariablesCountingEnd.Cells(dataRow, "U").Formula = "=SUM('Counting (Exhaustive)'!M2:M1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "V").Formula = "=SUM('Counting (Exhaustive)'!M2:M1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "W").Formula = "=STDEV.S('Counting (Exhaustive)'!M2:M1048576)"
    Else
    End If
    
    If Not txt_Now5.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "X").Value = txt_Name5.Text
    SavedVariablesCountingEnd.Cells(dataRow, "Y").Formula = "=SUM('Counting (Exhaustive)'!O2:O1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "Z").Formula = "=SUM('Counting (Exhaustive)'!O2:O1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AA").Formula = "=STDEV.S('Counting (Exhaustive)'!O2:O1048576)"
    Else
    End If
    
    If Not txt_Now6.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "AB").Value = txt_Name6.Text
    SavedVariablesCountingEnd.Cells(dataRow, "AC").Formula = "=SUM('Counting (Exhaustive)'!Q2:Q1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AD").Formula = "=SUM('Counting (Exhaustive)'!Q2:Q1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AE").Formula = "=STDEV.S('Counting (Exhaustive)'!Q2:Q1048576)"
    Else
    End If
    
    If Not txt_Now7.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "AF").Value = txt_Name7.Text
    SavedVariablesCountingEnd.Cells(dataRow, "AG").Formula = "=SUM('Counting (Exhaustive)'!S2:S1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AH").Formula = "=SUM('Counting (Exhaustive)'!S2:R1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AI").Formula = "=STDEV.S('Counting (Exhaustive)'!S2:S1048576)"
    Else
    End If
    
    If Not txt_Now8.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "AJ").Value = txt_Name8.Text
    SavedVariablesCountingEnd.Cells(dataRow, "AK").Formula = "=SUM('Counting (Exhaustive)'!U2:U1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AL").Formula = "=SUM('Counting (Exhaustive)'!U2:U1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AM").Formula = "=STDEV.S('Counting (Exhaustive)'!U2:U1048576)"
    Else
    End If
    
    If Not txt_Now9.Value = "" Then
    SavedVariablesCountingEnd.Cells(dataRow, "AN").Value = txt_Name9.Text
    SavedVariablesCountingEnd.Cells(dataRow, "AO").Formula = "=SUM('Counting (Exhaustive)'!W2:W1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AP").Formula = "=SUM('Counting (Exhaustive)'!W2:W1048576) / COUNT('Counting (Exhaustive)'!C2:C1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AQ").Formula = "=STDEV.S('Counting (Exhaustive)'!W2:W1048576)"
    Else
    End If
    
    SavedVariablesCountingEnd.Cells(dataRow, "AR").Formula = "=SUM('Counting (Exhaustive)'!X2:Y1048576)"
    SavedVariablesCountingEnd.Cells(dataRow, "AS").Formula = "=STDEV.S('Counting (Exhaustive)'!X2:X1048576)"
    
    '
    '
    '
    
    ' Populate Total, Y3 and SD output values.
    ' # 1
    Total1 = Worksheets("Counting (Summary)").Range("I2").Value
    Y3_1 = Worksheets("Counting (Summary)").Range("J2").Value
    If Not Now1 = Total1 Then
        SD1 = Worksheets("Counting (Summary)").Range("K2").Value
    Else
    End If
    
    If Len(txt_Name1.Text) > 0 Then
        txt_Total1.Enabled = True
        txt_Total1.Text = Format(Total1, "0")
        txt_Total1.BackColor = RGB(255, 255, 255)
        
        txt_Y3_1.Enabled = True
        txt_Y3_1.Text = Format(Y3_1, "0.000")
        txt_Y3_1.BackColor = RGB(255, 255, 255)
    
        txt_SD1.Enabled = True
        txt_SD1.Text = Format(SD1, "0.000")
        txt_SD1.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 2
    Total2 = Worksheets("Counting (Summary)").Range("M2").Value
    Y3_2 = Worksheets("Counting (Summary)").Range("N2").Value
    If Not Now2 = Total2 Then
        SD2 = Worksheets("Counting (Summary)").Range("O2").Value
    Else
    End If
    
    If Len(txt_Name2.Text) > 0 Then
        txt_Total2.Enabled = True
        txt_Total2.Text = Format(Total2, "0")
        txt_Total2.BackColor = RGB(255, 255, 255)
        
        txt_Y3_2.Enabled = True
        txt_Y3_2.Text = Format(Y3_2, "0.000")
        txt_Y3_2.BackColor = RGB(255, 255, 255)
        
        txt_SD2.Enabled = True
        txt_SD2.Text = Format(SD2, "0.000")
        txt_SD2.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 3
    Total3 = Worksheets("Counting (Summary)").Range("Q2").Value
    Y3_3 = Worksheets("Counting (Summary)").Range("R2").Value
    If Not Now3 = Total3 Then
        SD3 = Worksheets("Counting (Summary)").Range("s2").Value
    Else
    End If
  
    If Len(txt_Name3.Text) > 0 Then
        txt_Total3.Enabled = True
        txt_Total3.Text = Format(Total3, "0")
        txt_Total3.BackColor = RGB(255, 255, 255)
        
        txt_Y3_3.Enabled = True
        txt_Y3_3.Text = Format(Y3_3, "0.000")
        txt_Y3_3.BackColor = RGB(255, 255, 255)
        
        txt_SD3.Enabled = True
        txt_SD3.Text = Format(SD3, "0.000")
        txt_SD3.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 4
    Total4 = Worksheets("Counting (Summary)").Range("U2").Value
    Y3_4 = Worksheets("Counting (Summary)").Range("V2").Value
    If Not Now4 = Total4 Then
        SD4 = Worksheets("Counting (Summary)").Range("W2").Value
    Else
    End If
    
    If Len(txt_Name4.Text) > 0 Then
        txt_Total4.Enabled = True
        txt_Total4.Text = Format(Total4, "0")
        txt_Total4.BackColor = RGB(255, 255, 255)
        
        txt_Y3_4.Enabled = True
        txt_Y3_4.Text = Format(Y3_4, "0.000")
        txt_Y3_4.BackColor = RGB(255, 255, 255)
        
        txt_SD4.Enabled = True
        txt_SD4.Text = Format(SD4, "0.000")
        txt_SD4.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 5
    Total5 = Worksheets("Counting (Summary)").Range("Y2").Value
    Y3_5 = Worksheets("Counting (Summary)").Range("Z2").Value
    If Not Now5 = Total5 Then
        SD5 = Worksheets("Counting (Summary)").Range("AA2").Value
    Else
    End If
     
    If Len(txt_Name5.Text) > 0 Then
        txt_Total5.Enabled = True
        txt_Total5.Text = Format(Total5, "0")
        txt_Total5.BackColor = RGB(255, 255, 255)
        
        txt_Y3_5.Enabled = True
        txt_Y3_5.Text = Format(Y3_5, "0.000")
        txt_Y3_5.BackColor = RGB(255, 255, 255)

        txt_SD5.Enabled = True
        txt_SD5.Text = Format(SD5, "0.000")
        txt_SD5.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 6
    Total6 = Worksheets("Counting (Summary)").Range("AC2").Value
    Y3_6 = Worksheets("Counting (Summary)").Range("AD2").Value
    If Not Now6 = Total6 Then
        SD6 = Worksheets("Counting (Summary)").Range("AE2").Value
    Else
    End If
    
    If Len(txt_Name6.Text) > 0 Then
        txt_Total6.Enabled = True
        txt_Total6.Text = Format(Total6, "0")
        txt_Total6.BackColor = RGB(255, 255, 255)
        
        txt_Y3_6.Enabled = True
        txt_Y3_6.Text = Format(Y3_6, "0.000")
        txt_Y3_6.BackColor = RGB(255, 255, 255)
        
        txt_SD6.Enabled = True
        txt_SD6.Text = Format(SD6, "0.000")
        txt_SD6.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 7
    Total7 = Worksheets("Counting (Summary)").Range("AG2").Value
    Y3_7 = Worksheets("Counting (Summary)").Range("AH2").Value
    If Not Now7 = Total7 Then
        SD7 = Worksheets("Counting (Summary)").Range("AI2").Value
    Else
    End If
    
    If Len(txt_Name7.Text) > 0 Then
        txt_Total7.Enabled = True
        txt_Total7.Text = Format(Total7, "0")
        txt_Total7.BackColor = RGB(255, 255, 255)
        
        txt_Y3_7.Enabled = True
        txt_Y3_7.Text = Format(Y3_7, "0.000")
        txt_Y3_7.BackColor = RGB(255, 255, 255)
        
        txt_SD7.Enabled = True
        txt_SD7.Text = Format(SD7, "0.000")
        txt_SD7.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 8
    Total8 = Worksheets("Counting (Summary)").Range("AK2").Value
    Y3_8 = Worksheets("Counting (Summary)").Range("AL2").Value
    If Not Now8 = Total8 Then
        SD8 = Worksheets("Counting (Summary)").Range("AM2").Value
    Else
    End If

    If Len(txt_Name8.Text) > 0 Then
        txt_Total8.Enabled = True
        txt_Total8.Text = Format(Total8, "0")
        txt_Total8.BackColor = RGB(255, 255, 255)
        
        txt_Y3_8.Enabled = True
        txt_Y3_8.Text = Format(Y3_8, "0.000")
        txt_Y3_8.BackColor = RGB(255, 255, 255)
        
        txt_SD8.Enabled = True
        txt_SD8.Text = Format(SD8, "0.000")
        txt_SD8.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 9
    Total9 = Worksheets("Counting (Summary)").Range("AO2").Value
    Y3_9 = Worksheets("Counting (Summary)").Range("AP2").Value
    If Not Now8 = Total8 Then
        SD9 = Worksheets("Counting (Summary)").Range("AQ2").Value
    Else
    End If

    If Len(txt_Name9.Text) > 0 Then
        txt_Total9.Enabled = True
        txt_Total9.Text = Format(Total9, "0")
        txt_Total9.BackColor = RGB(255, 255, 255)
        
        txt_Y3_9.Enabled = True
        txt_Y3_9.Text = Format(Y3_9, "0.000")
        txt_Y3_9.BackColor = RGB(255, 255, 255)
        
        txt_SD9.Enabled = True
        txt_SD9.Text = Format(SD9, "0.000")
        txt_SD9.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' # 0
    Total0 = Worksheets("Counting (Summary)").Range("E2").Value
    Y3_0 = Worksheets("Counting (Summary)").Range("F2").Value
    If Not Now0 = Total0 Then
        SD0 = Worksheets("Counting (Summary)").Range("G2").Value
    Else
    End If
    
    If Len(txt_Name0.Text) > 0 Then
        txt_Total0.Enabled = True
        txt_Total0.Text = Format(Total0, "0")
        txt_Total0.BackColor = RGB(255, 255, 255)
        
        txt_Y3_0.Enabled = True
        txt_Y3_0.Text = Format(Y3_0, "0.000")
        txt_Y3_0.BackColor = RGB(255, 255, 255)
        
        txt_SD0.Enabled = True
        txt_SD0.Text = Format(SD0, "0.000")
        txt_SD0.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    CountingSaved = True
        
End Sub

'
' End and Save
'

Private Sub CommandButton_Save_Click()
    If SavedVariablesCountingEndExists Then
    
        If TimerRunning Then
        CommandButton_Timer_Click
        TimerRunning = False
        End If
        
        X = Worksheets("Counting (Summary)").Range("AR2").Value
        N = Total0
        
        If OriginFOVSTarget Or OriginFOVSMarker Then
            response = MsgBox("Is this dataset a calibration count (FOVS method)?", vbQuestion + vbYesNo, "Calibration Count?")
            ' Check user response
            If response = vbYes Then
                N3C = NumFOV + 1
                If Not Worksheets("Counting (Exhaustive)").Range("X2").Value = Worksheets("Counting (Summary)").Range("AR2").Value Then
                    s3 = Worksheets("Counting (Summary)").Range("AS2").Value
                Else
                End If
                
                If IsNumeric(txt_TimeTotal.Value) Then
                    TimeTotal = CLng(txt_TimeTotal.Value)
                End If
            Else
                N3E = NumFOV + 1
            End If
        Else
            N3C = NumFOV + 1
            
            If IsNumeric(txt_TimeTotal.Value) Then
                TimeTotal = CLng(txt_TimeTotal.Value)
            End If
        End If
        
        'Update userforms with new counts.
        If OriginStarter Then
            CalculatorStart.txt_X.Value = X
            CalculatorStart.txt_N.Value = N
            CalculatorStart.txt_N3c.Value = N3C
            CalculatorStart.txt_TimeTotal.Value = TimeTotal
            
            OriginStarter = False
            
        '   CalculatorStart.Hide
            CalculatorStart.Show
            
        ElseIf OriginLinear Then
            CalculatorLinear.txt_X.Value = X
            CalculatorLinear.txt_N.Value = N
            
            OriginLinear = False
            
            CalculatorLinear.Hide
            CalculatorLinear.Show
            
        ElseIf OriginFOVSTarget Then
            CalculatorFOVSTarget.txt_N_FOVS.Value = N
            If response = vbNo Then
                CalculatorFOVSTarget.txt_N3E.Value = N3E
            Else
            End If
            
            OriginFOVSTarget = False
            
            CalculatorFOVSTarget.Hide
            CalculatorFOVSTarget.Show
            
        ElseIf OriginFOVSMarker Then
            CalculatorFOVSMarker.txt_X_FOVS.Value = X
            If response = vbNo Then
                CalculatorFOVSMarker.txt_N3E.Value = N3E
            Else
            End If
            
            OriginFOVSMarker = False
            
            CalculatorFOVSMarker.Hide
            CalculatorFOVSMarker.Show
            
        ElseIf OriginCountingEffort Then
            CountingEffort.txt_X.Value = X
            CountingEffort.txt_N.Value = N
            If response = vbYes Then
                CountingEffort.txt_N3c.Value = N3C
            Else
            End If
            CountingEffort.txt_TimeTotal.Value = TimeTotal
            
            OriginCountingEffort = False
            
            CountingEffort.Hide
            CountingEffort.Show
            
        ElseIf OriginCalibrationFOV Then
            CalibratorFOV.txt_X_FOVS.Value = X
            CalibratorFOV.txt_N3c.Value = N3C
            CalibratorFOV.txt_S3.Value = s3
            
            OriginCalibrationFOV = False
            
            CalibratorFOV.Hide
            CountingEffort.Show
        Else
        End If
        
        ' Move saved count end variables to the next row so as to not overwrite
        dataRow = dataRow + 1
    Else
        CommandButton_Next_Click
    End If

End Sub


'
' Clear inputs
'

Private Sub CommandButton_Clear_Click()
    ' Display a message box confirming the action and asking for confirmation
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    
    ' Check user's response
    If response = vbYes Then
        txt_Now1.Text = ""
        txt_Now2.Text = ""
        txt_Now3.Text = ""
        txt_Now4.Text = ""
        txt_Now5.Text = ""
        txt_Now6.Text = ""
        txt_Now7.Text = ""
        txt_Now8.Text = ""
        txt_Now9.Text = ""
        txt_Now0.Text = ""
    Else
    ' User cancelled, do nothing
    End If
End Sub

'
' Extra
'

    Private Sub CommandButton_Timer_Click()
        If TimerRunning = False Then
           ' Start the timer
            TimerRunning = True
            StartTime = Timer ' Get the current time
    
            CommandButton_Timer.Caption = "Pause timer" 'Change text to this while timer is running.
            CommandButton_ClearTimer.Enabled = True
            
            Do While TimerRunning
                txt_TimeTotal.Locked = True ' Disable the textbox from being edited.
                ElapsedSeconds = Timer - StartTime + PausedTime ' Calculate elapsed time in seconds.
                txt_TimeTotal.Text = Format(ElapsedSeconds, "0.0") ' Display elapsed time in text box.
                DoEvents ' Allow other events to be processed.
            Loop
        Else
            TimerRunning = False
            PausedTime = ElapsedSeconds ' Store the elapsed time when pausing
            CommandButton_Timer.Caption = "Resume timer"
        End If
    End Sub

    Private Sub CommandButton_ClearTimer_Click()
    
        If IsNumeric(txt_TimeTotal) Then 'Check if time value is not empty.
            response = MsgBox("This will reset the timer. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Reset timer?")
    
            ' Check response.
            If response = vbNo Then 'Clear inputs.
                Exit Sub
            Else
                ' Do nothing.
            End If
        End If

        ' Clear the results by updating the captions of labels or values of text boxes.
        txt_TimeTotal.Locked = False
        ElapsedSeconds = 0
        PausedTime = 0
        txt_TimeTotal.Text = "0.0" ' Update the textbox to display zero.
        CommandButton_Timer.Caption = "Start timer" ' Reset the caption of the timer button.
        CommandButton_ClearTimer.Enabled = False
        
        ' If the timer was running, stop it.
        If TimerRunning Then
            TimerRunning = False
        End If
    End Sub
