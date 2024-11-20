VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CountingEffort 
   Caption         =   "Optimisation data"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "CountingEffort.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CountingEffort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private InputsSaved As Boolean
    Private InputEmptyAny As Boolean
    Dim InputEmptyX As Boolean
    Dim InputEmptyN As Boolean
    Dim InputEmptyN3C As Boolean
    Dim InputEmptyTimeFOV As Boolean
    Dim InputEmptyTimeTotal As Boolean
    Private UnsavedWarningGiven As Boolean

' Populate inputs with previous values if these exist. Sets the value of the textbox to the value of the public variable.
Private Sub UserForm_Initialize()
    
    UnsavedWarningGiven = False
        
    If X <> 0 Then
        txt_X.Text = X
    Else
        txt_X.Text = ""
    End If
    
    If N <> 0 Then
        txt_N.Text = N
    Else
        txt_N.Text = ""
    End If

    If N3C <> 0 Then
        txt_N3c.Text = N3C
    Else
        txt_N3c.Text = ""
    End If
    
    If TimeFOV <> 0 Then
        txt_TimeFOV.Text = TimeFOV
    Else
        txt_TimeFOV.Text = ""
    End If
    
    If TimeTotal <> 0 Then
        txt_TimeTotal.Text = TimeTotal
    Else
        txt_TimeTotal.Text = ""
    End If
End Sub

'
' Saving
'
    ' Check if there are unsaved inputs
    ' Activate standard deviation text boxes if number of samples > 1.
    
    ' X
    
    Private Sub txt_X_Change()
        If InputsSaved And txt_X.Value <> X Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' N
    
    Private Sub txt_N_Change()
        If InputsSaved And txt_N.Value <> N Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' N3C
    
    Private Sub txt_N3C_Change()
        If InputsSaved And txt_N3c.Value <> N3C Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' TimeFOV
    
    Private Sub txt_TimeFOV_Change()
        If InputsSaved And txt_TimeFOV.Value <> TimeFOV Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' TimeTotal
    
    Private Sub txt_TimeTotal_Change()
        If InputsSaved And txt_TimeTotal.Value <> TimeTotal Then
            UnsavedWarningGiven = False
        End If
    End Sub

Private Sub CommandButtonSave_Click()
    ' Validate other input fields to see if not empty
    
    ' X
    If IsNumeric(txt_X.Value) Then
        X = CLng(txt_X.Value)
        InputEmptyX = False
    Else
        InputEmptyX = True
    End If
    
    ' N
    If IsNumeric(txt_N.Value) Then
        N = CLng(txt_N.Value)
        InputEmptyN = False
    Else
        InputEmptyN = True
    End If
    
    ' N3C
    If IsNumeric(txt_N3c.Value) Then
        N3C = CLng(txt_N3c.Value)
        InputEmptyN3C = False
    Else
        InputEmptyN3C = True
    End If
    
    ' TimeFOV
    If IsNumeric(txt_TimeFOV.Value) Then
        TimeFOV = CLng(txt_TimeFOV.Value)
        InputEmptyTimeFOV = False
    Else
        InputEmptyTimeFOV = True
    End If
    
    ' TimeTotal
    If IsNumeric(txt_TimeTotal.Value) Then
        TimeTotal = CLng(txt_TimeTotal.Value)
        InputEmptyTimeTotal = False
    Else
        InputEmptyTimeTotal = True
    End If
       
    ' Avoid zeros in counts
    If X <= 0 Then
        MsgBox "Number of targets needs to be higher than 0.", vbExclamation
        Exit Sub
    End If
    
    If N <= 0 Then
        MsgBox "Number of markers needs to be higher than 0.", vbExclamation
        Exit Sub
    End If
    
    ' Defining formulas and keep variables in memory.
        
    If X > 0 And N > 0 And N3C > 0 And TimeFOV > 0 And TimeTotal > 0 Then

        Y3x = X / N3C
        
        Y3n = N / N3C
        
        uhat = X / N
        
        TimeFOVTotal = TimeFOV * (N3C - 1)
        
        TimeTotalNoFOV = TimeTotal - TimeFOVTotal
        
        TimePerSpecimen = TimeTotalNoFOV / (X + N)
        
            'Routine to check if TimePerSpecimen is greater than 0. If so, stops futher calculations to avoid a division by zero in the next step.
            If Not TimePerSpecimen > 0 Then
                MsgBox "The time it takes to count specimens must be greater than 0.", vbExclamation
                Exit Sub ' Stop futher calculations, avoid division by zero.
            End If
        
        FOVTransitionEffort = TimeFOV / TimePerSpecimen
        
        If uhat <> 1 Then
            Ystar3n = (2 * FOVTransitionEffort) * ((uhat * uhat) + Sqr((uhat * uhat * uhat) * (1 + (uhat * (uhat - 1))))) / (uhat * ((uhat + 1) * ((uhat - 1) * (uhat - 1)))) ' When targets [x] are more common than markers [n].
            Ystar3x = (2 * FOVTransitionEffort) * ((uhat * uhat) + Sqr((uhat * uhat * uhat) * (1 + (uhat * (uhat - 1))))) / ((uhat + 1) * ((uhat - 1) * (uhat - 1))) ' When markers [n] are more common than targets [x].
        Else
            ' uhat = 1 would mean same amount of targets [x] and markers [n], which would lead to a division by zero in the next step.
            ' Don't calculate.
        End If
         
        ' Select which equation to consider for method determination.
        
        If uhat <> 1 And X > N Then
            MethodDetFactor = Ystar3x
        ElseIf uhat <> 1 And N > X Then
            MethodDetFactor = Ystar3n
        Else
            MethodDetFactor = 0 ' Avoiding division by 0 in the previous step by directly assinging its value to be 0 here (equal amount of targets and markers).
        End If
            
    Else
        ' Do not run calculations.
    End If
    
    ' Check if MethodDetFactor is greater than, less than, or equal to Y3x.
    
    If X > N Then ' Targets [x] more common than markers [n].
        If MethodDetFactor <> 0 And MethodDetFactor > Y3x Then
            LinearSuggested = True
        ElseIf MethodDetFactor = 0 Then
            LinearSuggested = True
        ElseIf MethodDetFactor <> 0 And MethodDetFactor < Y3x Then
            TargetSuggested = True
        End If
    ElseIf N > X Then ' Markers [n] more common than targets [x].
        If MethodDetFactor <> 0 And MethodDetFactor > Y3n Then
            LinearSuggested = True
        ElseIf MethodDetFactor = 0 Then
            LinearSuggested = True
        ElseIf MethodDetFactor <> 0 And MethodDetFactor < Y3n Then
            MarkerSuggested = True
        End If
    ElseIf X = N Then ' Markers [n] and targets [x] have same amount.
        LinearSuggested = True
    End If
  
    If InputEmptyX Or InputEmptyN Or InputEmptyN3C Or InputEmptyTimeFOV Or InputEmptyTimeTotal Then
        InputEmptyAny = True
    Else
        InputEmptyAny = False
    End If
  
    CountingEffortCalibration = True
    MsgBox "Variables successfully saved.", vbInformation
    
    'Background check to see if the best method has changed following calibration.
    'Commented out for now after discussion with Chris Mays.
    
'    If LinearSuggested And Not LinearChosen Then
'        response = MsgBox("Based on the data you have entered, the most efficient method for your sample is the Linear method. Would you like to switch to the Linear method calculator now?", vbQuestion + vbYesNo + vbDefaultButton1, "FOVS Method")
'        If response = vbYes Then
'            If FOVSTargetChosen Then
'                FOVSTargetChosen = False
'                CalculatorFOVSTarget.Hide
'            ElseIf FOVSMarkerChosen Then
'                FOVSMarkerChosen = False
'                CalculatorFOVSMarker.Hide
'            End If
'            LinearSuggested = False
'            LinearChosen = True
'            CalculatorLinear.Show
'        End If
'    ElseIf (TargetSuggested Or MarkerSuggested) And LinearChosen Then
'        response = MsgBox("Based on the data you have entered, the most efficient method for your sample is the FOVS method. Would you like to switch to the FOVS method calculator now?", vbQuestion + vbYesNo + vbDefaultButton1, "FOVS Method")
'        If response = vbYes Then
'            If TargetSuggested Then
'                TargetSuggested = False
'                LinearChosen = False
'                FOVSTargetChosen = True
'                CalculatorLinear.Hide
'                CalculatorFOVSTarget.Show
'            ElseIf MarkerSuggested Then
'                MarkerSuggested = False
'                LinearChosen = False
'                FOVSMarkerChosen = True
'                CalculatorLinear.Hide
'                CalculatorFOVSMarker.Show
'            End If
'        Else
'        End If
'    End If

    If LinearChosen And Not InputEmptyAny Then ' For Linear method
        CalculatorLinear.CommandButton_CountingEffort.BackColor = RGB(212, 236, 214) ' Greenish color
        CalculatorLinear.CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data ready)"
    ElseIf LinearChosen And InputEmptyAny Then
        CalculatorLinear.CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CalculatorLinear.CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
        
    ElseIf FOVSTargetChosen And Not InputEmptyAny Then ' For FOVS Target method
        CalculatorFOVSTarget.CommandButton_CountingEffort.BackColor = RGB(212, 236, 214) ' Greenish color
        CalculatorFOVSTarget.CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data ready)"
    ElseIf FOVSTargetChosen And InputEmptyAny Then
        CalculatorFOVSTarget.CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CalculatorFOVSTarget.CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
        
    ElseIf FOVSMarkerChosen And Not InputEmptyAny Then ' For FOVS Target method
        CalculatorFOVSMarker.CommandButton_CountingEffort.BackColor = RGB(212, 236, 214) ' Greenish color
        CalculatorFOVSMarker.CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data ready)"
    ElseIf FOVSMarkerChosen And InputEmptyAny Then
        CalculatorFOVSMarker.CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CalculatorFOVSMarker.CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
    Else
    End If
    
    ' Adding flags.
    
    InputsSaved = True
    UnsavedWarningGiven = False
    
    Me.Hide
End Sub

Private Sub CommandButtonClear_Click()
    ' Display a message box confirming the action and asking for confirmation
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    
    ' Check user's response
    If response = vbYes Then
        txt_X.Text = ""
        txt_N.Text = ""
        txt_TimeFOV.Text = ""
        txt_N3c.Text = ""
        txt_TimeTotal.Text = ""
    Else
    ' User cancelled, do nothing
    End If
End Sub

Private Sub CommandButton_Assistant_Click()
    OriginCountingEffort = True
    AssistantCounting.Show
End Sub

Private Sub txt_X_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_X.Text) > 0 Then
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

Private Sub txt_N_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_N.Text) > 0 Then
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

Private Sub txt_N3c_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_N3c.Text) > 0 Then
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


Private Sub txt_TimeFOV_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_TimeFOV.Text) > 0 Then
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

Private Sub txt_X_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_N_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_N3c_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_TimeFOV_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_TimeTotal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub CommandButtonTimer_Click()
    If TimerRunning = False Then
       ' Start the timer
        TimerRunning = True
        StartTime = Timer ' Get the current time

        CommandButtonTimer.Caption = "Pause timer" 'Change text to this while timer is running.
        CommandButtonClearTimer.Enabled = True
        
        Do While TimerRunning
            txt_TimeTotal.Locked = True ' Disable the textbox from being edited.
            ElapsedSeconds = Timer - StartTime + PausedTime ' Calculate elapsed time in seconds.
            txt_TimeTotal.Text = Format(ElapsedSeconds, "0.0") ' Display elapsed time in text box.
            DoEvents ' Allow other events to be processed.
        Loop
    Else
        TimerRunning = False
        PausedTime = ElapsedSeconds ' Store the elapsed time when pausing
        CommandButtonTimer.Caption = "Resume timer"
    End If
End Sub

    Private Sub CommandButtonClearTimer_Click()
    
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
        CommandButtonTimer.Caption = "Start timer" ' Reset the caption of the timer button.
        CommandButtonClearTimer.Enabled = False
        
        ' If the timer was running, stop it.
        If TimerRunning Then
            TimerRunning = False
        End If
    End Sub

'
' Shutdown
'

    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
        ' Ask user if they want to close the form without saving.
        ' For this warning to appear, check if inputs were saved, if input boxes contain values, and that these values are different from those stored in memory.
        
        ' To avoid several warnings in the case of many unsaved variables, the flag 'UnsavedWarningGiven' checks if such a warning has come up yet.
            
        
        If Not InputsSaved Then
            Unload Me
        End If
        
        ' X
        If Not ClearedAllData Then
            If Not UnsavedWarningGiven Then
                If IsNumeric(txt_X.Value) And txt_X.Value <> X Then
                    If CloseMode = 0 Then
                        Cancel = 1 ' Cancel the close operation.
                    End If
                    response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                    If response = vbYes Then
                        CommandButtonSave_Click 'Run subroutine to save inputs.
                        UnsavedWarningGiven = True
                    Else
                        Cancel = 0
                        UnsavedWarningGiven = False
                        Unload Me
                    End If
                End If
            End If
            
            ' N
            
            If Not UnsavedWarningGiven Then
                If IsNumeric(txt_N.Value) And txt_N.Value <> N Then
                    If CloseMode = 0 Then
                        Cancel = 1 ' Cancel the close operation.
                    End If
                    response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                    If response = vbYes Then
                        CommandButtonSave_Click 'Run subroutine to save inputs.
                        UnsavedWarningGiven = True
                    Else
                        Cancel = 0
                        UnsavedWarningGiven = False
                        Unload Me
                    End If
                End If
            End If
            
             ' N3C
            
            If Not UnsavedWarningGiven Then
                If IsNumeric(txt_N3c.Value) And txt_N3c.Value <> N3C Then
                    If CloseMode = 0 Then
                        Cancel = 1 ' Cancel the close operation.
                    End If
                    response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                    If response = vbYes Then
                        CommandButtonSave_Click 'Run subroutine to save inputs.
                        UnsavedWarningGiven = True
                    Else
                        Cancel = 0
                        UnsavedWarningGiven = False
                        Unload Me
                    End If
                End If
            End If
            
            ' TimeFOV
            
            If Not UnsavedWarningGiven Then
                If IsNumeric(txt_TimeFOV.Value) And txt_TimeFOV.Value <> TimeFOV Then
                    If CloseMode = 0 Then
                        Cancel = 1 ' Cancel the close operation.
                    End If
                    response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                    If response = vbYes Then
                        CommandButtonSave_Click 'Run subroutine to save inputs.
                        UnsavedWarningGiven = True
                    Else
                        Cancel = 0
                        UnsavedWarningGiven = False
                        Unload Me
                    End If
                End If
            End If
            
            ' TimeTotal
            
            If Not UnsavedWarningGiven Then
                If IsNumeric(txt_TimeTotal.Value) And txt_TimeTotal.Value <> TimeTotal Then
                    If CloseMode = 0 Then
                        Cancel = 1 ' Cancel the close operation.
                    End If
                    response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                    If response = vbYes Then
                        CommandButtonSave_Click 'Run subroutine to save inputs.
                        UnsavedWarningGiven = True
                    Else
                        Cancel = 0
                        UnsavedWarningGiven = False
                        Unload Me
                    End If
                End If
            End If
        End If
    
    ClearedAllData = False
    
    End Sub

