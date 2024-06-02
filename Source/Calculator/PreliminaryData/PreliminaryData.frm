VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PreliminaryData 
   Caption         =   "Preliminary data"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390.001
   OleObjectBlob   =   "PreliminaryData.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "PreliminaryData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Variables
'
    Private InputsSaved As Boolean
    Private UnsavedWarningGiven As Boolean
    
'
' Startup
'
    Private Sub UserForm_Initialize()
            
        ' Populate inputs with previous values if these exist. Sets the value of the textbox to the value of the public variable.
        
        ' X
        
        If X <> 0 Then
            txt_X.Text = X
        Else
            txt_X.Text = ""
        End If
        
        ' N
        
        If N <> 0 Then
            txt_N.Text = N
        Else
            txt_N.Text = ""
        End If
        
        ' N3C
        
        If N3C <> 0 Then
            txt_N3c.Text = N3C
        Else
            txt_N3c.Text = ""
        End If
        
        ' TimeFOV
        
        If TimeFOV <> 0 Then
            txt_TimeFOV.Text = TimeFOV
        Else
            txt_TimeFOV.Text = ""
        End If
        
        ' TimeTotal
        
        If TimeTotal <> 0 Then
            txt_TimeTotal.Text = TimeTotal
        Else
            txt_TimeTotal.Text = ""
        End If
        
        
        ' N1
        
        If N1 <> 0 Then
            txt_N1.Text = N1
        Else
            txt_N1.Text = ""
        End If
        
        ' Y1
        
        If Y1 <> 0 Then
            txt_Y1.Text = Y1
        Else
            txt_Y1.Text = ""
        End If
        
        ' S1
        
        If s1 <> 0 Then
            txt_s1.Text = s1
        Else
            txt_s1.Text = ""
        End If
        
        ' LevelError
        
        If LevelError <> 0 Then
            txt_LevelError.Text = LevelError
        Else
            txt_LevelError.Text = ""
        End If
        
        ' Check if origin is from Linear or FOVS. Changes text and color from red to green.
            
        If LinearChosen Then
            CommandButtonSkipPerliminary.Caption = "Skip to linear data collection"
        ElseIf FOVSTargetChosen Or FOVSMarkerChosen Then
            CommandButtonSkipPerliminary.Caption = "Skip to FOVS data collection"
        End If
        
    End Sub
        
'
' Inputs
'
    'Make sure only numbers are inserted.
    
    Private Sub txt_X_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace (if not already entered).
        Select Case KeyAscii
            Case 8 ' Backspace.
            Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            Case 48, 96 ' Numbers 0 and Numpad 0.
            If Len(txt_X.Text) > 0 Then ' Allow input if the textbox is not empty.
                ' Do nothing, allow input.
            Else
                KeyAscii = 0 ' Disallow input if the textbox is empty.
            End If
            Case Else
                KeyAscii = 0 ' Disallow other characters.
        End Select
    End Sub
    
    Private Sub txt_N_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
        Select Case KeyAscii
            Case 8 ' Backspace.
            Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            Case 48, 96 ' Numbers 0 and Numpad 0.
            If Len(txt_N.Text) > 0 Then ' Allow input if the textbox is not empty.
                ' Do nothing, allow input.
            Else
                KeyAscii = 0 ' Disallow input if the textbox is empty.
            End If
            Case Else
                KeyAscii = 0 ' Disallow other characters.
        End Select
    End Sub
    
    Private Sub txt_N3c_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered).
        Select Case KeyAscii
            Case 8 ' Backspace.
            Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            Case 48, 96 ' Numbers 0 and Numpad 0.
            If Len(txt_N3c.Text) > 0 Then ' Allow input if the textbox is not empty.
                ' Do nothing, allow input.
            Else
                KeyAscii = 0 ' Disallow input if the textbox is empty.
            End If
            Case Else
                KeyAscii = 0 ' Disallow other characters.
        End Select
    End Sub
    
    Private Sub txt_TimeFOV_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered)
        Select Case KeyAscii
            Case 8 ' Backspace.
            Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            Case 48, 96 ' Numbers 0 and Numpad 0.
            If Len(txt_TimeFOV.Text) > 0 Then ' Allow input if the textbox is not empty.
                ' Do nothing, allow input.
            Else
                KeyAscii = 0 ' Disallow input if the textbox is empty.
            End If
            Case Else
                KeyAscii = 0 ' Disallow other characters.
        End Select
    End Sub
    
    Private Sub txt_TimeTotal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) ' Allow numbers (0-9) and Backspace key (if not already entered)
        Select Case KeyAscii
            Case 8
            Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            Case 48, 96 ' Numbers 0 and Numpad 0.
            If Len(txt_TimeTotal.Text) > 0 Then ' Allow input if the textbox is not empty.
                ' Do nothing, allow input
            Else
                
                KeyAscii = 0 ' Disallow input if the textbox is empty.
            End If
            Case Else
                KeyAscii = 0 ' Disallow other characters.
        End Select
    End Sub
    
    Private Sub txt_N1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Allow numbers (0-9) and Backspace key (if not already entered)
        Select Case KeyAscii
            Case 8
            Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            Case 48, 96 ' Numbers 0 and Numpad 0.
            If Len(txt_N1.Text) > 0 Then
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
    
    Private Sub txt_Y1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Allow numbers (0-9), Backspace, and Dot (.) key (if not already entered)
        Select Case KeyAscii
            Case 8 ' Backspace
                ' Do nothing, allow backspace
            Case 46 ' Dot
                If Len(txt_Y1.Text) = 0 Then
                    ' Disallow dot if textbox is empty
                    KeyAscii = 0
                ElseIf InStr(txt_Y1.Text, ".") > 0 Then
                    ' Disallow dot if dot already exists
                    KeyAscii = 0
                End If
            Case 48 To 57, 96 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
                If Len(txt_Y1.Text) = 0 Then
                    ' Allow input of 0 if textbox is empty
                    ' Do nothing, allow input
                ElseIf txt_Y1.Text = "0" Then
                    ' Disallow input of 0 if it's already present
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0 ' Disallow other characters
        End Select
    End Sub
    
    Private Sub txt_s1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Allow numbers (0-9), Backspace, and Dot (.) key (if not already entered)
        Select Case KeyAscii
            Case 8 ' Backspace
                ' Do nothing, allow backspace
            Case 46 ' Dot
                If Len(txt_s1.Text) = 0 Then
                    ' Disallow dot if textbox is empty
                    KeyAscii = 0
                ElseIf InStr(txt_s1.Text, ".") > 0 Then
                    ' Disallow dot if dot already exists
                    KeyAscii = 0
                End If
            Case 48 To 57, 96 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
                If Len(txt_s1.Text) = 0 Then
                    ' Allow input of 0 if textbox is empty
                    ' Do nothing, allow input
                ElseIf txt_s1.Text = "0" Then
                    ' Disallow input of 0 if it's already present
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0 ' Disallow other characters
        End Select
    End Sub
    
    Private Sub txt_LevelError_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Allow numbers (0-9), Backspace, and Dot (.) key (if not already entered)
        Select Case KeyAscii
            Case 8 ' Backspace
            Case 46 ' Dot
                If Len(txt_LevelError.Text) = 0 Then
                    ' Disallow dot if textbox is empty
                    KeyAscii = 0
                ElseIf InStr(txt_LevelError.Text, ".") > 0 Then
                    ' Disallow dot if dot already exists
                    KeyAscii = 0
                End If
            Case 48 To 57, 96 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
                ' Find position of decimal point if present
                DecimalPosition = InStr(txt_LevelError.Text, ".")
                If DecimalPosition > 0 Then
                    ' Calculate number of digits after decimal point
                    NumDigitsAfterDecimal = Len(txt_LevelError.Text) - DecimalPosition
                    If NumDigitsAfterDecimal >= 2 Then
                        ' Block more than two digits after the decimal
                        KeyAscii = 0
                    End If
                End If
                ' Additional check for '0' as first character
                If (KeyAscii = 48 Or KeyAscii = 96) And Len(txt_LevelError.Text) = 0 And DecimalPosition = 0 Then
                    ' Disallow '0' if it's the first character and no decimal point
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
    
    Private Sub txt_N1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
            KeyCode = 0
        End If
    End Sub

    Private Sub txt_Y1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
            KeyCode = 0
        End If
    End Sub

    Private Sub txt_S1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
            KeyCode = 0
        End If
    End Sub
    
    Private Sub txt_LevelError_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
            KeyCode = 0
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
    
    ' N1

    Private Sub txt_N1_Change()
        If InputsSaved And txt_N1.Value <> N1 Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' Y1
    
    Private Sub txt_Y1_Change()
        If InputsSaved And txt_Y1.Value <> Y1 Then
            UnsavedWarningGiven = False
        End If
    End Sub

' Both functions commented out below were for the userbox to be automatically populated by the square root of Y1 in case N1 = 1.

'    Private Sub txt_N1_Change()
        
'        If IsNumeric(txt_N1.Value) And txt_N1.Value > 1 Then
'            txt_s1.Value = ""
'            txt_s1.Enabled = True
'            txt_s1.BackColor = RGB(255, 255, 255) ' White colour.
'        Else
'            txt_s1.Value = ""
'            txt_s1.Enabled = False
'            txt_s1.BackColor = RGB(224, 224, 224) ' Grey colour.
'        End If
        
'        If InputsSaved And txt_N1.Value <> N1 Then
'            UnsavedWarningGiven = False
'        End If
               
'    End Sub
        
'    Private Sub txt_Y1_Change()
        
'        If IsNumeric(txt_N1.Value) And txt_N1.Value > 1 Then
'            txt_s1.Value = ""
'            txt_s1.Enabled = True
'            txt_s1.BackColor = RGB(255, 255, 255) ' White colour.
'        Else
'            txt_s1.Value = ""
'            txt_s1.Enabled = False
'            txt_s1.BackColor = RGB(224, 224, 224) ' Grey colour.
'        End If
    
'        If InputsSaved And txt_Y1.Value <> Y1 Then
'            UnsavedWarningGiven = False
'        End If
               
'    End Sub
    
    ' S1
    
    Private Sub txt_S1_Change()
        If InputsSaved And txt_s1.Value <> s1 Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' LevelError
    
    Private Sub txt_LevelError_Change()
        If InputsSaved And txt_LevelError.Value <> LevelError Then
            UnsavedWarningGiven = False
        End If
    End Sub
       
' Store values in memory after clicking save.

    Private Sub CommandButtonSave_Click()
                
        ' X
        If IsNumeric(txt_X.Value) Then
            X = CLng(txt_X.Value)
        Else
        End If
        
        ' N
        If IsNumeric(txt_N.Value) Then
            N = CLng(txt_N.Value)
        Else
        End If
        
        ' N3C
        If IsNumeric(txt_N3c.Value) Then
            N3C = CLng(txt_N3c.Value)
        Else
        End If
        
        ' TimeTotal
        If IsNumeric(txt_TimeTotal.Value) Then
            TimeTotal = CLng(txt_TimeTotal.Value)
        Else
        End If
        
        ' TimeFOV
        If IsNumeric(txt_TimeFOV.Value) Then
            TimeFOV = CLng(txt_TimeFOV.Value)
        Else
        End If

        ' N1
        
        If IsNumeric(txt_N1.Value) Then
            N1 = CLng(txt_N1.Value)
        Else
        End If
        
        ' Y1
        
        If IsNumeric(txt_Y1.Value) Then
            Y1 = CDbl(txt_Y1.Value)
        Else
        End If
        
        ' S1
        
        If IsNumeric(txt_s1.Value) Then
            s1 = CDbl(txt_s1.Value)
            
        ElseIf Not IsNumeric(txt_s1.Value) Then
        
            response = MsgBox("Note: You have not inserted the sample standard deviation of exotic markers per dose [s1]. Would you like to use the square-root of the mean numbers of exotic markers per dose [Y1] as an approximation?", vbQuestion + vbYesNo + vbDefaultButton2, "Determining s1")
        
            ' Check response.
            If response = vbYes And IsNumeric(txt_N1.Value) And IsNumeric(txt_Y1.Value) Then
                s1 = Sqr(Y1)
                
            ElseIf response = vbYes And Not (IsNumeric(txt_N1.Value) Or IsNumeric(txt_Y1.Value)) Then
                MsgBox "Please make sure the number of doses of exotic marker specimens [N1] and the mean number of exotic markers per dose [Y1] are filled out.", vbExclamation, "Input Required"
                Exit Sub
                
            Else
            End If
        Else
        End If
        
        ' LevelError
        
        If IsNumeric(txt_LevelError.Value) Then
            LevelError = CDbl(txt_LevelError.Value) 'TODO Can lead to negatives later on if less than 5.
        Else
        End If
        
        ' Avoid zeros in counts.
        
        If IsNumeric(txt_X.Value) And X = 0 Then
            MsgBox "Number of targets [x] needs to be higher than 0.", vbExclamation
            Exit Sub
        End If
        
        If IsNumeric(txt_N.Value) And N = 0 Then
            MsgBox "Number of markers [n] needs to be higher than 0.", vbExclamation
            Exit Sub
        End If
        
        ' Defining formulas and keep variables in memory.
        
        If X > 0 And N > 0 And N3C > 0 And TimeFOV > 0 And TimeTotal > 0 Then
               
            Y3x = X / N3C
            
            Y3n = N / N3C
            
            uhat = X / N
            
'            If LinearChosen Then
'                uhat = X / N
'            ElseIf FOVSTargetChosen Or TargetSuggested Then
'                uhat = xhat / N
'            ElseIf FOVSMarkerChosen Or MarkerSuggested Then
'                uhat = X / nhat
'            End If
            
            TimeFOVTotal = TimeFOV * (N3C - 1)
            
            TimeTotalNoFOV = TimeTotal - TimeFOVTotal
            
            TimePerSpecimen = TimeTotalNoFOV / (X + N)
                
                'Routine to check if TimePerSpecimen is greater than 0. If so, stops futher calculations to avoid a division by zero in the next step.
                If Not TimePerSpecimen > 0 Then
                    MsgBox "The time it takes to count specimens must be greater than 0.", vbExclamation
                    Exit Sub ' Stop futher calculations, avoid division by zero.
                End If
                    
            FOVTransitionEffort = TimeFOV / TimePerSpecimen
                                
            If LinearChosen Then
                If FOVTransitionEffort <> 0 Then 'Linear
                    eL = ((FOVTransitionEffort * (X / Y3x)) + X + N)
                Else
                    ' FOVTransitionEffort is equal to 0, do not run calculation
                End If
            End If
            
            If FOVSTargetChosen Or TargetSuggested Then
                If FOVTransitionEffort <> 0 Then 'FOVS Target
                    eF = (FOVTransitionEffort * N3C) + X + (FOVTransitionEffort * N3F) + N
                    deltastar = uhat * Sqr((FOVTransitionEffort + Y3x) / ((FOVTransitionEffort * uhat) + Y3x))
                Else
                    ' FOVTransitionEffort is equal to 0, do not run calculation
                End If
            End If
            
            If FOVSMarkerChosen Or MarkerSuggested Then
                If FOVTransitionEffort <> 0 Then 'FOVS Marker
                    eF = (FOVTransitionEffort * N3C) + X + (FOVTransitionEffort * N3F) + N
                    deltastar = uhat * Sqr((FOVTransitionEffort + (Y3n * uhat)) / ((FOVTransitionEffort * uhat) + (Y3n * uhat)))
                Else
                    ' FOVTransitionEffort is equal to 0, do not run calculation
                End If
            End If
        Else
            ' Do not run calculations.
        End If
        
        If FOVTransitionEffort <> 0 And N1 > 0 And Y1 > 0 And s1 > 0 Then
            If LinarChosen Then
                eL_sigmabar = (FOVTransitionEffort * (1 + uhat) + (Y3x * (2 + uhat)) + Y3x / uhat) / (Y3x * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
            ElseIf Not LinearChosen Then
                If FOVSTargetChosen Or TargetSuggested Then
                    Nstar3C = (1 / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3x + FOVTransitionEffort) + (Sqr(Y3x + (FOVTransitionEffort / uhat)))) / ((Y3x * (Sqr(Y3x + FOVTransitionEffort)))) 'TODO condition if LevelError is 0
                    Nstar3F = (uhat / (((LevelError / 100) * (LevelError / 100)) - ((N1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3x + FOVTransitionEffort) + (Sqr(Y3x + (uhat * FOVTransitionEffort)))) / (Y3x * (Sqr(Y3x + (uhat * FOVTransitionEffort))))
                    eF_sigmabar = ((2 * Y3x) + (FOVTransitionEffort * (1 + uhat) + 2 * (Sqr((Y3x + FOVTransitionEffort) * (Y3x + (uhat * FOVTransitionEffort)))))) / (Y3x * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
                ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                    Nstar3C = (1 / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3n + FOVTransitionEffort) + (Sqr(Y3n + (FOVTransitionEffort / uhat)))) / ((Y3n * (Sqr(Y3n + FOVTransitionEffort)))) 'TODO condition if LevelError is 0
                    Nstar3F = ((1 / uhat) / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3n + FOVTransitionEffort) + Sqr(Y3n + (FOVTransitionEffort / uhat))) / (Y3n * (Sqr(Y3n + (FOVTransitionEffort / uhat)))) 'TODO condition if LevelError is 0
                    eF_sigmabar = ((2 * (Y3n * uhat)) + (FOVTransitionEffort * (1 + uhat) + 2 * (Sqr(((Y3n * uhat) + FOVTransitionEffort) * ((Y3n * uhat) + (uhat * FOVTransitionEffort)))))) / ((Y3n * uhat) * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
                End If
            End If
        Else
            'Do not run calculation
        End If
                       
        MsgBox "Variables successfully saved.", vbInformation
        
        If LinearChosen Then
            CalculatorLinear.Show
        ElseIf FOVSTargetChosen Or TargetSuggested Then
            CalculatorFOVSTarget.Show
        ElseIf FOVSMarkerChosen Or MarkerSuggested Then
            CalculatorFOVSMarker.Show
        End If
        
        ' Adding flags.
        
        InputsSaved = True
        UnsavedWarningGiven = False
        
        Unload Me
        
    End Sub

'
' Associated tools
'
    'Counting Assistant.
    
    Private Sub CommandButton_Assistant_Click()
        OriginPreliminaryData = True ' Add flag to memory for any potential checks in future steps.
        AssistantCounting.Show
    End Sub
    
    ' Clear inputs.
    
    Private Sub CommandButtonClear_Click()
        ' Display a message box confirming the action and asking for confirmation.
        response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
        
        ' Check response.
        If response = vbYes Then 'Clear inputs.
            txt_X.Text = ""
            txt_N.Text = ""
            txt_TimeFOV.Text = ""
            txt_N3c.Text = ""
            txt_TimeTotal.Text = ""
            txt_N1.Text = ""
            txt_Y1.Text = ""
            txt_s1.Text = ""
            txt_LevelError.Text = ""
        Else
            ' Do nothing.
        End If
    End Sub
    
    'Linear method button.
    
    Private Sub CommandButtonSkipPerliminary_Click()
        OriginPreliminaryData = True
        
        If LinearChosen Then
            CalculatorLinear.Show
        ElseIf FOVSTargetChosen Or TargetSuggested Then
            CalculatorFOVSTarget.Show
        ElseIf FOVSMarkerChosen Or MarkerSuggested Then
            CalculatorFOVSMarker.Show
        End If
        
        Unload Me
    End Sub
       
    'Timer 'TODO maybe have a distinct userform for this?
    
    Private Sub CommandButtonTimer_Click()
        If TimerRunning = False Then
           ' Start the timer
            TimerRunning = True
            StartTime = Timer ' Get the current time
    
            CommandButtonTimer.Caption = "Pause timer" 'Change text to this while timer is running.
            CommandButtonClearTimer.Visible = True 'Make reset button visible.
            
            Do While TimerRunning
                txt_TimeTotal.Locked = True ' Disable the textbox from being edited.
                ElapsedSeconds = Timer - StartTime + PausedTime ' Calculate elapsed time in seconds.
                txt_TimeTotal.Text = Format(ElapsedSeconds, "0") ' Display elapsed time in text box.
                DoEvents ' Allow other events to be processed.
            Loop
        Else
            TimerRunning = False
            PausedTime = ElapsedSeconds ' Store the elapsed time when pausing
            CommandButtonTimer.Caption = "Resume timer"
        End If
    End Sub
    
    'Reset timer.
    
    Private Sub CommandButtonClearTimer_Click()
        ' Clear the results by updating the captions of labels or values of text boxes.
        txt_TimeTotal.Locked = False
        ElapsedSeconds = 0
        PausedTime = 0
        txt_TimeTotal.Text = "" ' Update the textbox to display nothing.
        CommandButtonTimer.Caption = "Start timer" ' Reset the caption of the timer button.
        CommandButtonClearTimer.Visible = False 'Make reset button invisible.
        
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
            If LinearChosen Then
                CalculatorLinear.Show
            ElseIf FOVSTargetChosen Or TargetSuggested Then
                CalculatorFOVSTarget.Show
            ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                CalculatorFOVSMarker.Show
            End If
            
            Unload Me
        End If
        
        ' X
        
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
                    Unload Me
                End If
            End If
        End If
        
        ' N1
        
        If Not UnsavedWarningGiven Then
            If IsNumeric(txt_N1.Value) And txt_N1.Value <> N1 Then
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
                    Unload Me
                End If
            End If
        End If
            
        ' Y1
        
        If Not UnsavedWarningGiven Then
            If IsNumeric(txt_Y1.Value) And txt_Y1.Value <> Y1 Then
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
                    Unload Me
                End If
            End If
        End If
            
        ' S1
            
        If Not UnsavedWarningGiven Then
            If IsNumeric(txt_s1.Value) And txt_s1.Value <> s1 Then
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
                    Unload Me
                End If
            End If
        End If
        
        ' LevelError
            
        If Not UnsavedWarningGiven Then
            If IsNumeric(txt_LevelError.Value) And txt_LevelError.Value <> LevelError Then
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
                    If LinearChosen Then
                        CalculatorLinear.Show
                    ElseIf FOVSTargetChosen Or TargetSuggested Then
                        CalculatorFOVSTarget.Show
                    ElseIf FOVSMarkerChosen Or MarkerSuggested Then
                        CalculatorFOVSMarker.Show
                    End If
                    Unload Me
                End If
            End If
        End If
        
    End Sub
