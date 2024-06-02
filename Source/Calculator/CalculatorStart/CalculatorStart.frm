VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalculatorStart 
   Caption         =   "Absolute abundance calculator v1.0 - Optimisation data & method determination"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135.001
   OleObjectBlob   =   "CalculatorStart.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CalculatorStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Variables
'

Private StarterCalculationRun As Boolean
   
'
' Startup
'

Private Sub UserForm_Initialize()
    'Introductory message. 'TODO include DOI when ready.
    If Not IntroductionGiven Then
        IntroductionGiven = True
        MsgBox "This calculator will assist in:" & vbNewLine & vbNewLine & "1) determining which method is most efficient for estimating absolute abundances in your population;" & vbNewLine & "2) estimating the absolute abundances of your specimens (and their associated precisions); and" & vbNewLine & "3) predicting the amount of effort required for a given precision." & vbNewLine & vbNewLine & "The next screen will ask for preliminary count data to determine the most efficient method. For this, please have a timer on hand, or use the provided timer before commencing counting. Approximately 10 fields of view are sufficient for most populations." & vbNewLine & vbNewLine & "Alternatively, you may wish to skip this step by selecting either the 'Linear' or 'FOVS' buttons." & vbNewLine & vbNewLine & "For more details on the terms used, the formulae and when to use these methods, see the original manuscript (https://doi.org/10.XXXXXX).", vbInformation
    Else
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

'
' Calculations
'

    Private Sub CommandButtonCalculate_Click()
    
        ' Check if calculation was already run.
        
        If StarterCalculationRun Then
            response = MsgBox("The new calculation will overwrite the old output data. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "New Calculation")
    
            ' Check response.
            If response = vbNo Then 'Clear inputs.
                Exit Sub
            Else
                ' Do nothing.
            End If
        End If
        
        ' Validate input fields to see if not empty.
        
        If Not IsNumeric(txt_X.Value) Then
            MsgBox "Please enter the number of targets counted [x].", vbExclamation, "Input Required"
            Exit Sub
        End If
        
        If Not IsNumeric(txt_N.Value) Then
            MsgBox "Please enter the number of markers counted [n].", vbExclamation, "Input Required"
            Exit Sub
        End If
        
        If Not IsNumeric(txt_TimeFOV.Value) Then
            MsgBox "Please enter the average time (seconds) it takes to change fields of view.", vbExclamation, "Input Required"
            Exit Sub
        End If
        
        If Not IsNumeric(txt_N3c.Value) Then
            MsgBox "Please enter the number of observed fields-of-view.", vbExclamation, "Input Required"
            Exit Sub
        End If
        
        If Not IsNumeric(txt_TimeTotal.Value) Then
            MsgBox "Please enter the total time (seconds) that counting took.", vbExclamation, "Input Required"
            Exit Sub
        End If
        
        ' Store input values in memory as Long (no expected decimal places).
        
        X = CLng(txt_X.Value)
        N = CLng(txt_N.Value)
        TimeTotal = CLng(txt_TimeTotal.Value)
        N3C = CLng(txt_N3c.Value)
        TimeFOV = CLng(txt_TimeFOV.Value)
        
        ' Avoid zeros in counts.
        
        If Not X > 0 Then
            MsgBox "Number of targets [x] needs to be higher than 0.", vbExclamation
            Exit Sub
        End If
        
        If Not N > 0 Then
            MsgBox "Number of markers [n] needs to be higher than 0.", vbExclamation
            Exit Sub
        End If
        
        If Not N3C > 0 Then
            MsgBox "Number of counted fields of view [N3C] needs to be higher than 0.", vbExclamation
            Exit Sub
        End If
        
        If Not TimeFOV > 0 Then
            MsgBox "Transition time needs to be higher than 0.", vbExclamation
            Exit Sub
        End If
        
        ' Defining formulas and keep variables in memory.
        
        Y3x = X / N3C
        
        Y3n = N / N3C
        
        uhat = X / N
        
        TimeFOVTotal = TimeFOV * (N3C - 1)
        
        TimeTotalNoFOV = TimeTotal - TimeFOVTotal
        
        TimePerSpecimen = TimeTotalNoFOV / (X + N)
            
            'Routine to check if TimePerSpecimen is greater than 0. If so, stops futher calculations to avoid a division by zero in the next step.
            If Not TimePerSpecimen > 0 Then
                MsgBox "The time it takes to count specimens must be greater than 0.", vbExclamation
                Exit Sub
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
            
        ' Display calculated results. Enable and colour output backgrounds to white.
        
        txtResult_Y3x.Text = Format(Y3x, "0.000")

        If IsNumeric(txtResult_Y3x.Value) Then ' Check if a number exists.
            txtResult_Y3x.Enabled = True
            txtResult_Y3x.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txtResult_Y3x.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        txtResult_Y3n.Text = Format(Y3n, "0.000")
        
        If IsNumeric(txtResult_Y3n.Value) Then ' Check if a number exists.
            txtResult_Y3n.Enabled = True
            txtResult_Y3n.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txtResult_Y3n.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        txtResult_uhat.Text = Format(uhat, "0.000")
        
        If IsNumeric(txtResult_uhat.Value) Then ' Check if a number exists.
            txtResult_uhat.Enabled = True
            txtResult_uhat.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txtResult_uhat.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        txtResult_FOVTransitionEffort.Text = Format(FOVTransitionEffort, "0.000")
             
        If IsNumeric(txtResult_FOVTransitionEffort.Value) Then ' Check if a number exists.
            txtResult_FOVTransitionEffort.Enabled = True
            txtResult_FOVTransitionEffort.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txtResult_FOVTransitionEffort.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        ' Check if MethodDetFactor is greater than, less than, or equal to Y3x.
        
        If X > N Then ' Targets [x] more common than markers [n].
            If MethodDetFactor <> 0 And MethodDetFactor > Y3x Then
                txtResult_BestMethod.Text = "Linear"
            ElseIf MethodDetFactor = 0 Then
                txtResult_BestMethod.Text = "Linear"
            ElseIf MethodDetFactor <> 0 And MethodDetFactor < Y3x Then
                txtResult_BestMethod.Text = "Field-of-view subsampling (FOVS)"
                TargetSuggested = True
            End If
        ElseIf N > X Then ' Markers [n] more common than targets [x].
            If MethodDetFactor <> 0 And MethodDetFactor > Y3n Then
                txtResult_BestMethod.Text = "Linear"
            ElseIf MethodDetFactor = 0 Then
                txtResult_BestMethod.Text = "Linear"
            ElseIf MethodDetFactor <> 0 And MethodDetFactor < Y3n Then
                txtResult_BestMethod.Text = "Field-of-view subsampling (FOVS)"
                MarkerSuggested = True
            End If
        ElseIf X = N Then ' Markers [n] and targets [x] have same amount.
            txtResult_BestMethod.Text = "Linear"
        End If
        
        ' Enable and colour Best Method background to white.
        
        If Len(txtResult_BestMethod.Text) > 0 Then ' Check if a text is present.
            txtResult_BestMethod.Enabled = True
            txtResult_BestMethod.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txtResult_BestMethod.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
            
        ' Change Linear button to green if best method.
        
        If txtResult_BestMethod = "Linear" Then
            CommandButtonLinear.BackColor = RGB(0, 255, 0) ' Green colour.
        Else
            CommandButtonLinear.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        ' Change FOVS button to green if best method.
        
        If txtResult_BestMethod = "Field-of-view subsampling (FOVS)" Then
            CommandButtonFOVS.BackColor = RGB(0, 255, 0) ' Green colour.
        Else
            CommandButtonFOVS.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        ' Add flag to memory for any potential checks in future steps.
        
        CountingEffortCalibration = True
        StarterCalculationRun = True
        
    End Sub

'
' Associated tools
'
    'Counting Assistant.
    
    Private Sub CommandButton_Assistant_Click()
        OriginStarter = True ' Add flag to memory for any potential checks in future steps.
        AssistantCounting.Show
        Hide
    End Sub
    
    ' Clear inputs.
    
    Private Sub CommandButtonClear_Click()
        ' Display a message box confirming the action and asking for confirmation.
        response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
        
        ' Check response.
        If response = vbYes Then 'C lear inputs.
            txt_X.Text = ""
            txt_N.Text = ""
            txt_TimeFOV.Text = ""
            txt_N3c.Text = ""
            txt_TimeTotal.Text = ""
        Else
            ' Do nothing.
        End If
    End Sub
    
    'Linear method button.
    
    Private Sub CommandButtonLinear_Click()
        LinearChosen = True ' Add flag to memory for any potential checks in future steps.
        
        ' Store inputs in case these exist but calculation was not run.
        
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
        
        ' TimeFOV
        
        If IsNumeric(txt_TimeFOV.Value) Then
            TimeFOV = CLng(txt_TimeFOV.Value)
        Else
        End If
                
        ' TimeTotal
        
        If IsNumeric(txt_TimeTotal.Value) Then
            TimeTotal = CLng(txt_TimeTotal.Value)
        Else
        End If
        
        response = MsgBox("The linear method is most efficient if preliminary data about the sample and the markers are entered. From these, the optimal amount of sampling effort can be calculated." & vbNewLine & vbNewLine & "Would you like to include preliminary data?", vbQuestion + vbYesNo + vbDefaultButton2, "Preliminary Data")

        ' Check response.
        If response = vbYes Then
            PreliminaryData.Show
        Else
            CalculatorLinear.Show
        End If
        
        Hide
    End Sub
    
    'FOVS method button.
    
    Private Sub CommandButtonFOVS_Click()
        If Not (TargetSuggested Or MarkerSuggested) And (txt_X.Value = "" Or txt_N.Value = "") Then
            response = MsgBox("Do targets appear to be more common than markers?", vbQuestion + vbYesNo, "Most common specimens?")
            ' Check user response
            If response = vbYes Then
                TargetSuggested = True ' Add flag to memory for any potential checks in future steps.
            Else
                MarkerSuggested = True ' Add flag to memory for any potential checks in future steps.
            End If
        ElseIf Not (TargetSuggested Or MarkerSuggested) And (txt_X.Value >= txt_N.Value) Then
            FOVSTargetChosen = True ' Add flag to memory for any potential checks in future steps.
            CalculatorFOVSTarget.Show ' FOVS calculator with equations that consider targets [x] being more common than markers [n].
        ElseIf Not (TargetSuggested Or MarkerSuggested) And (txt_X.Value < txt_N.Value) Then
            FOVSMarkerChosen = True ' Add flag to memory for any potential checks in future steps.
            CalculatorFOVSMarker.Show ' FOVS calculator with equations that consider markers [n] being more common than targets [x].
        End If
        
        response = MsgBox("The FOVS method is most efficient if preliminary data about the sample and the markers are entered. From these, the optimal field of view counts can be calculated for maximum efficiency. Would you like to insert preliminary data?", vbQuestion + vbYesNo + vbDefaultButton2, "Preliminary Data")
        ' Check response.
        If response = vbYes Then
            PreliminaryData.Show
        Else
            If TargetSuggested Then
                FOVSTargetChosen = True ' Add flag to memory for any potential checks in future steps.
                CalculatorFOVSTarget.Show ' FOVS calculator with equations that consider targets [x] being more common than markers [n].
                TargetSuggested = False ' Remove from memory.
            ElseIf MarkerSuggested Then
                FOVSMarkerChosen = True ' Add flag to memory for any potential checks in future steps.
                CalculatorFOVSMarker.Show ' FOVS calculator with equations that consider markers [n] being more common than targets [x].
                TargetSuggested = False ' Remove from memory.
            End If
        End If

        Hide
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
