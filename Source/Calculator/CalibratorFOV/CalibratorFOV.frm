VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalibratorFOV 
   Caption         =   "FOVS method calibration counts"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   OleObjectBlob   =   "CalibratorFOV.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CalibratorFOV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Variables
'

Private InputsSaved As Boolean
Private InputEmptyAny As Boolean
Dim InputEmptyX As Boolean
Dim InputEmptyN As Boolean
Dim InputEmptyN3C As Boolean
Dim InputEmptyS3 As Boolean
Private UnsavedWarningGiven As Boolean

Private Sub CommandButton_Assistant_Click()
    OriginCalibrationFOV = True
    AssistantCounting.Show
End Sub

    Private Sub txt_X_FOVS_Change()
                
        If InputsSaved And FOVSTargetChosen Then
            If txt_X_FOVS.Value <> X Then
                UnsavedWarningGiven = False
            End If
        ElseIf InputsSaved And FOVSMakerChosen Then
            If txt_X_FOVS.Value <> N Then
                UnsavedWarningGiven = False
            End If
        End If
               
    End Sub
    
    Private Sub txt_N3C_Change()
        
        If IsNumeric(txt_N3c.Value) And txt_N3c.Value > 1 Then
            txt_S3.Value = ""
            txt_S3.Enabled = True
            txt_S3.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txt_S3.Value = ""
            txt_S3.Enabled = False
            txt_S3.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
    
        If InputsSaved And txt_N3c.Value <> N3C Then
            UnsavedWarningGiven = False
        End If
               
    End Sub
    
    Private Sub txt_S3_Change()
        If InputsSaved And txt_S3.Value <> s3 Then
            UnsavedWarningGiven = False
        End If
    End Sub

Private Sub CommandButtonSaveCalibrationFOV_Click()
    ' Store values in memory
    If IsNumeric(txt_X_FOVS.Value) And FOVSTargetChosen Then
        X = CLng(txt_X_FOVS.Value)
        InputEmptyX = False
    Else
        InputEmptyX = True
    End If
    
    If IsNumeric(txt_X_FOVS.Value) And FOVSMarkerChosen Then
        N = CLng(txt_X_FOVS.Value)
        InputEmptyN = False
    Else
        InputEmptyN = True
    End If
    
    ' N3C
    
    If IsNumeric(txt_N3c.Value) Then
        N3C = CDbl(txt_N3c.Value)
        InputEmptyN3C = False
    Else
        InputEmptyN3C = True
    End If
    
    ' S3
    
    If IsNumeric(txt_S3.Value) And txt_N3c.Value > 1 Then
        s3 = CDbl(txt_S3.Value)
        InputEmptyS3 = False
    ElseIf txt_N3c.Value = 1 Then
        s3 = Sqr(txt_X_FOVS)
        txt_S3.Value = s3
        txt_S3.Enabled = True
        txt_S3.BackColor = RGB(255, 255, 255) ' White colour.
        InputEmptyS3 = False
    Else
        InputEmptyS3 = True
    End If
    
    
    If FOVSTargetChosen Then
        If InputEmptyX Or InputEmptyN3C Or InputEmptyS3 Then
            InputEmptyAny = True
        Else
            InputEmptyAny = False
        End If
    ElseIf FOVSMarkerChosen Then
        If InputEmptyN Or InputEmptyN3C Or InputEmptyS3 Then
            InputEmptyAny = True
        Else
            InputEmptyAny = False
        End If
    End If
 
    ' Avoid zeros in counts
    If FOVSTargetChosen And X <= 0 Then
        MsgBox "Number of targets needs to be higher than 0.", vbExclamation
        Exit Sub
    ElseIf FOVSMarkerChosen And N <= 0 Then
        MsgBox "Number of markers needs to be higher than 0.", vbExclamation
        Exit Sub
    End If
           
    MsgBox "Variables successfully saved.", vbInformation

    If FOVSTargetChosen And Not InputEmptyAny Then ' For FOVS Target method
        CalculatorFOVSTarget.CommandButton_CalibrationFOV.BackColor = RGB(212, 236, 214) ' Greenish color
        CalculatorFOVSTarget.CommandButton_CalibrationFOV.Caption = "Field of view (FOV) calibration count" & vbCrLf & "(data ready)"
        
        
        CalculatorFOVSTarget.txt_N_FOVS.Enabled = True
        CalculatorFOVSTarget.txt_N_FOVS.BackColor = RGB(255, 255, 255)
        CalculatorFOVSTarget.txt_N3f.Enabled = True
        CalculatorFOVSTarget.txt_N3f.BackColor = RGB(255, 255, 255)
        
        ' Refresh certain outputs.
        uhat = X / N
        Y3x = X / N3C
        
        If FOVTransitionEffort <> 0 Then
            Nstar3C = (1 / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3x + FOVTransitionEffort) + (Sqr(Y3x + (FOVTransitionEffort / uhat)))) / ((Y3x * (Sqr(Y3x + FOVTransitionEffort)))) 'TODO condition if LevelError is 0
            Nstar3F = (uhat / (((LevelError / 100) * (LevelError / 100)) - ((N1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3x + FOVTransitionEffort) + (Sqr(Y3x + (uhat * FOVTransitionEffort)))) / (Y3x * (Sqr(Y3x + (uhat * FOVTransitionEffort))))
            deltastar = uhat * Sqr((FOVTransitionEffort + Y3x) / ((FOVTransitionEffort * uhat) + Y3x))
            eF = (FOVTransitionEffort * N3C) + X + (FOVTransitionEffort * N3F) + N
            eF_sigmabar = ((2 * Y3x) + (FOVTransitionEffort * (1 + uhat) + 2 * (Sqr((Y3x + FOVTransitionEffort) * (Y3x + (uhat * FOVTransitionEffort)))))) / (Y3x * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
        Else
            ' FOVTransitionEffort is equal to 0, do not run calculation
        End If
        
        ' Display outputs in calculator.
        
        CalculatorFOVSTarget.LabelResult_OptimalCalibrationFOV.Value = Format(Y3x, "0.00")
        CalculatorFOVSTarget.LabelResult_OptimalCalibrationFOV.Value = Format(Nstar3C, "0.00")
        CalculatorFOVSTarget.LabelResult_OptimalFullFOV.Value = Format(Nstar3F, "0.00")
        CalculatorFOVSTarget.LabelResult_OptimalRatioFOV.Value = Format(deltastar, "0.00")
        CalculatorFOVSTarget.LabelResult_CollectionEffort_FOVS.Value = Format(eF, "0")
        CalculatorFOVSTarget.LabelResult_PredictedCollectionEffort_FOVS.Value = Format(eF_sigmabar, "0")
        
    ElseIf FOVSMarkerChosen And Not InputEmptyAny Then ' For FOVS Marker method
        CalculatorFOVSMarker.CommandButton_CalibrationFOV.BackColor = RGB(212, 236, 214) ' Greenish color
        CalculatorFOVSMarker.CommandButton_CalibrationFOV.Caption = "Field of view (FOV) calibration count" & vbCrLf & "(data ready)"
        
        CalculatorFOVSMarker.txt_X_FOVS.Enabled = True
        CalculatorFOVSMarker.txt_X_FOVS.BackColor = RGB(255, 255, 255)
        CalculatorFOVSMarker.txt_N3f.Enabled = True
        CalculatorFOVSMarker.txt_N3f.BackColor = RGB(255, 255, 255)
        
        ' Refresh certain outputs.
        uhat = X / N
        Y3n = N / N3C
                
        If FOVTransitionEffort <> 0 Then
            Nstar3C = (1 / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3n + FOVTransitionEffort) + (Sqr(Y3n + (FOVTransitionEffort / uhat)))) / ((Y3n * (Sqr(Y3n + FOVTransitionEffort)))) 'TODO condition if LevelError is 0
            Nstar3F = ((1 / uhat) / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3n + FOVTransitionEffort) + Sqr(Y3n + (FOVTransitionEffort / uhat))) / (Y3n * (Sqr(Y3n + (FOVTransitionEffort / uhat)))) 'TODO condition if LevelError is 0
            deltastar = uhat * Sqr((FOVTransitionEffort + (Y3n * uhat)) / ((FOVTransitionEffort * uhat) + (Y3n * uhat)))
            eF = (FOVTransitionEffort * N3C) + X + (FOVTransitionEffort * N3F) + N
            eF_sigmabar = ((2 * (Y3n * uhat)) + (FOVTransitionEffort * (1 + uhat) + 2 * (Sqr(((Y3n * uhat) + FOVTransitionEffort) * ((Y3n * uhat) + (uhat * FOVTransitionEffort)))))) / ((Y3n * uhat) * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
        Else
            ' FOVTransitionEffort is equal to 0, do not run calculation
        End If
        
        ' Display outputs in calculator.
        
        CalculatorFOVSMarker.LabelResult_OptimalCalibrationFOV.Value = Format(Y3n, "0.00")
        CalculatorFOVSMarker.LabelResult_OptimalCalibrationFOV.Value = Format(Nstar3C, "0.00")
        CalculatorFOVSMarker.LabelResult_OptimalFullFOV.Value = Format(Nstar3F, "0.00")
        CalculatorFOVSMarker.LabelResult_OptimalRatioFOV.Value = Format(deltastar, "0.00")
        CalculatorFOVSMarker.LabelResult_CollectionEffort_FOVS.Value = Format(eF, "0")
        CalculatorFOVSMarker.LabelResult_PredictedCollectionEffort_FOVS.Value = Format(eF_sigmabar, "0")
    End If
    
    ' Adding flags.
    
    InputsSaved = True
    
    If Not InputEmptyAny Then
        CalibratedFOV = True
    Else
        CalibratedFOV = False
    End If
    
    UnsavedWarningGiven = False
    
    Me.Hide
End Sub

Private Sub CommandButtonClear_Click()
    ' Display a message box confirming the action and asking for confirmation
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    
    ' Check user's response
    If response = vbYes Then
        txt_X_FOVS.Text = "" ' Will clear both when targets or markers are the focus.
        txt_N3c.Text = ""
        txt_S3.Text = ""
    Else
    ' User cancelled, do nothing
    End If
End Sub

' Populate inputs with previous values if these exist. Sets the value of the textbox to the value of the public variable.
Private Sub UserForm_Initialize()
    If FOVSTargetChosen Then
        Label_X_FOVS.Caption = "Number of counted target specimens" & vbCrLf & "during calibration counts [x]"
        Label_S3.Caption = "Standard deviation of target specimens per" & vbCrLf & "field of view (from calibration counts) [s ]"
        
        txt_X_FOVS.Enabled = True
        txt_X_FOVS.BackColor = RGB(255, 255, 255) ' White colour.
        
        If X <> 0 Then
            txt_X_FOVS.Text = X
        Else
            txt_X_FOVS.Text = ""
        End If
    ElseIf FOVSMarkerChosen Then
        Label_X_FOVS.Caption = "Number of counted marker specimens" & vbCrLf & "during calibration counts [n]"
        Label_S3.Caption = "Standard deviation of marker specimens per" & vbCrLf & "field of view (from calibration counts) [s ]"
        
        txt_X_FOVS.Enabled = True
        txt_X_FOVS.BackColor = RGB(255, 255, 255) ' White colour.
        
        If N <> 0 Then
            txt_X_FOVS.Text = N
        Else
            txt_X_FOVS.Text = ""
        End If
    Else
    End If
    
    If N3C <> 0 Then
        txt_N3c.Text = N3C
    Else
        txt_N3c.Text = ""
    End If
    
    ' S3
    
    If s3 <> 0 Then
        txt_S3.Text = s3
        txt_S3.Enabled = True
        txt_S3.BackColor = RGB(255, 255, 255) ' White colour.
    ElseIf N3C > 1 Then
        txt_S3.Enabled = True
        txt_S3.BackColor = RGB(255, 255, 255) ' White colour.
    Else
        txt_S3.Text = ""
        txt_S3.Enabled = False
        txt_S3.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
End Sub

Private Sub txt_X_FOVS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_X_FOVS.Text) > 0 Then
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

Private Sub txt_s3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9), Backspace, and Dot (.) key (if not already entered)
    Select Case KeyAscii
        Case 8 ' Backspace
            ' Do nothing, allow backspace
        Case 46 ' Dot
            If Len(txt_S3.Text) = 0 Then
                ' Disallow dot if textbox is empty
                KeyAscii = 0
            ElseIf InStr(txt_S3.Text, ".") > 0 Then
                ' Disallow dot if dot already exists
                KeyAscii = 0
            End If
        Case 48 To 57, 96 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            If Len(txt_S3.Text) = 0 Then
                ' Allow input of 0 if textbox is empty
                ' Do nothing, allow input
            ElseIf txt_S3.Text = "0" Then
                ' Disallow input of 0 if it's already present
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0 ' Disallow other characters
    End Select
End Sub

' Avoid pasting words and numbers.

Private Sub txt_X_FOVS_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_N3c_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_s3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' Ask user if they want to close the form without saving.
    ' For this warning to appear, check if inputs were saved, if input boxes contain values, and that these values are different from those stored in memory.
    
    ' To avoid several warnings in the case of many unsaved variables, the flag 'UnsavedWarningGiven' checks if such a warning has come up yet.
    
    
    If Not ClearedAllData Then
    
        ' X
        If Not UnsavedWarningGiven And FOVSTargetChosen Then
            UnsavedWarningGiven = True
            If IsNumeric(txt_X_FOVS.Value) And txt_X_FOVS.Value <> X Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveCalibrationFOV_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
        
        ' N
        If Not UnsavedWarningGiven And FOVSMarkerChosen Then
            UnsavedWarningGiven = True
            If IsNumeric(txt_X_FOVS.Value) And txt_X_FOVS.Value <> N Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveCalibrationFOV_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
            
        ' N3C
        
        If Not UnsavedWarningGiven Then
            UnsavedWarningGiven = True
            If IsNumeric(txt_N3c.Value) And txt_N3c.Value <> N3C Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveCalibrationFOV_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
            
        ' S3
            
        If Not UnsavedWarningGiven Then
            If IsNumeric(txt_S3.Value) And txt_S3.Value <> s3 Then
            UnsavedWarningGiven = True
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveCalibrationFOV_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
    End If
    
    ClearedAllData = False
    
End Sub
