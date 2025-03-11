VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalculatorFOVSTarget 
   Caption         =   "Absolute abundance calculator v1.1.3 - FOVS method (Target focus)"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11895
   OleObjectBlob   =   "CalculatorFOVSTarget.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CalculatorFOVSTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private InputsSaved As Boolean
    Private OutputsSaved As Boolean
    Private InfoExported As Boolean

Private Sub CommandButton_Clear_Linear_Click()
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    ' Check user's response
    If response = vbYes Then
        txt_N_FOVS.Text = ""
        txt_N3E.Text = ""
    Else
    ' User cancelled, do nothing
    End If
End Sub

Private Sub CommandButton_Assistant_Click()
    OriginFOVSTarget = True
    AssistantCounting.Show
End Sub

Private Sub CommandButton_CalibrationFOV_Click()
    FOVSTargetChosen = True
    CalibratorFOV.Show
End Sub

Private Sub CommandButton_Clear_FOVS_Click()
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    ' Check user's response
    If response = vbYes Then
        txt_N_FOVS.Text = ""
        txt_N3E.Text = ""
        txt_LevelError.Text = ""
        InputsSaved = False
    Else
    ' User cancelled, do nothing
    End If
End Sub

Private Sub CommandButton_ClearAll_Click()
    response = MsgBox("Are you sure you want to clear all data? Data stored in worksheets is unaffected.", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Data")
    ' Check user's response
    If response = vbYes Then
        
        InputsSaved = False
        OutputsSaved = False
        
        ' Inputs
        
        txt_N_FOVS.Text = ""
        N = Empty
        
        txt_N3E.Text = ""
        N3E = Empty
        
        txt_LevelError.Text = ""
        LevelError = Empty
        
        ' Concentration
                     
        LabelResult_Concentration_FOVS.Text = ""
        c = Empty
        LabelResult_Concentration_FOVS.Enabled = False
        LabelResult_Concentration_FOVS.BackColor = RGB(224, 224, 224)
        
        txt_ConcentrationUnits.Text = ""
        UnitSize = Empty
        txt_ConcentrationUnits.Enabled = False
        txt_ConcentrationUnits.BackColor = RGB(224, 224, 224)
        
        LabelResult_ConcentrationStandardError_FOVS.Text = ""
        sigma_Fx = Empty
        LabelResult_ConcentrationStandardError_FOVS.Enabled = False
        LabelResult_ConcentrationStandardError_FOVS.BackColor = RGB(224, 224, 224)
        
        ' Optimal field of view counts
        
        LabelResult_OptimalCalibrationFOV.Text = ""
        Nstar3C = Empty
        LabelResult_OptimalCalibrationFOV.Enabled = False
        LabelResult_OptimalCalibrationFOV.BackColor = RGB(224, 224, 224)
        
        LabelResult_OptimalFullFOV.Text = ""
        Nstar3E = Empty
        LabelResult_OptimalFullFOV.Enabled = False
        LabelResult_OptimalFullFOV.BackColor = RGB(224, 224, 224)
        
        LabelResult_OptimalRatioFOV.Text = ""
        deltastar = Empty
        LabelResult_OptimalRatioFOV.Enabled = False
        LabelResult_OptimalRatioFOV.BackColor = RGB(224, 224, 224)
        
        ' Sampling effort
        
        LabelResult_CollectionEffort_FOVS.Text = ""
        eF = Empty
        LabelResult_CollectionEffort_FOVS.Enabled = False
        LabelResult_CollectionEffort_FOVS.BackColor = RGB(224, 224, 224)
        
        LabelResult_PredictedCollectionEffort_FOVS.Text = ""
        eF_sigmabar = Empty
        LabelResult_PredictedCollectionEffort_FOVS.Enabled = False
        LabelResult_PredictedCollectionEffort_FOVS.BackColor = RGB(224, 224, 224)
        
        ' Other sample parameters
        
        LabelResult_Y3x.Text = ""
        Y3x = Empty
        LabelResult_Y3x.Enabled = False
        LabelResult_Y3x.BackColor = RGB(224, 224, 224)
        
        LabelResult_uhat_FOVS.Text = ""
        uhat = Empty
        LabelResult_uhat_FOVS.Enabled = False
        LabelResult_uhat_FOVS.BackColor = RGB(224, 224, 224)
        
        LabelResult_FOVTransitionEffort.Text = ""
        FOVTransitionEffort = Empty
        LabelResult_FOVTransitionEffort.Enabled = False
        LabelResult_FOVTransitionEffort.BackColor = RGB(224, 224, 224)
        
        LabelResult_xhat.Text = ""
        xhat = Empty
        LabelResult_xhat.Enabled = False
        LabelResult_xhat.BackColor = RGB(224, 224, 224)
        
        ' Field of view calibration count
        
        X = Empty
        N3C = Empty
        s3 = Empty
        LevelError = Empty
        
        CalibratedFOV = False
        CommandButton_CalibrationFOV.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_CalibrationFOV.Caption = "Field of view (FOV) calibration count" & vbCrLf & "(data missing)"
        
        ' Marker and sample related
        
        N1 = Empty
        Y1 = Empty
        s1 = Empty
        N2 = Empty
        Y2 = Empty
        SizeUnit = Empty
        s2 = Empty
                
        SavedMarkerDetails = False
        CommandButton_MarkerCharacteristics.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data missing)"
        
        ' Optimisation data related
        
        TimeFOV = Empty
        TimeTotal = Empty
        
        CountingEffortCalibration = False
        CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
        
        'Unload associated user forms
        
        ClearedAllData = True
        
        Unload CalibratorFOV
        Unload MarkerCharacteristics
        Unload CountingEffort
        
        ' Disallow exporting data
        
'        If CommandButton_SaveVariables_FOVS.Enabled = True Then
'            CommandButton_SaveVariables_FOVS.Enabled = False
'        End If
        
    Else
    ' User cancelled, do nothing
    End If
End Sub

Private Sub CommandButton_Glossary_Click()
    Glossary.Show
End Sub

Private Sub CommandButton_MarkerCharacteristics_Click()
    FOVSTargetChosen = True
    MarkerCharacteristics.Show
End Sub

Private Sub CommandButton_MethodDetermination_Click()
    FOVSTargetChosen = True
    CalculatorStart.Show
    Me.Hide
End Sub

Private Sub CommandButton_SaveVariables_FOVS_Click()
    ' Validate other input fields to see if not empty.
'    If Not IsNumeric(txt_N_FOVS.Value) And Not ShutdownRequested Then
'        MsgBox "Please enter the number of markers counted [n].", vbExclamation, "Input Required"
'        Exit Sub
'    End If
    
'    If Not IsNumeric(txt_N3E.Value) And Not ShutdownRequested Then
'        MsgBox "Please enter the amount of observed fields-of-view seen in the extrapolation count [N3E].", vbExclamation, "Input Required"
'        Exit Sub
'    End If
    
'    If Not SavedMarkerDetails And Not ShutdownRequested Then
'        MsgBox "Please enter marker characteristics.", vbExclamation, "Input Required"
'        MarkerCharacteristics.Show
'        Exit Sub
'    End If
    
'    If Not CalibratedFOV And Not ShutdownRequested Then
'        MsgBox "Please attempt a FOV calibration.", vbExclamation, "Input Required"
'        MarkerCharacteristics.Show
'        Exit Sub
'    End If
    
    ' Make sure calculations are run first
'    If Not ShutdownRequested Then
'        CommandButton_Calculate_FOVS_Click
'    End If
     
    ' Initialize the variable to False
    SavedFOVSTargetExists = False
     
    ' Create a new worksheet named "Saved Variables (FOVS_Target)"
    ' Check if the sheet "Saved Variables (FOVS_Target)" already exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Saved Variables (FOVS_Target)" Then
            SavedVariablesFOVSTargetExists = True
            Set SavedVariablesFOVSTarget = ws
            Exit For
        End If
    Next ws

    ' If the sheet doesn't exist, create a new one
    If Not SavedVariablesFOVSTargetExists Then
        Set SavedVariablesFOVSTarget = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets("Calculator"))
        SavedVariablesFOVSTarget.Name = "Saved Variables (FOVS_Target)"
        AddHeadersFOVSTarget SavedVariablesFOVSTarget
        
        ' Clear nextRow and lastNonEmptyRow
        nextRow = 1
        lastNonEmptyRow = 0
    End If
    
    ' Determine the next empty row by examining all used columns
    lastNonEmptyRow = 0
    For i = 1 To 29 ' Data is up to column AC
        nextRow = SavedVariablesFOVSTarget.Cells(Rows.Count, i).End(xlUp).Row
        If nextRow > lastNonEmptyRow Then
            lastNonEmptyRow = nextRow
        End If
    Next i
    nextRow = lastNonEmptyRow + 1

    ' Write values from the userform to specific cells in the next available row
    SavedVariablesFOVSTarget.Cells(nextRow, "A").Value = Now
    SavedVariablesFOVSTarget.Cells(nextRow, "B").Value = lastNonEmptyRow
    SavedVariablesFOVSTarget.Cells(nextRow, "C").Value = txt_SampleName.Text
    SavedVariablesFOVSTarget.Cells(nextRow, "D").Value = txt_N_FOVS.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "E").Value = txt_N3E.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "F").Value = txt_LevelError.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "G").Value = MarkerCharacteristics.txt_N1.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "H").Value = MarkerCharacteristics.txt_Y1.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "I").Value = MarkerCharacteristics.txt_s1.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "J").Value = MarkerCharacteristics.txt_N2.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "K").Value = MarkerCharacteristics.txt_Y2.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "L").Value = MarkerCharacteristics.ComboBox_Units.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "M").Value = MarkerCharacteristics.txt_s2.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "N").Value = LabelResult_Concentration_FOVS.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "O").Value = LabelResult_ConcentrationStandardError_FOVS.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "P").Value = LabelResult_OptimalCalibrationFOV.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "Q").Value = LabelResult_OptimalFullFOV.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "R").Value = LabelResult_OptimalRatioFOV.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "S").Value = LabelResult_CollectionEffort_FOVS.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "T").Value = LabelResult_PredictedCollectionEffort_FOVS.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "U").Value = LabelResult_Y3x.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "V").Value = LabelResult_uhat_FOVS.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "W").Value = LabelResult_FOVTransitionEffort.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "X").Value = LabelResult_xhat.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "Y").Value = CalibratorFOV.txt_X_FOVS.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "Z").Value = CalibratorFOV.txt_N3c.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "AA").Value = CalibratorFOV.txt_S3.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "AB").Value = CountingEffort.txt_TimeFOV.Value
    SavedVariablesFOVSTarget.Cells(nextRow, "AC").Value = CountingEffort.txt_TimeTotal.Value
        
    ' Inform the user that values have been saved
    InfoExported = True
    MsgBox "Values have been saved to the worksheet 'Saved Variables (FOVS_Target)'.", vbInformation
    
    If ShutdownRequested Then
        End
    End If
End Sub

Private Sub AddHeadersFOVSTarget(ByRef ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "Date and time (DD/MM/YYYY XX:XX)"
        .Cells(1, 2).Value = "Data export #"
        .Cells(1, 3).Value = "Sample name"
        .Cells(1, 4).Value = "Number of marker specimens from extrapolation counts [n]"
        .Cells(1, 5).Value = "Number of extrapolation-count fields of view counted [N3E]"
        .Cells(1, 6).Value = "Desired level of total error [sigma-bar]"
        .Cells(1, 7).Value = "Number of doses of exotic marker specimens [N1]"
        .Cells(1, 8).Value = "Mean number of exotic markers per dose [Ybar1]"
        .Cells(1, 9).Value = "Sample standard deviation of exotic markers per dose [s1]"
        .Cells(1, 10).Value = "Total number of samples [N2]"
        .Cells(1, 11).Value = "Sample size (or mean sample size, if N2 > 1) [Ybar2]"
        .Cells(1, 12).Value = "Size unit"
        .Cells(1, 13).Value = "Standard deviation of sample size [s2]"
        .Cells(1, 14).Value = "Concentration of target specimens [cF]"
        .Cells(1, 15).Value = "Total standard error of concentration [sigma-F]"
        .Cells(1, 16).Value = "Optimal number of calibration-count FOVs [Nstar3C]"
        .Cells(1, 17).Value = "Optimal number of extrapolation-count FOVs [Nstar3E]"
        .Cells(1, 18).Value = "Optimal FOV count ratio (extrapolation-to-calibration ratio) [deltastar]"
        .Cells(1, 19).Value = "Present data collection effort (time units) [eF]"
        .Cells(1, 20).Value = "Predicted data collection effort to achieve desired error rate (time units) [eF-sigma-bar]"
        .Cells(1, 21).Value = "Mean number of targets per field of view [Yline3x]"
        .Cells(1, 22).Value = "Target-to-market ratio [u-hat]"
        .Cells(1, 23).Value = "Field of view transition effort factor [omegaline]"
        .Cells(1, 24).Value = "Estimate of target specimens from extrapolation counts [xhat]"
        .Cells(1, 25).Value = "Number of counted target specimens during calibration counts [x]"
        .Cells(1, 26).Value = "Number of fields of view counted during calibration counts [N3C]"
        .Cells(1, 27).Value = "Standard deviation of target specimens per field of view from calibration counts [s3]"
        .Cells(1, 28).Value = "Transition time (in seconds)"
        .Cells(1, 29).Value = "Total count time (in seconds)"
        
        ' Force three decimal places
        .Cells(1, 22).NumberFormat = "0.000"
    End With
End Sub

' Populate inputs with previous values if these exist. Sets the value of the textbox to the value of the public variable.
Private Sub UserForm_Initialize()
    If Not FOVSTargetIntroGiven Then
        MsgBox "The FOVS method requires a series of 'calibration counts' followed by a series of 'extrapolation counts'. To insert the calibration count data, press the 'Field of view (FOV) calibration count.'" & vbNewLine & vbNewLine & "Once these are filled, the 'extrapolation count' fields will be available." & vbNewLine & vbNewLine & "Absolute abundances (and associated error) will require the addition of data from the marker specimens being used. To do so, press the 'marker and sample characteristrics' button." & vbNewLine & vbNewLine & "(Optional: To predict the amount of sampling effort required for a given assemblage, insert the relevant data by pressing the 'optimisation data' button.)", vbInformation
        FOVSTargetIntroGiven = True
    Else
    End If
    ' Check if certain sheets are present. Iterate through all worksheets in the workbook.
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Exported data (FOVS-T)" Then
            ' Set the flag to True if the worksheet exists
            SavedVariablesFOVSTargetExists = True
            Exit For
        End If
    Next ws
    
    ' Inputs
    
    If Len(SampleName) > 1 Then
        txt_SampleName.Text = SampleName
    Else
    End If
    
    If LevelError <> 0 Then
        txt_LevelError = LevelError
    End If
             
'    If CalibratedFOV Then
'        txt_N3E.Enabled = True
'        txt_N3E.BackColor = RGB(255, 255, 255)
'    Else
'        txt_N3E.BackColor = RGB(224, 224, 224)
'    End If
    
    If Nstar3C <> 0 Then ' If Nstar3C is not equal to 0, render it in the label.
        LabelResult_OptimalCalibrationFOV.Enabled = True
        LabelResult_OptimalCalibrationFOV.Text = Format(Nstar3C, "0")
        LabelResult_OptimalCalibrationFOV.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_OptimalCalibrationFOV.BackColor = RGB(224, 224, 224)
    End If
    
    If Nstar3E <> 0 Then ' If Nstar3E is not equal to 0, render it in the label.
        LabelResult_OptimalFullFOV.Enabled = True
        LabelResult_OptimalFullFOV.Text = Format(Nstar3E, "0")
        LabelResult_OptimalFullFOV.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_OptimalFullFOV.BackColor = RGB(224, 224, 224)
    End If
    
    If deltastar <> 0 Then ' If deltastar is not equal to 0, render it in the label.
        LabelResult_OptimalRatioFOV.Enabled = True
        LabelResult_OptimalRatioFOV.Text = Format(deltastar, "0.00")
        LabelResult_OptimalRatioFOV.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_OptimalRatioFOV.BackColor = RGB(224, 224, 224)
    End If
    
    If Y3x <> 0 Then ' If Y3x is not equal to 0, render it in the label.
        LabelResult_Y3x.Enabled = True
        LabelResult_Y3x.Text = Format(Y3x, "0.000")
        LabelResult_Y3x.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_Y3x.BackColor = RGB(224, 224, 224)
    End If
    
    If uhat <> 0 Then ' If uhat is not equal to 0, render it in the label.
        LabelResult_uhat_FOVS.Enabled = True
        LabelResult_uhat_FOVS.Text = Format(uhat, "0.000")
        LabelResult_uhat_FOVS.BackColor = RGB(255, 255, 255)
    Else ' If uhat is equal to 0, do nothing.
        LabelResult_uhat_FOVS.BackColor = RGB(224, 224, 224)
    End If
        
    If FOVTransitionEffort <> 0 Then ' If FOVTransitionEffort is not equal to 0, render it in the label.
        LabelResult_FOVTransitionEffort.Enabled = True
        LabelResult_FOVTransitionEffort.Text = Format(FOVTransitionEffort, "0.000")
        LabelResult_FOVTransitionEffort.BackColor = RGB(255, 255, 255)
    Else ' If uhat is equal to 0, do nothing.
        LabelResult_FOVTransitionEffort.BackColor = RGB(224, 224, 224)
    End If
    
'    If xhat <> 0 Then ' If xhat is not equal to 0, render it in the label.
'        LabelResult_xhat.Enabled = True
'        LabelResult_xhat.Text = Format(xhat, "0")
'        LabelResult_xhat.BackColor = RGB(255, 255, 255)
'    Else ' If uhat is equal to 0, do nothing.
'        LabelResult_xhat.BackColor = RGB(224, 224, 224)
'    End If
    
    If eF <> 0 Then ' If eF is not equal to 0, render it in the label.
        LabelResult_CollectionEffort_FOVS.Enabled = True
        LabelResult_CollectionEffort_FOVS.Text = Format(eF, "0")
        LabelResult_CollectionEffort_FOVS.BackColor = RGB(255, 255, 255)
    Else ' If uhat is equal to 0, do nothing.
        LabelResult_CollectionEffort_FOVS.BackColor = RGB(224, 224, 224)
    End If
    
    If eF_sigmabar <> 0 Then ' If eF_sigmabar is not equal to 0, render it in the label.
        LabelResult_PredictedCollectionEffort_FOVS.Enabled = True
        LabelResult_PredictedCollectionEffort_FOVS.Text = Format(eF_sigmabar, "0")
        LabelResult_PredictedCollectionEffort_FOVS.BackColor = RGB(255, 255, 255)
    Else ' If uhat is equal to 0, do nothing.
        LabelResult_PredictedCollectionEffort_FOVS.BackColor = RGB(224, 224, 224)
    End If
    
    ' Check if marker details were saved. Changes color from red to green.
            
    If Not SavedMarkerDetails Then
        CommandButton_MarkerCharacteristics.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data missing)"
    Else
        CommandButton_MarkerCharacteristics.BackColor = RGB(212, 236, 214) ' Greenish color
        CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data ready)"
    End If
    
    ' Check if FOVS calibration was done. Changes color from red to green.
            
    If Not CalibratedFOV Then
        CommandButton_CalibrationFOV.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_CalibrationFOV.Caption = "Field of view (FOV) calibration count" & vbCrLf & "(data missing)"
    Else
        CommandButton_CalibrationFOV.BackColor = RGB(212, 236, 214) ' Greenish color
        CommandButton_CalibrationFOV.Caption = "Field of view (FOV) calibration count" & vbCrLf & "(data ready)"
    End If
    
    ' Check if counting effort calibration was done. Changes color from red to green.
            
    If Not CountingEffortCalibration Then
        CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
    Else
        CommandButton_CountingEffort.BackColor = RGB(212, 236, 214) ' Greenish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data ready)"
    End If
    
    ' Check if Preliminary calculations were done. Changes color from red to green.
            
    If Not X <> 0 And Not N <> 0 And Not N3C <> 0 And Not TimeFOV <> 0 And Not TimeTotal <> 0 Then
        CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
    Else
        CommandButton_CountingEffort.BackColor = RGB(212, 236, 214) ' Greenish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data ready)"
    End If
End Sub

' Check for changes in inputs.

Private Sub txt_N_FOVS_Change()
    If IsNumeric(txt_N_FOVS.Value) And txt_N_FOVS.Value <> N Then
        InputsSaved = False
    End If
End Sub

Private Sub txt_N3E_Change()
    If IsNumeric(txt_N3E.Value) And txt_N3E.Value <> N3E Then
        InputsSaved = False
    End If
End Sub

Private Sub txt_LevelError_Change()
    If IsNumeric(txt_LevelError) And txt_LevelError <> LevelError Then
        InputsSaved = False
    End If
End Sub

'
' Calculate
'

Private Sub CommandButton_Calculate_FOVS_Click()
   
' Validate other input fields to see if not empty.
      
    ' Store values in memory
    
    If Len(txt_SampleName.Text) > 1 Then
        SampleName = txt_SampleName.Text
    Else
    End If
     
    InputsSaved = True
    
   ' Perform background calculations
    If Not CalibratedFOV Or N3C = 0 Then
        MsgBox "Please attempt a FOV calibration.", vbExclamation, "Input Required"
        FOVSTargetChosen = True
        CalibratorFOV.Show
        Exit Sub
    End If
    
    ' Y3x
    
        Y3x = X / N3C
        
        LabelResult_Y3x.Text = Format(Y3x, "0.000")
        
        If IsNumeric(LabelResult_Y3x.Value) Then
            LabelResult_Y3x.Enabled = True
            LabelResult_Y3x.BackColor = RGB(255, 255, 255)
        Else
            LabelResult_Y3x.BackColor = RGB(224, 224, 224)
        End If
    
    'Dim Vline As Double ' Total mass or volume of samples ' TODO Include as in Linear?
    'Vline = N2 * Y2
   
    Dim c4 As Double ' Bias correction for calibration count no. of FOVs (N3)
    
    If N3C > 1 Then
        c4 = Sqr(2 / (N3C - 1)) * WorksheetFunction.Gamma(N3C / 2) / WorksheetFunction.Gamma((N3C - 1) / 2)
    Else
        MsgBox "Number of calibration count fields-of-view needs to be higher than 1.", vbExclamation
        Exit Sub
    End If
    
    Dim sigmahat3 As Double ' Unbiased estimator for the population standard deviation
    sigmahat3 = s3 / c4
   
    Dim s3P As Double ' Proportional corrected sample standard deviation - common grains/FOV
    s3P = (sigmahat3 / Y3x)

    ' Check if N values exist before running following calculations.
    
    If Len(txt_N_FOVS.Text) = 0 Then
        MsgBox "Please enter the number of markers counted in the extrapolation counts [n].", vbExclamation, "Input Required"
        txt_N_FOVS.SetFocus
        Exit Sub
    End If
    
    N = CLng(txt_N_FOVS.Value)
    
    If N <= 0 Then
        MsgBox "Number of markers needs to be higher than 0.", vbExclamation
        txt_N_FOVS.SetFocus
        Exit Sub
    End If
    
    ' Variable defined as public in ShowCalculator module
    
    If Len(txt_N3E.Value) > 0 Then
        N3E = CLng(txt_N3E.Value)
    Else
    End If
    
    If N3E = Empty Then
        uhat = X / N
    Else
        xhat = N3E * Y3x
        uhat = xhat / N
        
        LabelResult_xhat.Text = Format(xhat, "0")
        If LabelResult_xhat.Value >= 0 Then
            LabelResult_xhat.Enabled = True
            LabelResult_xhat.BackColor = RGB(255, 255, 255)
        Else
            LabelResult_xhat.BackColor = RGB(224, 224, 224)
        End If
    End If
    
    LabelResult_uhat_FOVS.Text = Format(uhat, "0.000")
    
    If IsNumeric(LabelResult_uhat_FOVS.Value) Then
        LabelResult_uhat_FOVS.Enabled = True
        LabelResult_uhat_FOVS.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_uhat_FOVS.BackColor = RGB(224, 224, 224)
    End If

    ' Perform visible calculations
    
    If Not SavedMarkerDetails Then
        MsgBox "Please enter marker and sample characteristics.", vbExclamation, "Input Required"
        MarkerCharacteristics.Show
        Exit Sub
    End If
    
    Dim c As Double ' Mean number of target specimens per unit mass or volume
    If xhat <> Empty Then
        ' C = (xhat * Y1 * N1) / (N * Vline) ' TODO Include as in Linear? Going with below code.
        c = (xhat * Y1 * N1) / (N * Y2)
    Else
        c = (uhat * Y1 * N1) / (N * Y2)
    End If
    
    LabelResult_Concentration_FOVS = Format(c, "0")
    txt_ConcentrationUnits = SizeUnit
    
    'Enable and colour output backgrounds to white
    If IsNumeric(LabelResult_Concentration_FOVS.Value) Then
        LabelResult_Concentration_FOVS.Enabled = True
        LabelResult_Concentration_FOVS.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_Concentration_FOVS.BackColor = RGB(224, 224, 224)
    End If
    
    If Len(txt_ConcentrationUnits) > 0 Then
        txt_ConcentrationUnits.Enabled = True
        txt_ConcentrationUnits.BackColor = RGB(255, 255, 255)
    Else
        txt_ConcentrationUnits.BackColor = RGB(224, 224, 224)
    End If
          
    If FOVTransitionEffort = 0 Then
        MsgBox "Please enter Optimisation Data.", vbExclamation, "Input Required"
        CountingEffort.Show
        Exit Sub
    Else
        LabelResult_FOVTransitionEffort.Text = Format(FOVTransitionEffort, "0.000")
        If IsNumeric(LabelResult_FOVTransitionEffort.Value) And FOVTransitionEffort <> 0 Then
            LabelResult_FOVTransitionEffort.Enabled = True
            LabelResult_FOVTransitionEffort.BackColor = RGB(255, 255, 255)
        ElseIf IsNumeric(LabelResult_FOVTransitionEffort.Value) And FOVTransitionEffort = 0 Then
        ' FOVTransitionEffort is equal to 0, do not show
            LabelResult_FOVTransitionEffort.Text = ""
        Else
            LabelResult_FOVTransitionEffort.BackColor = RGB(224, 224, 224)
        End If
        
        deltastar = uhat * Sqr((FOVTransitionEffort + Y3x) / ((FOVTransitionEffort * uhat) + Y3x))
        LabelResult_OptimalRatioFOV.Text = Format(deltastar, "0.00")
        
        If IsNumeric(LabelResult_OptimalRatioFOV.Value) And FOVTransitionEffort <> 0 Then
            LabelResult_OptimalRatioFOV.Enabled = True
            LabelResult_OptimalRatioFOV.BackColor = RGB(255, 255, 255)
        ElseIf IsNumeric(LabelResult_OptimalRatioFOV.Value) And FOVTransitionEffort = 0 Then
            ' FOVTransitionEffort is equal to 0, do not show
            LabelResult_OptimalRatioFOV.Text = ""
        Else
            LabelResult_OptimalRatioFOV.BackColor = RGB(224, 224, 224)
        End If
    End If
    
    If FOVTransitionEffort <> 0 Then
        eF = (FOVTransitionEffort * N3C) + X + (FOVTransitionEffort * N3E) + N
        LabelResult_CollectionEffort_FOVS.Text = Format(eF, "0")
        
        If IsNumeric(LabelResult_CollectionEffort_FOVS.Value) And FOVTransitionEffort <> 0 Then
            LabelResult_CollectionEffort_FOVS.Enabled = True
            LabelResult_CollectionEffort_FOVS.BackColor = RGB(255, 255, 255)
        ElseIf IsNumeric(LabelResult_CollectionEffort_FOVS.Value) And FOVTransitionEffort = 0 Then
            ' FOVTransitionEffort is equal to 0, do not show
            LabelResult_CollectionEffort_FOVS.Text = ""
        Else
            LabelResult_CollectionEffort_FOVS.BackColor = RGB(224, 224, 224)
        End If
    Else
        ' FOVTransitionEffort is equal to 0, do not run calculation
    End If
    
    If Not IsNumeric(txt_LevelError.Value) Then
        MsgBox "Please enter the desired target level of error as a percentage (e.g., 10).", vbExclamation, "Input Required"
        txt_LevelError.SetFocus
        Exit Sub
    End If
    
    LevelError = CDbl(txt_LevelError.Value) 'TODO Can lead to negatives later on if less than 5.
    
    If FOVTransitionEffort <> 0 And SavedMarkerDetails Then
        Nstar3C = (1 / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3x + FOVTransitionEffort) + (Sqr(Y3x + (FOVTransitionEffort * uhat)))) / ((Y3x * (Sqr(Y3x + FOVTransitionEffort)))) 'TODO condition if LevelError is 0
        LabelResult_OptimalCalibrationFOV.Text = Format(Nstar3C, "0")
        
        If IsNumeric(LabelResult_OptimalCalibrationFOV.Value) And FOVTransitionEffort <> 0 Then
            LabelResult_OptimalCalibrationFOV.Enabled = True
            LabelResult_OptimalCalibrationFOV.BackColor = RGB(255, 255, 255)
        ElseIf IsNumeric(LabelResult_OptimalCalibrationFOV.Value) And FOVTransitionEffort = 0 Then
            ' FOVTransitionEffort is equal to 0, do not show
            LabelResult_OptimalCalibrationFOV.Text = ""
        Else
            LabelResult_OptimalFullFOV.BackColor = RGB(224, 224, 224)
        End If
        
        Nstar3E = (uhat / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3x + FOVTransitionEffort) + (Sqr(Y3x + (uhat * FOVTransitionEffort)))) / (Y3x * (Sqr(Y3x + (uhat * FOVTransitionEffort))))
        LabelResult_OptimalFullFOV.Text = Format(Nstar3E, "0")
        
        If IsNumeric(LabelResult_OptimalFullFOV.Value) And FOVTransitionEffort <> 0 Then
            LabelResult_OptimalFullFOV.Enabled = True
            LabelResult_OptimalFullFOV.BackColor = RGB(255, 255, 255)
        ElseIf IsNumeric(LabelResult_OptimalFullFOV.Value) And FOVTransitionEffort = 0 Then
            ' FOVTransitionEffort is equal to 0, do not show
            LabelResult_OptimalFullFOV.Text = ""
        Else
            LabelResult_OptimalFullFOV.BackColor = RGB(224, 224, 224)
        End If
        
        eF_sigmabar = ((2 * Y3x) + (FOVTransitionEffort * (1 + uhat) + 2 * (Sqr((Y3x + FOVTransitionEffort) * (Y3x + (uhat * FOVTransitionEffort)))))) / (Y3x * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
        LabelResult_PredictedCollectionEffort_FOVS.Text = Format(eF_sigmabar, "0")
        
        If IsNumeric(LabelResult_PredictedCollectionEffort_FOVS.Value) And FOVTransitionEffort <> 0 Then
            LabelResult_PredictedCollectionEffort_FOVS.Enabled = True
            LabelResult_PredictedCollectionEffort_FOVS.BackColor = RGB(255, 255, 255)
        ElseIf IsNumeric(LabelResult_PredictedCollectionEffort_FOVS.Value) And FOVTransitionEffort = 0 Then
            LabelResult_PredictedCollectionEffort_FOVS.Text = ""
            ' FOVTransitionEffort is equal to 0, do not show
        Else
            LabelResult_PredictedCollectionEffort_FOVS.BackColor = RGB(224, 224, 224)
        End If
    Else
        ' FOVTransitionEffort is equal to 0, do not run calculation
    End If
      
    If Not IsNumeric(txt_N3E.Value) Then
        MsgBox "Please enter the amount of observed fields-of-view seen in the extrapolation count [N3E].", vbExclamation, "Input Required"
        txt_N3E.SetFocus
        Exit Sub
    End If
     
    Dim sigma_Fx As Double ' Total target concentration standard error with FOVS method
    sigma_Fx = 100 * Sqr((((s1 / Y1) / Sqr(N1)) ^ 2) + ((s3P / Sqr(N3C)) ^ 2) + (Sqr(N) / N) ^ 2)
    LabelResult_ConcentrationStandardError_FOVS.Text = Format(sigma_Fx, "0.00")
    
    If IsNumeric(LabelResult_ConcentrationStandardError_FOVS.Value) Then
        LabelResult_ConcentrationStandardError_FOVS.Enabled = True
        LabelResult_ConcentrationStandardError_FOVS.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_ConcentrationStandardError_FOVS.BackColor = RGB(224, 224, 224)
    End If
    
    OutputsSaved = True
    
    ' Enable ability to save variables.
'    CommandButton_SaveVariables_FOVS.Enabled = True
    
    ' Check if targets ended up being more common, and ask user if they would like to switch to the appropriated method.
    If N > xhat And Not MethodSwitchIgnored Then
        response = MsgBox("Warning: To enhance data collection efficiency based on your target-to-marker ratio (u-hat), consider focusing calibration counts on markers. Would you like to make this change?", vbQuestion + vbYesNo, "Most common specimens?")
            
        ' Check user response
        If response = vbYes Then
            FOVSTargetChosen = False
            FOVSMarkerChosen = True
            CalculatorFOVSMarker.Show ' FOVS calculator with equations that consider markers [n] being more common than targets [x].
            Hide
        Else
            'Do nothing, but show option to switch as a button.
            MethodSwitchIgnored = True 'TODO: Probably remove later as button is no longer present.
        End If
    Else
    End If
   
End Sub

Private Sub CommandButton_CountingEffort_Click()
    FOVSTargetChosen = True
    CountingEffort.Show
End Sub

Private Sub txt_N_FOVS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_N_FOVS.Text) > 0 Then
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

Private Sub txt_N3E_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_N3E.Text) > 0 Then
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

Private Sub txt_N_FOVS_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_N3E_FOVS_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

Private Sub txt_LevelError_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    AvoidCopyPaste KeyCode, Shift
End Sub

'
' Shutdown
'
    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) ' Once UserForm is closed, close all open windows and clear variables.
        
        ' Ask user if they want to close the form without saving.
        ' For this warning to appear, check if inputs were saved, if input boxes contain values, and that these values are different from those stored in memory.
        
        If CloseMode = 0 Then
            Unload AssistantCounting
            If InputsSaved And SavedVariablesFOVSTargetExists And InfoExported Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("Would you like to export your data one last time before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_FOVS_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            ElseIf InputsSaved And SavedVariablesFOVSTargetExists And Not InfoExported Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("There are data in the 'Exported data (FOVS-T) spreadsheet from previous trials. Would you like to export your data here before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_FOVS_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            ElseIf Not (InputsSaved Or SavedMarkerDetails Or CalibratedFOV) And SavedVariablesFOVSTargetExists Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("There are variables that differ from those in the latest export. Would you like to export these before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_FOVS_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            ElseIf Not SavedVariablesFOVSTargetExists And OutputsSaved Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("Would you like to export the saved information to a spreadsheet before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_FOVS_Click 'Run subroutine to export data.
                Else
                    Cancel = 0
                    End ' Terminates application and erases all data from memory.
                End If
            ElseIf (txt_N_FOVS.Value <> "" Or txt_N3E.Value <> "" Or txt_LevelError.Value <> "") Or (SavedMarkerDetails Or CalibratedFOV) Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("Stored variables will be deleted if the application is closed. Would you like to export these first?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
'                    CommandButton_SaveVariables_FOVS.Enabled = True
'                    CommandButton_SaveVariables_FOVS_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            Else
                Cancel = 0
                End
            End If
        End If
    End Sub
