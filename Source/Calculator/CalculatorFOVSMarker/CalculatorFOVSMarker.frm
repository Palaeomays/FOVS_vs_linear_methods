VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalculatorFOVSMarker 
   Caption         =   "Absolute abundance calculator v1.0 - FOVS method"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "CalculatorFOVSMarker.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CalculatorFOVSMarker"
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
        txt_X_FOVS.Text = ""
        txt_N3f.Text = ""
    Else
    ' User cancelled, do nothing
    End If
End Sub

Private Sub CommandButton_Assistant_Click()
    OriginFOVSMarker = True
    AssistantCounting.Show
End Sub

Private Sub CommandButton_CalibrationFOV_Click()
    FOVSMarkerChosen = True
    CalibratorFOV.Show
End Sub

Private Sub CommandButton_Clear_FOVS_Click()
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    ' Check user's response
    If response = vbYes Then
        txt_X_FOVS.Text = ""
        txt_N3f.Text = ""
        txt_LevelError.Text = ""
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
        
        txt_X_FOVS.Text = ""
        X = Empty
        
        txt_N3f.Text = ""
        N3F = Empty
        
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
        sigma_Fn = Empty
        LabelResult_ConcentrationStandardError_FOVS.Enabled = False
        LabelResult_ConcentrationStandardError_FOVS.BackColor = RGB(224, 224, 224)
        
        ' Optimal field of view counts
        
        LabelResult_OptimalCalibrationFOV.Text = ""
        Nstar3C = Empty
        LabelResult_OptimalCalibrationFOV.Enabled = False
        LabelResult_OptimalCalibrationFOV.BackColor = RGB(224, 224, 224)
        
        LabelResult_OptimalFullFOV.Text = ""
        Nstar3F = Empty
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
        
        LabelResult_Y3n.Text = ""
        Y3n = Empty
        LabelResult_Y3n.Enabled = False
        LabelResult_Y3n.BackColor = RGB(224, 224, 224)
        
        LabelResult_uhat_FOVS.Text = ""
        uhat = Empty
        LabelResult_uhat_FOVS.Enabled = False
        LabelResult_uhat_FOVS.BackColor = RGB(224, 224, 224)
        
        LabelResult_FOVTransitionEffort.Text = ""
        FOVTransitionEffort = Empty
        LabelResult_FOVTransitionEffort.Enabled = False
        LabelResult_FOVTransitionEffort.BackColor = RGB(224, 224, 224)
        
        LabelResult_nhat.Text = ""
        nhat = Empty
        LabelResult_nhat.Enabled = False
        LabelResult_nhat.BackColor = RGB(224, 224, 224)
        
        ' Field of view calibration count
        
        N = Empty
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

Private Sub CommandButton_FocusTargets_Click()
    response = MsgBox("Are you sure you want to change the focus to targets [x]?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    ' Check user's response
    If response = vbYes Then
        CalculatorFOVSTarget.Show
        FOVSMarkerChosen = False
        FOVSTargetChosen = True
        Hide
    Else
    ' User cancelled, do nothing
    End If
End Sub

Private Sub CommandButton_MarkerCharacteristics_Click()
    FOVSMarkerChosen = True
    MarkerCharacteristics.Show
End Sub

Private Sub CommandButton_SaveVariables_FOVS_Click()
    ' Validate other input fields to see if not empty.
    If Not IsNumeric(txt_X_FOVS.Value) And Not ShutdownRequested Then
        MsgBox "Please enter the number of targets counted [x].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not IsNumeric(txt_N3f.Value) And Not ShutdownRequested Then
        MsgBox "Please enter the amount of observed fields-of-view seen in the full count [N3f].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not SavedMarkerDetails And Not ShutdownRequested Then
        MsgBox "Please enter marker characteristics.", vbExclamation, "Input Required"
        MarkerCharacteristics.Show
        Exit Sub
    End If
    
    If Not CalibratedFOV And Not ShutdownRequested Then
        MsgBox "Please attempt a FOV calibration.", vbExclamation, "Input Required"
        MarkerCharacteristics.Show
        Exit Sub
    End If
    
    ' Make sure calculations are run first
    If Not ShutdownRequested Then
        CommandButton_Calculate_FOVS_Click
    End If
    
    ' Initialize the variable to False
    SavedFOVSMarkerExists = False
    
    ' Create a new worksheet named "Saved Variables (FOVS_Marker)"
    ' Check if the sheet "Saved Variables (FOVS)" already exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Exported data (FOVS-M)" Then
            SavedVariablesFOVSMarkerExists = True
            Set SavedVariablesFOVSMarker = ws
            Exit For
        End If
    Next ws

    ' If the sheet doesn't exist, create a new one
    If Not SavedVariablesFOVSMarkerExists Then
        Set SavedVariablesFOVSMarker = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Calculator"))
        SavedVariablesFOVSMarker.Name = "Saved Variables (FOVS_Marker)"
        AddHeadersFOVSMarker SavedVariablesFOVSMarker
        
        ' Clear nextRow and lastNonEmptyRow
        nextRow = 1
        lastNonEmptyRow = 0
    End If
    
    ' Determine the next empty row by examining all used columns
    lastNonEmptyRow = 0
    For i = 1 To 29 ' Data is up to column AC
        nextRow = SavedVariablesFOVSMarker.Cells(Rows.Count, i).End(xlUp).Row
        If nextRow > lastNonEmptyRow Then
            lastNonEmptyRow = nextRow
        End If
    Next i
    nextRow = lastNonEmptyRow + 1
   
    ' Write values from the userform to specific cells in the next available row
    SavedVariablesFOVSMarker.Cells(nextRow, "A").Value = Now
    SavedVariablesFOVSMarker.Cells(nextRow, "B").Value = lastNonEmptyRow
    SavedVariablesFOVSMarker.Cells(nextRow, "C").Value = txt_SampleName.Text
    SavedVariablesFOVSMarker.Cells(nextRow, "D").Value = txt_X_FOVS.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "E").Value = txt_N3f.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "F").Value = txt_LevelError.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "G").Value = MarkerCharacteristics.txt_N1.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "H").Value = MarkerCharacteristics.txt_Y1.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "I").Value = MarkerCharacteristics.txt_s1.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "J").Value = MarkerCharacteristics.txt_N2.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "K").Value = MarkerCharacteristics.txt_Y2.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "L").Value = MarkerCharacteristics.ComboBox_Units.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "M").Value = MarkerCharacteristics.txt_s2.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "N").Value = LabelResult_Concentration_FOVS.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "O").Value = LabelResult_ConcentrationStandardError_FOVS.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "P").Value = LabelResult_OptimalCalibrationFOV.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "Q").Value = LabelResult_OptimalFullFOV.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "R").Value = LabelResult_OptimalRatioFOV.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "S").Value = LabelResult_CollectionEffort_FOVS.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "T").Value = LabelResult_PredictedCollectionEffort_FOVS.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "U").Value = LabelResult_Y3n.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "V").Value = LabelResult_uhat_FOVS.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "W").Value = LabelResult_FOVTransitionEffort.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "X").Value = LabelResult_nhat.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "Y").Value = CalibratorFOV.txt_X_FOVS.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "Z").Value = CalibratorFOV.txt_N3c.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "AA").Value = CalibratorFOV.txt_S3.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "AB").Value = CountingEffort.txtTimeFOV.Value
    SavedVariablesFOVSMarker.Cells(nextRow, "AC").Value = CountingEffort.txtTimeTotal.Value
    
    ' Inform the user that values have been saved
    InfoExported = True
    MsgBox "Values have been saved to the worksheet 'Exported data (FOVS-M)'.", vbInformation
    
    If ShutdownRequested Then
        End
    End If
End Sub

Private Sub AddHeadersFOVSMarker(ByRef ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "Date and time (DD/MM/YYYY XX:XX)"
        .Cells(1, 2).Value = "Data export #"
        .Cells(1, 3).Value = "Sample name"
        .Cells(1, 4).Value = "Number of target specimens from full counts [x]"
        .Cells(1, 5).Value = "Number of fields of view counted [N3F]"
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
        .Cells(1, 17).Value = "Optimal number of full-count FOVs [Nstar3F]"
        .Cells(1, 18).Value = "Optimal FOV count ratio (full-to-calibration ratio) [deltastar]"
        .Cells(1, 19).Value = "Present data collection effort (time units) [eF]"
        .Cells(1, 20).Value = "Predicted data collection effort to achieve desired error rate (time units) [eF-sigma-bar]"
        .Cells(1, 21).Value = "Mean number of markers per field of view [Yline3n]"
        .Cells(1, 22).Value = "Target-to-market ratio [u-hat]"
        .Cells(1, 23).Value = "Field of view transition effort factor [omegaline]"
        .Cells(1, 24).Value = "Estimate of marker specimens extrapolated from full counts [nhat]"
        .Cells(1, 25).Value = "Number of counted marker specimens during calibration counts [n]"
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
    MsgBox "The FOVS method requires a series of 'calibration counts' followed by a series of 'full counts'. To insert the calibration count data, press the 'field of view (FOV) calibration count' button." & vbNewLine & vbNewLine & "Once these are filled, the 'full count' fields will be available." & vbNewLine & vbNewLine & "Absolute abundances (and associated error) will require the addition of data from the marker specimens being used. To do so, press the 'marker and sample characteristrics' button." & vbNewLine & vbNewLine & "(Optional: To predict the amount of sampling effort required for a given assemblage, insert the relevant data by pressing the 'optimisation data' button.)", vbInformation
    
    ' Check if certain sheets are present. Iterate through all worksheets in the workbook.
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Exported data (FOVS-M)" Then
            ' Set the flag to True if the worksheet exists
            SavedVariablesFOVSMarkerExists = True
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
    
    If CalibratedFOV Then
        txt_X_FOVS.Enabled = True
        txt_X_FOVS.BackColor = RGB(255, 255, 255)
        
        txt_N3f.Enabled = True
        txt_N3f.BackColor = RGB(255, 255, 255)
    Else
        txt_X_FOVS.BackColor = RGB(224, 224, 224)
        txt_N3f.BackColor = RGB(224, 224, 224)
    End If
              
    If Nstar3C <> 0 Then ' If Nstar3C is not equal to 0, render it in the label.
        LabelResult_OptimalCalibrationFOV.Enabled = True
        LabelResult_OptimalCalibrationFOV.Text = Format(Nstar3C, "0.00")
        LabelResult_OptimalCalibrationFOV.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_OptimalCalibrationFOV.BackColor = RGB(224, 224, 224)
    End If
    
    If Nstar3F <> 0 Then ' If Nstar3F is not equal to 0, render it in the label.
        LabelResult_OptimalFullFOV.Enabled = True
        LabelResult_OptimalFullFOV.Text = Format(Nstar3F, "0.00")
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
        
    If Y3n <> 0 Then ' If Y3n is not equal to 0, render it in the label.
        LabelResult_Y3n.Enabled = True
        LabelResult_Y3n.Text = Format(Y3x, "0.000")
        LabelResult_Y3n.BackColor = RGB(255, 255, 255)
    Else ' If uhat is equal to 0, do nothing.
        LabelResult_Y3n.BackColor = RGB(224, 224, 224)
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
    
'    If nhat <> 0 Then ' If nhat is not equal to 0, render it in the label.
'        LabelResult_nhat.Enabled = True
'        LabelResult_nhat.Text = Format(nhat, "0")
'        LabelResult_nhat.BackColor = RGB(255, 255, 255)
'    Else ' If uhat is equal to 0, do nothing.
'        LabelResult_nhat.BackColor = RGB(224, 224, 224)
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
End Sub

' Check for changes in inputs.

Private Sub txt_X_FOVS_Change()
    If IsNumeric(txt_X_FOVS.Value) And txt_X_FOVS.Value <> X Then
        InputsSaved = False
    End If
End Sub

Private Sub txt_N3f_Change()
    If IsNumeric(txt_N3f.Value) And txt_N3f.Value <> N3F Then
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
    If Not CalibratedFOV Or N3C = 0 Then
        MsgBox "Please attempt a FOV calibration.", vbExclamation, "Input Required"
        CalibratorFOV.Show
        Exit Sub
    End If
    
    If Not IsNumeric(txt_X_FOVS.Value) Then
        MsgBox "Please enter the number of targets counted in the full counts [x].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not IsNumeric(txt_N3f.Value) Then
        MsgBox "Please enter the amount of observed fields-of-view seen in the full count [N3F].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not IsNumeric(txt_LevelError.Value) Then
        MsgBox "Please enter the desired target level of error as a percentage (e.g., 10).", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not SavedMarkerDetails Then
        MsgBox "Please enter marker and sample characteristics.", vbExclamation, "Input Required"
        MarkerCharacteristics.Show
        Exit Sub
    End If
    
    ' Store values in memory
    
    If Len(txt_SampleName.Text) > 1 Then
        SampleName = txt_SampleName.Text
    Else
    End If
    
    X = CLng(txt_X_FOVS.Value)
    N3F = CLng(txt_N3f.Value)
    LevelError = CDbl(txt_LevelError.Value) 'TODO Can lead to negatives later on if less than 5.
    
    If X <= 0 Then
        MsgBox "Number of targets needs to be higher than 0.", vbExclamation
        Exit Sub
    End If
    
    InputsSaved = True
    
   ' Perform background calculations
    Y3n = N / N3C
    
    nhat = N3F * Y3n
    
    'Dim Vline As Double ' Total mass or volume of samples ' TODO Include as in Linear?
    'Vline = N2 * Y2
   
    Dim c4 As Double ' Bias correction for calibration count no. of FOVs (N3)
    c4 = Sqr(2 / (N3C - 1)) * WorksheetFunction.Gamma(N3C / 2) / WorksheetFunction.Gamma((N3C - 1) / 2)
    
    Dim sigmahat3 As Double ' Unbiased estimator for the population standard deviation
    sigmahat3 = s3 / c4
   
    Dim s3P As Double ' Proportional corrected sample standard deviation - common grains/FOV
    s3P = (sigmahat3 / Y3n)

    ' Variable defined as public in ShowCalculator module
    uhat = X / nhat ' Marker-specific
 
    ' Perform visible calculations
    Dim c As Double ' Mean number of target specimens per unit mass or volume
    ' C = (X * Y1 * N1) / (nhat * Vline) ' TODO Include as in Linear? Going with below code.
    c = (X * Y1 * N1) / (nhat * Y2)
    
    Dim sigma_Fn As Double ' Total target concentration standard error with FOVS method
    sigma_Fn = 100 * Sqr((((s1 / Y1) / Sqr(N1)) ^ 2) + ((Sqr(X) / X) ^ 2) + (s3P / Sqr(N3C)) ^ 2)
    
    If FOVTransitionEffort <> 0 Then
        Nstar3C = (1 / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3n + FOVTransitionEffort) + (Sqr(Y3n + (FOVTransitionEffort / uhat)))) / ((Y3n * (Sqr(Y3n + FOVTransitionEffort)))) 'TODO condition if LevelError is 0
        Nstar3F = ((1 / uhat) / (((LevelError / 100) * (LevelError / 100)) - ((s1 / Y1) * (s1 / Y1) / N1))) * (Sqr(Y3n + FOVTransitionEffort) + Sqr(Y3n + (FOVTransitionEffort / uhat))) / (Y3n * (Sqr(Y3n + (FOVTransitionEffort / uhat)))) 'TODO condition if LevelError is 0
        deltastar = uhat * Sqr((FOVTransitionEffort + (Y3n * uhat)) / ((FOVTransitionEffort * uhat) + (Y3n * uhat)))
        eF = (FOVTransitionEffort * N3C) + X + (FOVTransitionEffort * N3F) + N
        eF_sigmabar = ((2 * (Y3n * uhat)) + (FOVTransitionEffort * (1 + uhat) + 2 * (Sqr(((Y3n * uhat) + FOVTransitionEffort) * ((Y3n * uhat) + (uhat * FOVTransitionEffort)))))) / ((Y3n * uhat) * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
    
    Else
        ' FOVTransitionEffort is equal to 0, do not run calculation
    End If
    
    ' Display calculated results
    LabelResult_Concentration_FOVS = Format(c, "0")
    txt_ConcentrationUnits = SizeUnit
    
    LabelResult_ConcentrationStandardError_FOVS.Text = Format(sigma_Fn, "0.00")
    
    LabelResult_OptimalCalibrationFOV.Text = Format(Nstar3C, "0.00")
    LabelResult_OptimalFullFOV.Text = Format(Nstar3F, "0.00")
    LabelResult_OptimalRatioFOV.Text = Format(deltastar, "0.00")
    
    LabelResult_Y3n.Text = Format(Y3n, "0.00")
    LabelResult_uhat_FOVS.Text = Format(uhat, "0.000")
    LabelResult_FOVTransitionEffort.Text = Format(FOVTransitionEffort, "0.000")
    LabelResult_nhat.Text = Format(nhat, "0")
    LabelResult_CollectionEffort_FOVS.Text = Format(eF, "0")
    LabelResult_PredictedCollectionEffort_FOVS.Text = Format(eF_sigmabar, "0")

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
    
    If IsNumeric(LabelResult_ConcentrationStandardError_FOVS.Value) Then
        LabelResult_ConcentrationStandardError_FOVS.Enabled = True
        LabelResult_ConcentrationStandardError_FOVS.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_ConcentrationStandardError_FOVS.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_OptimalCalibrationFOV.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_OptimalCalibrationFOV.Enabled = True
        LabelResult_OptimalCalibrationFOV.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_OptimalCalibrationFOV.Value) And FOVTransitionEffort = 0 Then
        ' FOVTransitionEffort is equal to 0, do not show
        LabelResult_OptimalCalibrationFOV.Text = ""
    Else
        LabelResult_OptimalFullFOV.BackColor = RGB(224, 224, 224)
    End If
        
    If IsNumeric(LabelResult_OptimalFullFOV.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_OptimalFullFOV.Enabled = True
        LabelResult_OptimalFullFOV.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_OptimalFullFOV.Value) And FOVTransitionEffort = 0 Then
        ' FOVTransitionEffort is equal to 0, do not show
        LabelResult_OptimalFullFOV.Text = ""
    Else
        LabelResult_OptimalFullFOV.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_OptimalRatioFOV.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_OptimalRatioFOV.Enabled = True
        LabelResult_OptimalRatioFOV.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_OptimalRatioFOV.Value) And FOVTransitionEffort = 0 Then
        ' FOVTransitionEffort is equal to 0, do not show
        LabelResult_OptimalRatioFOV.Text = ""
    Else
        LabelResult_OptimalRatioFOV.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_Y3n.Value) Then
        LabelResult_Y3n.Enabled = True
        LabelResult_Y3n.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_Y3n.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_uhat_FOVS.Value) Then
        LabelResult_uhat_FOVS.Enabled = True
        LabelResult_uhat_FOVS.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_uhat_FOVS.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_FOVTransitionEffort.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_FOVTransitionEffort.Enabled = True
        LabelResult_FOVTransitionEffort.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_FOVTransitionEffort.Value) And FOVTransitionEffort = 0 Then
    ' FOVTransitionEffort is equal to 0, do not show
        LabelResult_FOVTransitionEffort.Text = ""
    Else
        LabelResult_FOVTransitionEffort.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_nhat.Value) Then
        LabelResult_nhat.Enabled = True
        LabelResult_nhat.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_LabelResult_nhat.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_CollectionEffort_FOVS.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_CollectionEffort_FOVS.Enabled = True
        LabelResult_CollectionEffort_FOVS.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_CollectionEffort_FOVS.Value) And FOVTransitionEffort = 0 Then
        ' FOVTransitionEffort is equal to 0, do not show
        LabelResult_CollectionEffort_FOVS.Text = ""
    Else
        LabelResult_PredictedCollectionEffort_Linear.BackColor = RGB(224, 224, 224)
    End If

    If IsNumeric(LabelResult_PredictedCollectionEffort_FOVS.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_PredictedCollectionEffort_FOVS.Enabled = True
        LabelResult_PredictedCollectionEffort_FOVS.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_PredictedCollectionEffort_FOVS.Value) And FOVTransitionEffort = 0 Then
        LabelResult_PredictedCollectionEffort_FOVS.Text = ""
        ' FOVTransitionEffort is equal to 0, do not show
    Else
        LabelResult_PredictedCollectionEffort_FOVS.BackColor = RGB(224, 224, 224)
    End If
    
    OutputsSaved = True
    
    ' Enable ability to save variables.
'    CommandButton_SaveVariables_FOVS.Enabled = True
    
    ' Check if targets ended up being more common, and ask user if they would like to switch to the appropriated method.
    If X > nhat And Not MethodSwitchIgnored Then
        response = MsgBox("Warning: To enhance data collection efficiency based on your target-to-marker ratio (u-hat), consider focusing calibration counts on targets. Would you like to make this change?", vbQuestion + vbYesNo, "Most common specimens?")
            
        ' Check user response
        If response = vbYes Then
            FOVSMarkerChosen = False
            FOVSTargetChosen = True
            CalculatorFOVSTarget.Show ' FOVS calculator with equations that consider targets [x] being more common than markers [n].
            Hide
        Else
            'Do nothing, but show option to switch as a button.
            MethodSwitchIgnored = True
            CommandButton_FocusTargets.Visible = True
        End If
    Else
    End If
End Sub

Private Sub CommandButton_CountingEffort_Click()
    FOVSMarkerChosen = True
    CountingEffort.Show
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

Private Sub txt_N3f_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_N3f.Text) > 0 Then
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

Private Sub txt_X_FOVS_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_N3f_FOVS_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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
' Shutdown
'
    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) ' Once UserForm is closed, close all open windows and clear variables.
        
        ' Ask user if they want to close the form without saving.
        ' For this warning to appear, check if inputs were saved, if input boxes contain values, and that these values are different from those stored in memory.

        If CloseMode = 0 Then
            Unload AssistantCounting
            If InputsSaved And SavedVariablesFOVSMarkerExists And InfoExported Then
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
            ElseIf InputsSaved And SavedVariablesFOVSMarkerExists And Not InfoExported Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("There are data in the 'Exported data (FOVS-M) spreadsheet from previous trials. Would you like to export your data here before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_FOVS_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            ElseIf Not (InputsSaved Or SavedMarkerDetails Or CalibratedFOV) And SavedVariablesFOVSMarkerExists Then
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
            ElseIf Not SavedVariablesFOVSMarkerExists And OutputsSaved Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("Would you like to export the saved information to a spreadsheet before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_FOVS_Click 'Run subroutine to export data.
                Else
                    Cancel = 0
                    End ' Terminates application and erases all data from memory.
                End If
            ElseIf (txt_X_FOVS.Value <> "" Or txt_N3f.Value <> "" Or txt_LevelError.Value <> "") Or (SavedMarkerDetails Or CalibratedFOV) Then
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
