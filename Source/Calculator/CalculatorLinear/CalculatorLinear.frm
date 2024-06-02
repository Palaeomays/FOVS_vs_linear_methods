VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalculatorLinear 
   Caption         =   "Absolute abundance calculator v1.0 - Linear method"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "CalculatorLinear.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CalculatorLinear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Variables
'

    Dim ConfidenceInterval As Double
    Private InputsSaved As Boolean
    Private OutputsSaved As Boolean
    Private InfoExported As Boolean
    

Private Sub CommandButton_Assistant_Click()
    OriginLinear = True
    AssistantCounting.Show
End Sub

Private Sub CommandButton_Clear_Linear_Click()
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    ' Check user's response
    If response = vbYes Then
        txt_X.Text = ""
        txt_N.Text = ""
        txt_ConfidenceInterval.Text = ""
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
        
        txt_X.Text = ""
        X = Empty
        
        txt_N.Text = ""
        N = Empty
              
        txt_ConfidenceInterval.Text = ""
        ConfidenceInterval = Empty
        
        txt_LevelError.Text = ""
        LevelError = Empty
        
        LabelResult_Concentration_Linear.Text = ""
        c = Empty
        LabelResult_Concentration_Linear.Enabled = False
        LabelResult_Concentration_Linear.BackColor = RGB(224, 224, 224)
        
        txt_ConcentrationUnits.Text = ""
        UnitSize = Empty
        txt_ConcentrationUnits.Enabled = False
        txt_ConcentrationUnits.BackColor = RGB(224, 224, 224)
        
        LabelResult_ConcentrationStandardError_Linear.Text = ""
        sigma_L = Empty
        LabelResult_ConcentrationStandardError_Linear.Enabled = False
        LabelResult_ConcentrationStandardError_Linear.BackColor = RGB(224, 224, 224)
        
        LabelResult_ConcentrationMax_Linear.Text = ""
        CI_max = Empty
        txt_CImaxUnits.Text = ""
        txt_CImaxUnits.Enabled = False
        txt_CImaxUnits.BackColor = RGB(224, 224, 224)
                
        LabelResult_ConcentrationMin_Linear.Text = ""
        CI_min = Empty
        txt_CIminUnits = ""
        txt_CImaxUnits.Enabled = False
        txt_CImaxUnits.BackColor = RGB(224, 224, 224)
        
        LabelResult_uhat_Linear.Text = ""
        uhat = Empty
        LabelResult_uhat_Linear.Enabled = False
        LabelResult_uhat_Linear.BackColor = RGB(224, 224, 224)
        
        LabelResult_CollectionEffort_Linear.Text = ""
        eL = Empty
        LabelResult_CollectionEffort_Linear.Enabled = False
        LabelResult_CollectionEffort_Linear.BackColor = RGB(224, 224, 224)
        
        LabelResult_PredictedCollectionEffort_Linear.Text = ""
        eL_sigmabar = Empty
        LabelResult_PredictedCollectionEffort_Linear.Enabled = False
        LabelResult_PredictedCollectionEffort_Linear.BackColor = RGB(224, 224, 224)
        
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
        
        N3C = Empty
        TimeFOV = Empty
        TimeTotal = Empty
        
        CountingEffortCalibration = False
        CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
        
        'Unload associated user forms
        
        ClearedAllData = True
        
        Unload MarkerCharacteristics
        Unload CountingEffort
        
        ' Disallow exporting data
        
        'If CommandButton_SaveVariables_Linear.Enabled = True Then
        '    CommandButton_SaveVariables_Linear.Enabled = False
        'End If
        
    Else
    ' User cancelled, do nothing
    End If
End Sub

Private Sub CommandButton_MethodDetermination_Click()
    LinearChosen = True
    CalculatorStart.Show
    Me.Hide
End Sub

Private Sub UserForm_Initialize() ' Runs as soon as userform is opened

    ' Check if certain sheets are present. Iterate through all worksheets in the workbook.
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Exported data (Linear)" Then
            ' Set the flag to True if the worksheet exists
            SavedVariablesLinearExists = True
            Exit For
        End If
    Next ws

    ' Populate inputs with previous values if these exist. Sets the value of the textbox to the value of the public variable.
    If Len(SampleName) > 1 Then
        txt_SampleName.Text = SampleName
    Else
    End If
        
    
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
    
    If LevelError <> 0 Then
        txt_LevelError = LevelError
    Else
        txt_LevelError.Text = ""
    End If
           
    If uhat <> 0 Then ' If uhat is not equal to 0, render it in the label.
        LabelResult_uhat_Linear.Enabled = True
        LabelResult_uhat_Linear.Text = Format(uhat, "0.000")
        LabelResult_uhat_Linear.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    If eL <> 0 Then
        LabelResult_CollectionEffort_Linear.Enabled = True
        LabelResult_CollectionEffort_Linear.Text = Format(eL, "0.000")
        LabelResult_CollectionEffort_Linear.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    If eL_sigmabar <> 0 Then
        LabelResult_PredictedCollectionEffort_Linear.Enabled = True
        LabelResult_PredictedCollectionEffort_Linear.Text = Format(eL_sigmabar, "0.000")
        LabelResult_PredictedCollectionEffort_Linear.BackColor = RGB(255, 255, 255)
    Else
    End If
    
    ' Check if marker details were saved. Changes color from red to green.
            
    If Not SavedMarkerDetails Then
        CommandButton_MarkerCharacteristics.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data missing)"
    Else
        CommandButton_MarkerCharacteristics.BackColor = RGB(212, 236, 214) ' Greenish color
        CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data ready)"
    End If
    
    ' Check if counting effort calibration was done. Changes color from red to green.
            
    If Not CountingEffortCalibration Or (CountingEffortCalibration And (X = 0 Or N = 0)) Then
        CommandButton_CountingEffort.BackColor = RGB(245, 148, 146) ' Reddish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data missing)"
    Else
        CommandButton_CountingEffort.BackColor = RGB(212, 236, 214) ' Greenish color
        CommandButton_CountingEffort.Caption = "Optional: Optimisation data" & vbCrLf & "(data ready)"
    End If
    
End Sub

' Check for changes in inputs.

Private Sub txt_X_Change()
    If IsNumeric(txt_X.Value) And txt_X.Value <> X Then
        InputsSaved = False
    End If
End Sub

Private Sub txt_N_Change()
    If IsNumeric(txt_N.Value) And txt_N.Value <> N Then
        InputsSaved = False
    End If
End Sub

Private Sub txt_ConfidenceInterval_Change()
    If IsNumeric(txt_ConfidenceInterval) And txt_ConfidenceInterval <> ConfidenceInterval Then
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

Private Sub CommandButton_Calculate_Linear_Click()
    ' Validate other input fields to see if not empty.
    If Not IsNumeric(txt_X.Value) Then
        MsgBox "Please enter the number of targets counted [x].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not IsNumeric(txt_N.Value) Then
        MsgBox "Please enter the number of markers counted [n].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not IsNumeric(txt_ConfidenceInterval.Value) Then
        MsgBox "Please enter the desired confidence interval as a percentage (e.g., 95).", vbExclamation, "Input Required"
        Exit Sub
    End If

    If Not IsNumeric(txt_LevelError.Value) Then
        MsgBox "Please enter the desired target level of error as a percentage (e.g., 10).", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not SavedMarkerDetails Then
        MsgBox "Please enter marker characteristics.", vbExclamation, "Input Required"
        LinearChosen = True
        MarkerCharacteristics.Show
        Exit Sub
    End If
    
    ' Store values in memory as doubles
    If Len(txt_SampleName.Text) > 1 Then
        SampleName = txt_SampleName.Text
    Else
    End If
    
    X = CLng(txt_X.Value)
    N = CLng(txt_N.Value)
    ConfidenceInterval = CDbl(txt_ConfidenceInterval.Value)
    LevelError = CDbl(txt_LevelError.Value) 'TODO Can lead to negatives later on if less than 5.
    
    InputsSaved = True

    ' Avoid zeros in counts
    If X <= 0 Then
        MsgBox "Number of targets needs to be higher than 0.", vbExclamation
        Exit Sub
    End If
    
    If N <= 0 Then
        MsgBox "Number of markers needs to be higher than 0.", vbExclamation
        Exit Sub
    End If
    
    ' Avoid inversion of values for confidence intervals lower than 25 and infinity by values equal or higher than 100.
    
    If ConfidenceInterval < 25 Then
        MsgBox "Please insert a confidence interval higher than 25.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If ConfidenceInterval >= 100 Then
        MsgBox "Please insert a confidence interval equal or lower than 100.", vbExclamation, "Input Required"
        Exit Sub
    End If

    ' Perform background calculations
    Dim Vline As Double ' Total mass or volume of samples
    Vline = N2 * Y2
    
    uhat = X / N
    
    Dim mline As Double ' Estimated number of exotic markers added to the sample
    mline = N1 * Y1
    
    Dim ConfidenceInterval_Percentile As Double ' MAYS
    ConfidenceInterval_Percentile = (Sqr(ConfidenceInterval) / 10)
    
    Dim CL As Double ' MAYS
    CL = Round(ConfidenceInterval_Percentile, 3)
    
    Dim Zscore As Double ' Distance in standard deviations of an observed value from the mean.
    Zscore = WorksheetFunction.NormInv(CL, 0, 1)

    Dim uhat_max As Double ' MAYS
    uhat_max = (uhat + (1 / (2 * N)) + Sqr(uhat * (1 + uhat) / N + (1 / (4 * N * N)))) / (1 - (1 / N))
    
    Dim uhat_min As Double ' MAYS
    uhat_min = (uhat + (1 / (2 * N)) - Sqr(uhat * (1 + uhat) / N + (1 / (4 * N * N)))) / (1 - (1 / N))
    
    Dim s_log_uhat As Double ' MAYS
    s_log_uhat = (Application.WorksheetFunction.Log(CDbl(uhat_max)) - Application.WorksheetFunction.Log(CDbl(uhat_min))) / 2

    Dim sm As Double ' MAYS
    sm = Sqr(N1) * s1
    
    Dim sv As Double ' MAYS
    sv = Sqr(N2) * s2
    
    Dim alpha As Double ' MAYS
    alpha = Atn((mline / sm) / (Vline / sv))
    
    Dim beta As Double ' MAYS
    beta = WorksheetFunction.Asin(1 / Sqr((mline / sm) * (mline / sm) + (Vline / sv) * (Vline / sv)))
        
    Dim MV_max As Double ' MAYS
    MV_max = sm * Tan(alpha + beta) / sv

    Dim MV_min As Double ' MAYS
    MV_min = sm * Tan(alpha - beta) / sv
    
    Dim s_log_mv As Double ' MAYS
    s_log_mv = (Application.WorksheetFunction.Log(CDbl(MV_max)) - Application.WorksheetFunction.Log(CDbl(MV_min))) / 2
    
    Dim logF As Double '
    logF = Zscore * Sqr((s_log_uhat * s_log_uhat) + (s_log_mv * s_log_mv))
    
    Dim F As Double ' MAYS
    F = 10 ^ logF
       
    ' Perform visible calculations
    Dim c As Double ' Mean number of target specimens per unit mass or volume
    c = (X * Y1 * N1) / (N * Vline)
    
    Dim sigma_L As Double ' Total concentration standard error for linear method
    sigma_L = 100 * (Sqr((((s1 / Y1) / (Sqr(N1))) ^ 2) + ((Sqr(X) / X) ^ 2) + ((Sqr(N) / N) ^ 2)))
    
    Dim CI_max As Double ' Predicted maximum concentation of targets
    CI_max = uhat * mline * F / Vline
    
    Dim CI_min As Double ' Predicted minimum concentation of targets
    CI_min = uhat * mline / (Vline * F)
    
    If FOVTransitionEffort <> 0 Then
        eL = ((FOVTransitionEffort * (X / Y3x)) + X + N)
    Else
        ' FOVTransitionEffort is equal to 0, do not run calculation
    End If
    
    If FOVTransitionEffort <> 0 Then
        eL_sigmabar = (FOVTransitionEffort * (1 + uhat) + (Y3x * (2 + uhat)) + Y3x / uhat) / (Y3x * ((LevelError / 100) * (LevelError / 100) - ((s1 / Y1) * (s1 / Y1) / N1)))
    Else
        ' FOVTransitionEffort is equal to 0, do not run calculation
    End If
    
    ' Display calculated results
    LabelResult_Concentration_Linear = Format(c, "0")
    txt_ConcentrationUnits = SizeUnit
    
    LabelResult_ConcentrationStandardError_Linear.Text = Format(sigma_L, "0.00")
    
    LabelResult_ConcentrationMax_Linear.Text = Format(CI_max, "0")
    txt_CImaxUnits = SizeUnit
    
    LabelResult_ConcentrationMin_Linear.Text = Format(CI_min, "0")
    txt_CIminUnits = SizeUnit
    
    LabelResult_uhat_Linear.Text = Format(uhat, "0.000")
    
    LabelResult_CollectionEffort_Linear.Text = Format(eL, "0")
    
    LabelResult_PredictedCollectionEffort_Linear.Text = Format(eL_sigmabar, "0") 'TODO If on estimating effort targets are more common than markers, and then markets are shown to be greater than targets in the actual count, predicted time is lower than actual ti
    
    'Enable and colour output backgrounds to white
    If IsNumeric(LabelResult_Concentration_Linear.Value) Then
        LabelResult_Concentration_Linear.Enabled = True
        LabelResult_Concentration_Linear.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_Concentration_Linear.BackColor = RGB(224, 224, 224)
    End If
    
    If Len(txt_ConcentrationUnits) > 0 Then
        txt_ConcentrationUnits.Enabled = True
        txt_ConcentrationUnits.BackColor = RGB(255, 255, 255)
    Else
        txt_ConcentrationUnits.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_ConcentrationStandardError_Linear.Value) Then
        LabelResult_ConcentrationStandardError_Linear.Enabled = True
        LabelResult_ConcentrationStandardError_Linear.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_ConcentrationStandardError_Linear.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_ConcentrationMax_Linear.Value) Then
        LabelResult_ConcentrationMax_Linear.Enabled = True
        LabelResult_ConcentrationMax_Linear.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_ConcentrationMax_Linear.BackColor = RGB(224, 224, 224)
    End If
    
    If Len(txt_CImaxUnits) > 0 Then
        txt_CImaxUnits.Enabled = True
        txt_CImaxUnits.BackColor = RGB(255, 255, 255)
    Else
        txt_CImaxUnits.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_ConcentrationMin_Linear.Value) Then
        LabelResult_ConcentrationMin_Linear.Enabled = True
        LabelResult_ConcentrationMin_Linear.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_ConcentrationMin_Linear.BackColor = RGB(224, 224, 224)
    End If
    
    If Len(txt_CIminUnits) > 0 Then
        txt_CIminUnits.Enabled = True
        txt_CIminUnits.BackColor = RGB(255, 255, 255)
    Else
        txt_CIminUnits.BackColor = RGB(224, 224, 224)
    End If
    
    If IsNumeric(LabelResult_uhat_Linear.Value) Then
        LabelResult_uhat_Linear.Enabled = True
        LabelResult_uhat_Linear.BackColor = RGB(255, 255, 255)
    Else
        LabelResult_uhat_Linear.BackColor = RGB(224, 224, 224)
    End If

    If IsNumeric(LabelResult_CollectionEffort_Linear.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_CollectionEffort_Linear.Enabled = True
        LabelResult_CollectionEffort_Linear.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_CollectionEffort_Linear.Value) And FOVTransitionEffort = 0 Then
        ' FOVTransitionEffort is equal to 0, do not show
        LabelResult_CollectionEffort_Linear.Text = ""
    Else
        LabelResult_PredictedCollectionEffort_Linear.BackColor = RGB(224, 224, 224)
    End If

    If IsNumeric(LabelResult_PredictedCollectionEffort_Linear.Value) And FOVTransitionEffort <> 0 Then
        LabelResult_PredictedCollectionEffort_Linear.Enabled = True
        LabelResult_PredictedCollectionEffort_Linear.BackColor = RGB(255, 255, 255)
    ElseIf IsNumeric(LabelResult_PredictedCollectionEffort_Linear.Value) And FOVTransitionEffort = 0 Then
        LabelResult_PredictedCollectionEffort_Linear.Text = ""
        ' FOVTransitionEffort is equal to 0, do not show
    Else
        LabelResult_PredictedCollectionEffort_Linear.BackColor = RGB(224, 224, 224)
    End If
    
    OutputsSaved = True
    
    ' Enable ability to save variables.
    'CommandButton_SaveVariables_Linear.Enabled = True
    
End Sub

Private Sub CommandButton_CountingEffort_Click()
    LinearChosen = True
    CountingEffort.Show
End Sub

Private Sub CommandButton_SaveVariables_Linear_Click()
    ' Validate other input fields to see if not empty.
    If Not IsNumeric(txt_X.Value) And Not ShutdownRequested Then
        MsgBox "Please enter the number of targets counted [x].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not IsNumeric(txt_N.Value) And Not ShutdownRequested Then
        MsgBox "Please enter the number of markers counted [n].", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not IsNumeric(txt_ConfidenceInterval.Value) And Not ShutdownRequested Then
        MsgBox "Please enter the desired confidence interval as a percentage (e.g., 95).", vbExclamation, "Input Required"
        Exit Sub
    End If

    If Not IsNumeric(txt_LevelError.Value) And Not ShutdownRequested Then
        MsgBox "Please enter the desired target level of error as a percentage (e.g., 10).", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    If Not SavedMarkerDetails And Not ShutdownRequested Then
        MsgBox "Please enter marker characteristics.", vbExclamation, "Input Required"
        MarkerCharacteristics.Show
        Exit Sub
    End If
    
    ' Make sure calculations are run first
    If Not ShutdownRequested Then
        CommandButton_Calculate_Linear_Click
    End If

    ' Initialize the variable to False
    SavedVariablesLinearExists = False

    ' Create a new worksheet named "Exported data (Linear)"
    ' Check if the sheet "Exported data (Linear)" already exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Exported data (Linear)" Then
            SavedVariablesLinearExists = True
            Set SavedVariablesLinear = ws
            Exit For
        End If
    Next ws

    ' If the sheet doesn't exist, create a new one
    If Not SavedVariablesLinearExists Then
        Set SavedVariablesLinear = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Calculator"))
        SavedVariablesLinear.Name = "Exported data (Linear)"
        AddHeadersLinear SavedVariablesLinear
        
        ' Clear nextRow and lastNonEmptyRow
        nextRow = 1
        lastNonEmptyRow = 0
    End If
    
    ' Determine the next empty row by examining all used columns
    lastNonEmptyRow = 0
    For i = 1 To 24 ' Data is up to column X
        nextRow = SavedVariablesLinear.Cells(Rows.Count, i).End(xlUp).Row
        If nextRow > lastNonEmptyRow Then
            lastNonEmptyRow = nextRow
        End If
    Next i
    nextRow = lastNonEmptyRow + 1
       
    ' Write values from the userform to specific cells in the next available row
    SavedVariablesLinear.Cells(nextRow, "A").Value = Now
    SavedVariablesLinear.Cells(nextRow, "B").Value = lastNonEmptyRow
    SavedVariablesLinear.Cells(nextRow, "C").Value = txt_SampleName.Text
    SavedVariablesLinear.Cells(nextRow, "D").Value = txt_X.Value
    SavedVariablesLinear.Cells(nextRow, "E").Value = txt_N.Value
    SavedVariablesLinear.Cells(nextRow, "F").Value = txt_ConfidenceInterval.Value
    SavedVariablesLinear.Cells(nextRow, "G").Value = txt_LevelError.Value
    SavedVariablesLinear.Cells(nextRow, "H").Value = MarkerCharacteristics.txt_N1.Value
    SavedVariablesLinear.Cells(nextRow, "I").Value = MarkerCharacteristics.txt_Y1.Value
    SavedVariablesLinear.Cells(nextRow, "J").Value = MarkerCharacteristics.txt_s1.Value
    SavedVariablesLinear.Cells(nextRow, "K").Value = MarkerCharacteristics.txt_N2.Value
    SavedVariablesLinear.Cells(nextRow, "L").Value = MarkerCharacteristics.txt_Y2.Value
    SavedVariablesLinear.Cells(nextRow, "M").Value = MarkerCharacteristics.ComboBox_Units.Value
    SavedVariablesLinear.Cells(nextRow, "N").Value = MarkerCharacteristics.txt_s2.Value
    SavedVariablesLinear.Cells(nextRow, "O").Value = LabelResult_Concentration_Linear.Value
    SavedVariablesLinear.Cells(nextRow, "P").Value = LabelResult_ConcentrationStandardError_Linear.Value
    SavedVariablesLinear.Cells(nextRow, "Q").Value = LabelResult_ConcentrationMax_Linear.Value
    SavedVariablesLinear.Cells(nextRow, "R").Value = LabelResult_ConcentrationMin_Linear.Value
    SavedVariablesLinear.Cells(nextRow, "S").Value = LabelResult_uhat_Linear.Value
    SavedVariablesLinear.Cells(nextRow, "T").Value = LabelResult_CollectionEffort_Linear.Value
    SavedVariablesLinear.Cells(nextRow, "U").Value = LabelResult_PredictedCollectionEffort_Linear.Value
    SavedVariablesLinear.Cells(nextRow, "V").Value = N3C
    SavedVariablesLinear.Cells(nextRow, "W").Value = TimeFOV
    SavedVariablesLinear.Cells(nextRow, "X").Value = TimeTotal
    
    ' Inform the user that values have been saved
    InfoExported = True
    MsgBox "Values have been saved to the worksheet 'Exported data (Linear)'.", vbInformation
    
    If ShutdownRequested Then
        End
    End If
End Sub

Private Sub AddHeadersLinear(ByRef ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "Date and time (DD/MM/YYYY XX:XX)"
        .Cells(1, 2).Value = "Data export #"
        .Cells(1, 3).Value = "Sample name"
        .Cells(1, 4).Value = "Number of target specimens [x]"
        .Cells(1, 5).Value = "Number of marker specimens [n]"
        .Cells(1, 6).Value = "Confidence interval [CI] size (in %)"
        .Cells(1, 7).Value = "Desired level of total error [sigma-bar]"
        .Cells(1, 8).Value = "Number of doses of exotic marker specimens [N1]"
        .Cells(1, 9).Value = "Mean number of exotic markers per dose [Ybar1]"
        .Cells(1, 10).Value = "Sample standard deviation of exotic markers per dose [s1]"
        .Cells(1, 11).Value = "Total number of samples [N2]"
        .Cells(1, 12).Value = "Sample size (or mean sample size, if N2 > 1) [Ybar2]"
        .Cells(1, 13).Value = "Size unit"
        .Cells(1, 14).Value = "Standard deviation of sample size [s2]"
        .Cells(1, 15).Value = "Concentration of target specimens [cL]"
        .Cells(1, 16).Value = "Total standard error of concentration [sigma-L]"
        .Cells(1, 17).Value = "Confidence interval maximum for concentration estimate [CImax]"
        .Cells(1, 18).Value = "Confidence interval minimum for concentration estimate [CImin]"
        .Cells(1, 19).Value = "Target-to-market ratio [u-hat]"
        .Cells(1, 20).Value = "Present data collection effort (time units) [eL]"
        .Cells(1, 21).Value = "Predicted data collection effort to achieve desired error rate (time units) [eL-sigma-bar]"
        .Cells(1, 22).Value = "Number of fields of view counted [N3C]"
        .Cells(1, 23).Value = "Transition time (in seconds)"
        .Cells(1, 24).Value = "Total count time (in seconds)"
        
        ' Force three decimal places
        .Cells(1, 19).NumberFormat = "0.000"
    End With
End Sub

Private Sub CommandButton_MarkerCharacteristics_Click()
    LinearChosen = True
    MarkerCharacteristics.Show
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

Private Sub txt_ConfidenceInterval_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9), Backspace, and Dot (.) key (if not already entered)
    Select Case KeyAscii
        Case 8 ' Backspace
        Case 46 ' Dot
            If Len(txt_ConfidenceInterval.Text) = 0 Then
                ' Disallow dot if textbox is empty
                KeyAscii = 0
            ElseIf InStr(txt_ConfidenceInterval.Text, ".") > 0 Then
                ' Disallow dot if dot already exists
                KeyAscii = 0
            End If
        Case 48 To 57, 96 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            ' Find position of decimal point if present
            DecimalPosition = InStr(txt_ConfidenceInterval.Text, ".")
            If DecimalPosition > 0 Then
                ' Calculate number of digits after decimal point
                NumDigitsAfterDecimal = Len(txt_ConfidenceInterval.Text) - DecimalPosition
                If NumDigitsAfterDecimal >= 2 Then
                    ' Block more than two digits after the decimal
                    KeyAscii = 0
                End If
            End If
            ' Additional check for '0' as first character
            If (KeyAscii = 48 Or KeyAscii = 96) And Len(txt_ConfidenceInterval.Text) = 0 And DecimalPosition = 0 Then
                ' Disallow '0' if it's the first character and no decimal point
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

Private Sub txt_ConfidenceInterval_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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
            If InputsSaved And SavedVariablesLinearExists And InfoExported Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("Would you like to export your data one last time before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_Linear_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            ElseIf InputsSaved And SavedVariablesLinearExists And Not InfoExported Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("There are data in the 'Exported data (Linear) spreadsheet from previous trials. Would you like to export your data here before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_Linear_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            ElseIf Not (InputsSaved Or SavedMarkerDetails) And SavedVariablesLinearExists Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("There are variables that differ from those in the latest export. Would you like to export these before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_Linear_Click ' Run subroutine to export data
                Else
                    ' User chose not to save, proceed to close
                    Cancel = 0
                    End
                End If
            ElseIf Not SavedVariablesLinearExists And OutputsSaved Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("Would you like to export the saved information to a spreadsheet before closing the application?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    CommandButton_SaveVariables_Linear_Click 'Run subroutine to export data.
                Else
                    Cancel = 0
                    End ' Terminates application and erases all data from memory.
                End If
            ElseIf (txt_X.Value <> "" Or txt_N.Value <> "" Or txt_ConfidenceInterval.Value <> "" Or txt_LevelError.Value <> "") Or SavedMarkerDetails Then
                Cancel = 1 ' Cancel the close operation
                response = MsgBox("Stored variables will be deleted if the application is closed. Would you like to export these first?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Data")
                If response = vbYes Then
                    ShutdownRequested = True
                    'CommandButton_SaveVariables_Linear.Enabled = True
                    CommandButton_SaveVariables_Linear_Click ' Run subroutine to export data
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
