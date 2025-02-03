Attribute VB_Name = "ShowCalculator"
'
' Public Variables
'
    ' Stores values for inputs, outputs, and background calculations that have use in various user forms.
    ' Values expected to have decimal places are stored as Double. Values without expected decimal values (e.g., counts) are Long.
    
    Option Explicit ' Declare all the variables, even if not used during a calculation run.
    
    ' Counts
        Public X As Long ' Number of targets.
        Public N As Long ' Number of markers.
    
    ' Time
        Public NumFOV As Long ' Number of counted fields-of-view. Here defined independently, will later help define N3C (calibration count fields-of-views) or N3E (extrapolated counts fields-of-view).
        Public TimeTotal As Long ' Seconds since calibration count start meant to help determine effort calculations. Not used for extrapolation counts.
        Public TimeFOV As Long ' Seconds it takes to transition between fields-of-view. Used to later be substracted with TimeTotal in order to see the actual time it takes to count specimens.
    
    ' Timer related.
    
        Public StartTime As Double 'Stores OS system time by calling the 'Timer' function.
        Public ElapsedSeconds As Double 'Elapsed time in seconds.
        Public PausedTime As Double 'Stores how many seconds a timer was a paused.
        Public TimerRunning As Boolean 'Flag to check if timer subroutine is running or not.
    
    ' Effort and Calibrations
        Public TimeFOVTotal As Long ' Total time it takes to transition between fields-of-view.
        Public TimeTotalNoFOV As Long ' Total time it takes to count with field-of-view transitioning time removed.
        Public TimePerSpecimen As Double 'Time it takes to count each specimen.
        Public Ystar3n As Double ' Critical value of field-of-view target density (Y3) whereby either FOVS or linear methods are better for when targets [x] are more common than markers [n].
        Public Ystar3x As Double ' Critical value of field-of-view target density (Y3) whereby either FOVS or linear methods are better for when markers [n] are more common than targets [x].
        Public MethodDetFactor As Double
        Public LevelError As Double 'Lowercase sigma-line; User-defined target level of error in percentage.
        Public DecimalPosition As Integer 'Defines at which position the dot is in a integer.
        Public NumDigitsAfterDecimal As Integer 'Defines how many digits after the dot certain action must happen.
        Public FOVTransitionEffort As Double ' Lowercase omega; field-o-fview transition factor. Used to determine best method for counting.
        Public Y3x As Double ' Mean value of targets per field-of-view.
        Public Y3n As Double ' Mean value of markers per field-of-view.
        Public N3C As Double ' Calibration count fields-of-view seen.
        Public N3E As Double ' Extrapolated counts fields-of-view seen.
        Public uhat As Double ' Target-to-marker ratio.
        Public xhat As Double ' Extrapolated number of counted target specimens for the extrapolation counts.
        Public nhat As Double ' Extrapolated number of counted target specimens for the extrapolation counts.
        Public eL As Double ' Data collection effort for linear method.
        Public eL_sigmabar As Double ' Pedicted data collection effort for linear method.
        Public eF As Double ' Data collection effort for FOVS method.
        Public eF_sigmabar As Double ' Pedicted data collection effort for FOVS method.
        Public Nstar3C As Double ' Optimal number of calibration count fields-of-view.
        Public Nstar3E As Double ' Optimal number of extrapolation count fields-of-view.
        Public deltastar As Double ' Optimal field-of-view count ratio.
    
    ' Marker characteristics
        Public N1 As Long ' Number of doses of marker specimens.
        Public Y1 As Double ' Mean number of markers for one dose.
        Public s1 As Double ' Sample standard deviation for one dose of markers.
        Public N2 As Long ' Total number of samples.
        Public Y2 As Double ' Mean sample size.
        Public SizeUnit As String ' Type of unit selected for sample size.
        Public s2 As Double ' Sample standard deviation for mass or volume.
        Public s3 As Double ' Sample standard deviation for targets (if more common) or markers (if more common).
    
    ' Flags
        ' True or False variables meant to check if certain processes were run and to allow some conditional calculations.
        Public IntroductionGiven As Boolean
        Public LinearSuggested As Boolean
        Public TargetSuggested As Boolean ' Flag to automatically assign FOVS button to run caculations that assume targets [x] are more common.
        Public MarkerSuggested As Boolean ' Flag to automatically assign FOVS button to run caculations that assume markers [n] are more common.
        Public SavedMarkerDetails As Boolean
        Public CountingEffortCalibration As Boolean
        Public CalibratedFOV As Boolean
        Public LinearChosen As Boolean
        Public FOVSTargetChosen As Boolean
        Public FOVSMarkerChosen As Boolean
        Public FOVSTargetIntroGiven As Boolean
        Public FOVSMarkerIntroGiven As Boolean
        Public ClearedAllData As Boolean
        Public ShutdownRequested As Boolean
        Public CountingSaved As Boolean
        Public OriginStarter As Boolean
        Public OriginLinear As Boolean
        Public OriginFOVSTarget As Boolean
        Public OriginFOVSMarker As Boolean
        Public OriginCountingEffort As Boolean
        Public OriginCalibrationFOV As Boolean
        Public OriginPreliminaryData As Boolean
        Public MethodSwitchIgnored As Boolean
        
    ' Worksheet
        Public ws As Worksheet
        Public SavedVariablesLinear As Worksheet
        Public SavedVariablesFOVSTarget As Worksheet
        Public SavedVariablesFOVSMarker As Worksheet
        Public nextRow As Long
        Public lastNonEmptyRow As Long
        Public SavedVariablesLinearExists As Boolean
        Public SavedVariablesFOVSTargetExists As Boolean
        Public SavedVariablesFOVSMarkerExists As Boolean
    
    ' Infoboxes
        Public response As VbMsgBoxResult 'Meant to define information boxes that can be rewritten or have options included.
        
    ' Strings
        Public SampleName As String

'
' Calculators
'
    ' Defines each user box and allows them to be usable in calculations.

    Sub CalculatorStart() ' Method Determination; first to appear for users.
        Dim CalculatorStart As New CalculatorStart
        
        CalculatorStart.Show ' Show the user form.
    End Sub
    
    Sub CalculatorLinear() ' Linear method.
        Dim CalculatorLinear As New CalculatorLinear
        
        CalculatorLinear.Show ' Show the user form.
    End Sub
    
    Sub CalculatorFOVS() ' FOVS method.
        Dim CalculatorFOVSTarget As New CalculatorFOVSTarget ' For when targets are more common.
        Dim CalculatorFOVSMarker As New CalculatorFOVSMarker ' For when markers are more common.
        
        'Message to check which specimen type is more common to later apply the correct calculations.
        response = MsgBox("Are targets more common than markers?", vbQuestion + vbYesNo, "Most common specimens?")
        
        ' Check user response
        If response = vbYes Then
            FOVSTargetChosen = True ' Sets flag.
            CalculatorFOVSTarget.Show ' Show the user form.
        Else
            FOVSMarkerChosen = True ' Sets flag.
            CalculatorFOVSMarker.Show ' Show the user form.
        End If
    End Sub
    
'
' Associated
'

    Sub PreliminaryData() ' Preliminary Data
        Dim PreliminaryData As New PreliminaryData
    End Sub
    
    Sub AssistantCounting() ' Counting Assistant.
        Dim AssistantCounting As New AssistantCounting
    End Sub
    
    Sub CountingEffort() ' Optimisation Data.
        Dim CountingEffort As New CountingEffort
    End Sub
    
    Sub MarkerCharacteristics() ' Marker and Sample Characteristics.
        Dim MarkerCharacteristics As New MarkerCharacteristics
    End Sub
    
    Sub CalibratorFOV() ' Field of view Calibrator.
        Dim CalibratorFOV As New CalibratorFOV
    End Sub
