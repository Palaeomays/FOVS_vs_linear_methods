VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MarkerCharacteristics 
   Caption         =   "Marker and sample characteristics"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5670
   OleObjectBlob   =   "MarkerCharacteristics.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MarkerCharacteristics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Variables
'

    Private InputsSaved As Boolean
    Private InputEmptyAny As Boolean
    Dim InputEmptyN1 As Boolean
    Dim InputEmptyY1 As Boolean
    Dim InputEmptyS1 As Boolean
    Dim InputEmptyN2 As Boolean
    Dim InputEmptyY2 As Boolean
    Dim InputEmptyUnits As Boolean
    Dim InputEmptyS2 As Boolean
    Private UnsavedWarningGiven As Boolean
    
'
' Startup
'
    Private Sub UserForm_Initialize()
    
        UnsavedWarningGiven = False
    
        ' Populate inputs with previous values if these exist. Sets the value of the textbox to the value of the public variable.
        
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
            txt_s1.Enabled = True
            txt_s1.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txt_s1.Text = ""
            txt_s1.Enabled = False
            txt_s1.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
    
        ' N2
        
        If N2 <> 0 Then
            txt_N2.Text = N2
        Else
            txt_N2.Text = ""
        End If
        
        ' Y2
        
        If Y2 <> 0 Then
            txt_Y2.Text = Y2
        Else
            txt_Y2.Text = ""
        End If
        
        ' Units
        
        If SizeUnit <> "" Then
            ComboBox_Units.Value = SizeUnit
        End If
        
        ' Populate Units drop-down with examples.
        
        With ComboBox_Units
            .AddItem "kg"
            .AddItem "g"
            .AddItem "mg"
            .AddItem "m³"
            .AddItem "cm³"
            .AddItem "mm³"
            .AddItem "m²"
            .AddItem "cm²"
            .AddItem "cm²"
            .AddItem "Other"
        End With
        
        ' S2
        
        If s2 <> 0 Then
            txt_s2.Text = s2
            txt_s2.Enabled = True
            txt_s2.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txt_s2.Text = ""
            txt_s2.Enabled = False
            txt_s2.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
    End Sub

'
' Units
'

    Private Sub ComboBox_Units_Change()
        If ComboBox_Units.Value = "Other" Then
            ComboBox_Units.Visible = False
            txt_UnitsOther.Visible = True ' Show the text box for user input.
            txt_UnitsOther.SetFocus ' Sets focus to the text box
        Else
        End If
    End Sub
    
    Private Sub txt_UnitsOther_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        If Len(txt_UnitsOther.Text) = 0 Then
            txt_UnitsOther.Visible = False
            ComboBox_Units.Visible = True
            ComboBox_Units.Value = "" ' Clear the selection
        End If
    End Sub
   
'
' Saving
'
    ' Activate standard deviation text boxes if number of samples > 1.
    
    ' S1
      
    Private Sub txt_N1_Change()
        
        If IsNumeric(txt_N1.Value) And txt_N1.Value > 1 Then
            txt_s1.Value = ""
            txt_s1.Enabled = True
            txt_s1.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txt_s1.Value = ""
            txt_s1.Enabled = False
            txt_s1.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        If InputsSaved And txt_N1.Value <> N1 Then
            UnsavedWarningGiven = False
        End If
               
    End Sub
    
    Private Sub txt_Y1_Change()
        
        If IsNumeric(txt_N1.Value) And txt_N1.Value > 1 Then
            txt_s1.Value = ""
            txt_s1.Enabled = True
            txt_s1.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txt_s1.Value = ""
            txt_s1.Enabled = False
            txt_s1.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
    
        If InputsSaved And txt_Y1.Value <> Y1 Then
            UnsavedWarningGiven = False
        End If
               
    End Sub
    
    Private Sub txt_S1_Change()
        If InputsSaved And txt_s1.Value <> s1 Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' S2
    
    Private Sub txt_N2_Change()
        
        If IsNumeric(txt_N2.Value) And txt_N2.Value > 1 Then
            txt_s2.Value = ""
            txt_s2.Enabled = True
            txt_s2.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txt_s2.Value = ""
            txt_s2.Enabled = False
            txt_s2.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        If InputsSaved And txt_N2.Value <> N2 Then
            UnsavedWarningGiven = False
        End If
        
    End Sub
    
    Private Sub txt_Y2_Change()
        
        If IsNumeric(txt_N2.Value) And txt_N2.Value > 1 Then
            txt_s2.Value = ""
            txt_s2.Enabled = True
            txt_s2.BackColor = RGB(255, 255, 255) ' White colour.
        Else
            txt_s2.Value = ""
            txt_s2.Enabled = False
            txt_s2.BackColor = RGB(224, 224, 224) ' Grey colour.
        End If
        
        If InputsSaved And txt_Y2.Value <> Y2 Then
            UnsavedWarningGiven = False
        End If
               
    End Sub
    
    Private Sub ComboBox_Units_AfterUpdate()
        If InputsSaved And ComboBox_Units.Value <> SizeUnit Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    Private Sub txt_s2_Change()
        
        If InputsSaved And txt_s2.Value <> s2 Then
            UnsavedWarningGiven = False
        End If
    End Sub
    
    ' Store values in memory after clicking save.
    
    Private Sub CommandButtonSaveMarkers_Click()
        
        ' N1
        
        If IsNumeric(txt_N1.Value) Then
            N1 = CLng(txt_N1.Value)
            InputEmptyN1 = False
        Else
            InputEmptyN1 = True
        End If
        
        ' Y1
        
        If IsNumeric(txt_Y1.Value) Then
            Y1 = CDbl(txt_Y1.Value)
            InputEmptyY1 = False
        Else
            InputEmptyY1 = True
        End If
        
        ' S1
        
        If IsNumeric(txt_s1.Value) And txt_N1.Value > 1 Then
            s1 = CDbl(txt_s1.Value)
            InputEmptyS1 = False
        ElseIf txt_N1.Value = 1 Then
            s1 = Sqr(Y1)
            txt_s1.Value = s1
            txt_s1.Enabled = True
            txt_s1.BackColor = RGB(255, 255, 255) ' White colour.
            InputEmptyS1 = False
        Else
            InputEmptyS1 = True
        End If
        
        ' N2
        
        If IsNumeric(txt_N2.Value) Then
            N2 = CLng(txt_N2.Value)
            InputEmptyN2 = False
        Else
            InputEmptyN2 = True
        End If
        
        ' Y2
        
        If IsNumeric(txt_Y2.Value) Then
            Y2 = CDbl(txt_Y2.Value)
            InputEmptyY2 = False
        Else
            InputEmptyY2 = True
        End If
        
        ' Units
            
        If IsNumeric(txt_Y2.Value) And ComboBox_Units.Value <> "" Then
            SizeUnit = ComboBox_Units.Value
            InputEmptyUnits = False
        ElseIf IsNumeric(txt_Y2.Value) And ComboBox_Units.Value = "" Then
            MsgBox "Please select the unit of size [Y2].", vbExclamation, "Input Required"
            ComboBox_Units.SetFocus
            Exit Sub
        Else
            InputEmptyUnits = True
        End If
        
        ' S2
        
        If IsNumeric(txt_s2.Value) And txt_N2.Value > 1 Then
            s2 = CDbl(txt_s2.Value)
            InputEmptyS2 = False
        ElseIf txt_N2.Value = 1 Then
            s2 = Sqr(Y2)
            txt_s2.Value = s2
            txt_s2.Enabled = True
            txt_s2.BackColor = RGB(255, 255, 255) ' White colour.
            InputEmptyS2 = False
        Else
            InputEmptyS2 = True
        End If
        
        If InputEmptyN1 Or InputEmptyY1 Or InputEmptyS1 Or InputEmptyN2 Or InputEmptyY2 Or InputEmptyUnits Or InputEmptyS2 Then
            InputEmptyAny = True
        Else
            InputEmptyAny = False
        End If
        
        MsgBox "Variables successfully saved.", vbInformation
    
        If LinearChosen And Not InputEmptyAny Then ' For Linear method
            CalculatorLinear.CommandButton_MarkerCharacteristics.BackColor = RGB(212, 236, 214) ' Greenish color
            CalculatorLinear.CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data ready)"
        ElseIf LinearChosen And InputEmptyAny Then
            CalculatorLinear.CommandButton_MarkerCharacteristics.BackColor = RGB(245, 148, 146) ' Reddish color
            CalculatorLinear.CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data missing)"
            
        ElseIf FOVSTargetChosen And Not InputEmptyAny Then ' For FOVS Target method
            CalculatorFOVSTarget.CommandButton_MarkerCharacteristics.BackColor = RGB(212, 236, 214) ' Greenish color
            CalculatorFOVSTarget.CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data ready)"
        ElseIf FOVSTargetChosen And InputEmptyAny Then
            CalculatorFOVSTarget.CommandButton_MarkerCharacteristics.BackColor = RGB(245, 148, 146) ' Reddish color
            CalculatorFOVSTarget.CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data missing)"
            
        ElseIf FOVSMarkerChosen And Not InputEmptyAny Then ' For FOVS Target method
            CalculatorFOVSMarker.CommandButton_MarkerCharacteristics.BackColor = RGB(212, 236, 214) ' Greenish color
            CalculatorFOVSMarker.CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data ready)"
        ElseIf FOVSMarkerChosen And InputEmptyAny Then
            CalculatorFOVSMarker.CommandButton_MarkerCharacteristics.BackColor = RGB(245, 148, 146) ' Reddish color
            CalculatorFOVSMarker.CommandButton_MarkerCharacteristics.Caption = "Marker and sample characteristics" & vbCrLf & "(data missing)"
        End If
        
        ' Adding flags.
        
        InputsSaved = True
        
        If Not InputEmptyAny Then
            SavedMarkerDetails = True
        Else
            SavedMarkerDetails = False
        End If
        
        UnsavedWarningGiven = False
        
        Me.Hide
    End Sub

Private Sub CommandButtonClear_Click()
    ' Display a message box confirming the action and asking for confirmation
    response = MsgBox("Are you sure you want to clear the inputs?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Inputs")
    
    ' Check user's response
    If response = vbYes Then
        txt_N1.Text = ""
        txt_Y1.Text = ""
        txt_s1.Text = ""
        txt_N2.Text = ""
        txt_Y2.Text = ""
        ComboBox_Units.Value = ""
        txt_s2.Text = ""
    Else
    ' User cancelled, do nothing
    End If
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

Private Sub txt_N2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9) and Backspace key (if not already entered)
    Select Case KeyAscii
        Case 8
        Case 49 To 57, 97 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
        Case 48, 96 ' Numbers 0 and Numpad 0.
        If Len(txt_N2.Text) > 0 Then
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

Private Sub txt_Y2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9), Backspace, and Dot (.) key (if not already entered)
    Select Case KeyAscii
        Case 8 ' Backspace
            ' Do nothing, allow backspace
        Case 46 ' Dot
            If Len(txt_Y2.Text) = 0 Then
                ' Disallow dot if textbox is empty
                KeyAscii = 0
            ElseIf InStr(txt_Y2.Text, ".") > 0 Then
                ' Disallow dot if dot already exists
                KeyAscii = 0
            End If
        Case 48 To 57, 96 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            If Len(txt_Y2.Text) = 0 Then
                ' Allow input of 0 if textbox is empty
                ' Do nothing, allow input
            ElseIf txt_Y2.Text = "0" Then
                ' Disallow input of 0 if it's already present
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0 ' Disallow other characters
    End Select
End Sub

Private Sub txt_S2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow numbers (0-9), Backspace, and Dot (.) key (if not already entered)
    Select Case KeyAscii
        Case 8 ' Backspace
            ' Do nothing, allow backspace
        Case 46 ' Dot
            If Len(txt_s2.Text) = 0 Then
                ' Disallow dot if textbox is empty
                KeyAscii = 0
            ElseIf InStr(txt_s2.Text, ".") > 0 Then
                ' Disallow dot if dot already exists
                KeyAscii = 0
            End If
        Case 48 To 57, 96 To 105 ' Numbers 1-9 and Numpad numbers 0-9.
            If Len(txt_s2.Text) = 0 Then
                ' Allow input of 0 if textbox is empty
                ' Do nothing, allow input
            ElseIf txt_s2.Text = "0" Then
                ' Disallow input of 0 if it's already present
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0 ' Disallow other characters
    End Select
End Sub

' Avoid pasting words and numbers.

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

Private Sub txt_N2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_Y2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub txt_S2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And (KeyCode = 86) Then ' Disable Ctrl+V (paste)
        KeyCode = 0
    End If
End Sub

Private Sub CommandButton_Assistant_Click()
    MarkerCharacteristicsAssistant.Show
End Sub

'
' Shutdown
'

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' Ask user if they want to close the form without saving.
    ' For this warning to appear, check if inputs were saved, if input boxes contain values, and that these values are different from those stored in memory.
    
    ' To avoid several warnings in the case of many unsaved variables, the flag 'UnsavedWarningGiven' checks if such a warning has come up yet.
    
    ' N1
    If Not ClearedAllData Then
        If Not UnsavedWarningGiven Then
            UnsavedWarningGiven = True
            If IsNumeric(txt_N1.Value) And txt_N1.Value <> N1 Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveMarkers_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
            
        ' Y1
        
        If Not UnsavedWarningGiven Then
            UnsavedWarningGiven = True
            If IsNumeric(txt_Y1.Value) And txt_Y1.Value <> Y1 Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveMarkers_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
            
        ' S1
            
        If Not UnsavedWarningGiven Then
            If IsNumeric(txt_s1.Value) And txt_s1.Value <> s1 Then
            UnsavedWarningGiven = True
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveMarkers_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
            
        ' N2
        
        If Not UnsavedWarningGiven Then
            If IsNumeric(txt_N2.Value) And txt_N2.Value <> N2 Then
            UnsavedWarningGiven = True
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveMarkers_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
            
        ' Y2
            
        If Not UnsavedWarningGiven Then
        UnsavedWarningGiven = True
            If IsNumeric(txt_Y2.Value) And txt_Y2.Value <> Y2 Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveMarkers_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
        
        ' Units
            
        If Not UnsavedWarningGiven Then
        UnsavedWarningGiven = True
            If IsNumeric(ComboBox_Units.Value) And ComboBox_Units.Value <> SizeUnit Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveMarkers_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
            
        ' S2
            
        If Not UnsavedWarningGiven Then
        UnsavedWarningGiven = True
            If IsNumeric(txt_s2.Value) And txt_s2.Value <> s2 Then
                If CloseMode = 0 Then
                    Cancel = 1 ' Cancel the close operation.
                End If
                response = MsgBox("You have unsaved inputs. Would you like to export these input data to a spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Unsaved Inputs")
                If response = vbYes Then
                    CommandButtonSaveMarkers_Click 'Run subroutine to save inputs.
                Else
                    Cancel = 0
                    Unload Me
                End If
            End If
        End If
    End If
    
    ClearedAllData = False
    
End Sub
