Attribute VB_Name = "Changelog"
'#####
'# 1.1.1
'#
'# 2025-01-XX
'#####

'# Features
'- Pasting values in input boxes is re-enabled.
'- Exporting inputs and variables no longer requires all input fields to be filled out.
'- FOVS calculations no longer require all inputs to be inserted before running calculations. Outputs are now given whenever the conditions required for their equation are fulfilled.

'# GUI
'- Hovering input boxes now also gives description of what it is for.
'- Added toltip description text to certain text elements that were lacking them.
'- In the FOVS (Target and Marker) forms the Y3x and Y3n values now show automatically after saving the variables in the "Field of view (FOVS) calibration count" (CalibratorFOV).
'- Empty (= 0) values for N*3C, N*3E, delta* no longer show in the FOVS (Target and Marker) forms after saving variables in the "Field of view (FOVS) calibration count" (CalibratorFOV).
'- N3E and both N and X values can now be inserted from the get-go once FOVS (Target and Marker) forms are initialised.
'- Tabbing now goes vertically instead of horizontally in the Counting Assistant form.
'- Counting Assistant now imports the Sample Name if there is one stored in memory.
'- X and N values now transfer to the "Optimisation data" if values are present in the text boxes for both FOVS forms.

'# Calculations
'- Order of calculations now considers each variable independently in order to allow specific ones to be readily calculated without waiting for inputs for other variables.

'# Code
'- Exporting inputs and variables no longer runs re-runs calculations beforehand.
'- Exporting multiple times no longer crashes the application. Before, the code was looking for example for the sheet "FOVS-T" when in fact it is currently defined as "FOVS_Target".
'- Clearing all data no longer spams the user with "unsaved input" warnings and properly removes the variables in "Marker Characteristics" and "Optimisation data".

'# Text
'- When exporting the confirmation message now correctly refers to the sheet, for example from "Exported data (FOVS-M)" to "Saved Variables (FOVS_Marker)".
'- Removed the "Optional" in "Optional: Optimisation data" present in the FOVS forms.

'#####
'# 1.1.0 (Dubgall)
'#
'# 2024-11-20
'#####

'# Features
'- Added a Glossary userbox that explains the symbology.
'- Resetting the timer now requires user confirmation.

'# GUI
'- Made interfaces more compact.
'- Increased size of arrows by two units in the Counting Assistant userform.
'- FOVS method are now explicit in the title (Target/Marker focus).
'- Descriptions are now captions that appear as a tool-tip hover in the FOVS method userforms.
'- Timer now shows tenths of a second.
'- Reset timer button is now always visible.

'# Text
'- Fixed title of pop-up box for when users want to change the focus to targets/markers in the FOVS method userform.
'- Included a text on the FOVS method to let users know they can over symbols for a description of what they are.
'- The term "full counts" is renamed to "extrapolation counts".
'- N3F (full counts) changed to N3E (extrapolation counts).

'# Calculations
'- Fixed issue with the Gamma function for C4 calculation when the number of calibration count fields-of-view was equal to 1.
'- Avoided division by 0 when calculating u-hat values when marker information is not presently stored in memory in the FOV calibration userform.

'# Code
'- Streamlined code responsible for number input in Method Determination userform.

'#####
'# 1.0.0 (Release)
'#
'# 2024-06-03
'#####
