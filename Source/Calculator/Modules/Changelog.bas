Attribute VB_Name = "Changelog"
'#####
'# 1.1.0 (Dubgall)
'#
'# 2024-11-20
'#####

'# Features
'- TODO Added import dataset button to populate calculator with previously stored information.
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
'- Fixe issue with the Gamma function for C4 calculation when the number of calibration count fields-of-view was equal to 1.
'- Avoided division by 0 when calculating u-hat values when marker information is not presently stored in memory in the FOV calibration userform.

'# Code
'- Streamlined code responsible for number input in Method Determination userform.

'#####
'# 1.0.0 (Release)
'#
'# 2024-06-03
'#####
