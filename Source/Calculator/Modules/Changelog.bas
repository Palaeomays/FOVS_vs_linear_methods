Attribute VB_Name = "Changelog"
'#####
'# 1.1.2
'#
'# 2025-02-25
'#####

'# GUI
'- Preliminary data now fills up if user returns back to it. Same for the Method Determination.
'- Preliminary screen now updates accordingly if it was arrived via Linear or FOVS buttons and redirects accordingly.
'- X and N values are now filled up automatically in the "Optional: Optimisation data" if these are stored in memory.
'- Modified button that changes between marker/target mode in FOVS to open up the Method Determination screen instead.
'- Glossary font size is now the same as other nearby buttons.
'- N and N3E is now greyed out in FOVS screen, with user needing to do a FOV calibration first. Once the "FOV calibration count" goes green these become enabled.
'- Output boxes for xline and nline now correctly turn white and selectable if there is a value present.
'- X and N values in FOVS screens are no longer automatically filled if these are given in the "Optimisation data" screen.
'- X and N values given in "Optimisation data" screen now automatically fill Linear method screen.
'- Nstar3E and Nstar3C now round out to the nearest integer.

'# Calculations
'- Fixed wrong value issues with Nstar3C and Nstar3E. For the first a division was required instead of a multiplication, and for the second the wrong variable was present (N1 instead of s1).
'- Linear method now does not ask for preliminary data.
'- Linear method screen no longer requires all inputs to be filled out before running calculations.
'- Level of error is no longer expressly required to run calculations and show outputs that make no use of it.
'- s3 and unit of measurement are now rounded off to 3 decimal places if automatically determined.

'# Code
'- "Optimisation data" screen now considers if targets or markers are more common and prompts the user to change the method.
'- Counting Assistant now requires sample name to be inserted first. User is prompted of this when pressing "Next FOV" or "End and export to spreadsheet".

'# Text
'- Added back the "Optional" in "Optional: Optimisation data" present in the FOVS forms.
'- The pop-up screen in case counting time is 0 is now "The time it takes to count specimens and field of view transitions must be greater than 0."
'- Changed hover text of FOVS screen X and N to "Number of target/marker specimens from extrapolation counts".
'- Changed Counting Assistant introductory first paragraph to "The counting assistant is designed for counting up to 10 concurrent specimen categories (e.g., one marker and nine targets.)"
'- Added "(FOV)" after "Second: perform a count of the first field of view."
'- Added sentence at end of Hotkey note saying "FOV = Field of view".
'- Changed default text for "Target 1 (x1)" to "Target 1 (x #1).

'#####
'# 1.1.1
'#
'# 2025-01-03
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
