# FOVS_vs_linear_methods
This ReadMe file was generated on 2024-11-20 by Marcos Amores

## General information

### 1. Title:
   
  Absolute abundance calculator (v1.1.0)

### 2. Leading/corresponding author:
   
  Name: Chris Mays

  Institution: Department of Geology & Palaeontology, Natural History Museum Vienna
  
  Address: Burgring 7, 1010 Vienna, Austria
  
  Email: cmays@ucc.ie

### 3. Date of present version:
   
  2024-11-20


## Sharing/Access Information

### 1. Licenses/restrictions placed on the code:
   
  Absolute abundance calculator © 2024 by Chris Mays, Marcos Amores, and Anthony Mays. Licensed under Creative Commons Attribution-ShareAlike 4.0 International

## Troubleshooting

### If you see the warning message: "Microsoft has blocked macros from running because the source of this file is untrusted."

  You can unblock macros by modifying the properties of the file as follows:

    1) Open Windows File Explorer and go to the folder where you saved the file.

    2) Right-click the file and choose Properties from the context menu.

    3) At the bottom of the General tab, select the Unblock checkbox and select OK.
    
  For additional details, please see the official guide on how to resolve this at the Microsoft Support website, link below:
  
  https://support.microsoft.com/en-us/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216





# Matlab code for simulations

The codes used to generate the simulation data in this paper have not been optimised, and have some components that are either not used or not fully implemented. However, in the interests of full transparency, we include the exact versions of the code that we used for our results below.

## Data for S5–S15 Tables and S17 Fig

### Code:
[Main] BigFossilSimsV3.m

[Dependent] MicrofossilSimV3.m

[Dependent] MicrofossilSim_iV3.m

[Dependent] FOVoptimiserV1.m

### Use:

Specify the following variables in BigFossilSimsV3.m:

[Line 3] its: the number of independent Monte Carlo instances to generate for each set of parameters.

[Line 22] params: [Mx, Mn, tab]

Mx: the total number of targets on each virtual study area.

Mn: the total number of markers on each virtual study area.

tab: value of the dose error used in Eqns 2 and 5.

Note: Multiple rows of this variable can be specified to run multiple batches, via: [(first batch parameters); (second batch parameters); ...]

[Line 27] alpha: this is the field-of-view transition factor (ω). Default is ω=2.

[Line 33] work: the fixed value of work that the program tries to achieve for each method.

Linear method: Eqn 9 is used to choose the number of targets to count.

FOVS method: Eqns 14 and 15 are used to choose the optimal number of calibration- and extrapolation-count fields of view, via the code FOVoptimiserV1.m.

### Notes:

Ensure that there is a \SimData\ subdirectory for the program in which to store the data files.

The command line output will be saved in a file called BigFossilSimsV3_Opt_TX_itsY.txt, where

X is 10000 times the tablet error (to ensure an integer); and

Y is the value of its.

## Data for Fig 3

### Code:

[Main] SimStatsChecker.m

[Dependent] MicrofossilSim_iCheck.m

### Use: 

Call the function SimStatsChecker(Mx,Mn,tlim,fn,its,fopt), where the arguments are:

Mx: The total number of targets on each virtual slide.

Mn: The total number of markers on each virtual slide.

tlim: (Linear method) the number of targets to count in the window.

fn: (FOVS method) the number of extrapolation-count fields of view in which to count markers.

its: The number of independent Monte Carlo instances to generate.

fopt: Not used. Set to 1.


## Data for Fig 8

### Code:

[Main] PrecWRTWorkV2.m
[Dependent] WorkSimV2.m
[Dependent] WorkSimV2_i.m

### Use:

Specify the following variables in PrecWRTWorkV2.m

* [Line 7] its: The number of independent Monte Carlo instances to generate for each set of parameters.

* [Line 8] bigfx: The number of calibration counts for the "high calibration counts" sequence in Fig 8 (black plus).

* [Line 9] medfx: The number of calibration counts for the "medium calibration counts" sequence in Fig 8 (blue stars).

* [Line 10] smallfx: The number of calibration counts for the "low calibration counts" sequence in Fig 8 (red stars).

* [Lines 19–24] params: [Mx, Mx, tlim, fnmax, omega]

* Mx: The total number of targets on each virtual slide.

* Mn: The total number of markers on each virtual slide.

* tlim: (Linear method) the number of targets to count in the window.

* fnmax: Not used. Set to 1.

* omega: This is the field of view transition factor (ω).

Note: Multiple rows of this variable can be specified to run multiple batches, via: [(first batch parameters); (second batch parameters); ...]
 
The simulations currently assume that the marker dose (e.g., tablet of Lycopodium spores) error is zero, i.e.: (s_1P/√(N_1 ))^2=0. If you wish to increase this, then change the following variable:

[WorkSimV2_i.m, Line 33] tab: Value of the marker dose error used in Eqns 2 and 5.

### Notes:

Ensure that there is a \SimData\ subdirectory for the program to store the data files in.

The command line output will be saved in a file called WorkSimOpt_tab0_itsY.txt, where:

Y is the value of its tab0 records that the marker dose error is zero for the simulations. This is hard-coded and will not update if the value of tab is changed in WorkSimV2_i.m.
