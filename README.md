# FOVS_vs_linear_methods
This ReadMe file was generated on 2025-02-25 by Marcos Amores.

## General information
   
  The following has been developed for the manuscript "Field-of-view subsampling: A novel ‘exotic marker’ method for absolute abundances, validated by simulation and microfossil case studies".

  The manuscript has been accepted for publication in the journal PLoS One; you can access it here: https://doi.org/10.1371/journal.pone.0320887.
  
### Corresponding author and affiliation:

  Name: Chris Mays

  Institution: Department of Geology & Palaeontology, Natural History Museum Vienna
  
  Address: Burgring 7, 1010 Vienna, Austria
  
  Email: chris.mays@nhm.at




## Absolute abundance calculator

### Latest version:

  v1.1.2


### Date of present version:
   
  2025-02-25


### Licenses/restrictions placed on the code:
   
  Absolute abundance calculator © 2025 by Chris Mays, Marcos Amores, and Anthony Mays. Licensed under Creative Commons Attribution-ShareAlike 4.0 International


### Troubleshooting

If you see the warning message: "Microsoft has blocked macros from running because the source of this file is untrusted."

  You can unblock macros by modifying the properties of the file as follows:

    1) Open Windows File Explorer and go to the folder where you saved the file.

    2) Right-click the file and choose Properties from the context menu.

    3) At the bottom of the General tab, select the Unblock checkbox and select OK.
    
  For additional details, please see the official guide on how to resolve this at the Microsoft Support website, link below:
  
  https://support.microsoft.com/en-us/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216





## Matlab code for simulations 

### Quick Start Guide

  Get Matlab ready:
  * Download the Matlab files to your working directory.
  * Make a \SimData\ subdirectory in your working directory.
  * Note that Matlab must have the Statistics and Machine Learning Toolbox installed (which is required to perform the t-tests).
     + If you run the code without this Toolbox installed, then you will receive an error. There should be a link in the error message that will install the Toolbox for you (depending on your Matlab distribution and licence).

  To generate the data for:
  * For Tables S4--S14 and Fig S1, use the file BigFossilSimsV3.m by typing
  ```
	>> BigFossilSimsV3
  ```
    on the Matlab command line and hit Enter.
  * For Fig 3, use the file SimStatsChecker.m by typing
  ```
	>> SimStatsChecker(30000,1000,700,15,100000,1)
  ```
    on the Matlab command line and hit Enter.
  * For Fig 6, use the file PrecWRTWorkV3.m by typing
  ```
	>> PrecWRTWorkV3
  ```
    on the Matlab command line and hit Enter.

  * Note that some of these files will take many hours (even days) to run. If you wish to run shorter versions, change the "its" parameter (i.e., iterations) to a smaller number. You can generate new simulations with different parameters using the more detailed instructions contained in "Supporting information file 2".
