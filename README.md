# FOVS_vs_linear_methods
This ReadMe file was generated on 2024-11-20 by Marcos Amores

## General information

### 1. Title:
   
  Absolute abundance calculator (v1.1.0)

### 2. Leading/corresponding author:
   
  Name: Chris Mays

  Institution: Department of Geology & Palaeontology, Natural History Museum Vienna
  
  Address: Burgring 7, 1010 Vienna, Austria
  
  Email: chris.mays@nhm.at

### 3. Date of present version:
   
  2024-12-14


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





# Matlab code for simulations --- Quick Start Guide

Get Matlab ready:
* Download the Matlab files to your working directory.
* Make a \SimData\ subdirectory in your working directory.
* Note that Matlab must have the Statistics and Machine Learning Toolbox installed (which is required to perform the t-tests).
   + If you run the code without this Toolbox installed, then you will receive an error. There should be a link in the error message that will install the Toolbox for you (depending on your Matlab distribution and licence).

To generate the data for:
* Tables S4--S14 and Fig S1 use the file BigFossilSimsV3.m, by typing
```
	>> BigFossilSimsV3
```
  on the Matlab command line and hit Enter.
* Fig 3 use the file SimStatsChecker.m by typing
```
	>> SimStatsChecker(30000,1000,700,15,100000,1)
```
   on the Matlab command line and hit Enter.
* Fig 8 use the file PrecWRTWorkV3.m by typing
```
	>> PrecWRTWorkV3
```
   on the Matlab command line and hit Enter.

Note that some of these files will take many hours (even days) to run. If you wish to run shorter versions, change the "its" parameter (i.e., iterations) to a smaller number. You can generate new simulations with different parameters using the more detailed instructions contained in "Supporting information file 2".
