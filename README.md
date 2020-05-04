# covid19

## Analysis considerations

### Part 1 - Process results of RT-PCR assay
- Compare the Ct values to expected threshold for each of three reactions (N1, N2, RNaseP).
  - N1 and N2 reactions need to meet a specific Ct threshold (per CDC reference document < 40) to qualify as positive result. 
  - A log is made detailing experiment information and the Ct values from NTC (negative), HSC (extraction), and nCoVPC (positive) controls.
  
### Part 2 - Convert Meditech to BSI out file for upload into BSI (COVID-friendly)
- Convert output files from Meditch to BSI R script for upload into BSI.
- Duplicate selected columns.
- Change study field to COVID.

### Part 3 - Process results of ELISA assay
- Calculate average absorbance for all samples and controls. 
  - The average value of the absorbance of the negative control must be less than 0.25.
  - The absorbance of the positive control must not be less than 0.30. 
  - For interpretation of results the positive and negative cutoffs are defined based on the average value of the absorbance of the negative control.
    - Positive cutoff = 1.1 * (xNC + 0.18)
    - Negative cutoff = 0.9 * (xNC + 0.18)
  - A log file is made detailing experiment information. 
