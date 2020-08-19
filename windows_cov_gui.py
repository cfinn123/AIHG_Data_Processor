"""
By: Jeffrey Beck and Casey Finnicum

Date of inception: March 16, 2020
1. Program for determining results of 2019-nCoV testing at the Avera Institute for Human Genetics (AIHG).
Ingests files from RT-qPCR assay and creates summarized results for upload.
Reference: CDC-006-00019, Revision: 02

Date of addition: April 5, 2020
2. Added a button to convert output files from the "Meditech to BSI" R scripts to make them COVID-compatible
for upload into BSI.

Date of addition: April 29, 2020
3. Additional logic added for interpretation of ELISA used for detecting COVID-19 IgG antibody in human serum.
The assay is intended for qualitative detection only.
Reference: EAGLE Biosciences EDI Novel Coronavirus COVID-19 IgG ELISA Kit.

Date of addition: May 13, 2020
4. Added a button for plotting cumulative positive, negative, inconclusive results for all testing done at AIHG.
The button works at the level of the "resulting_completed" directory within the RT_PCR/results/processed parent
directory.

Date of addition: August 17, 2020
5. Added button to replace the 'Meditech to BSI' R script that was originally created by Matthijs Van Der Zee.
There were two components to the script, one for extracting only the current label, the other for full conversion of
the Meditech report to a BSI-friendly upload file.
"""

from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
from pandas import ExcelWriter
import os
import ntpath
import time
import logging
from PIL import ImageTk, Image
import numpy as np
import re
import glob
import matplotlib.pyplot as plt

root = Tk()
root.configure(bg='white')

img = ImageTk.PhotoImage(Image.open("./misc/aihg.gif"))
panel = Label(root, image=img)
panel.pack(side="bottom", fill="both", expand="yes")


class AIHGdataprocessor:
    def __init__(self, master):
        master.minsize(width=200, height=100)
        self.master = master
        master.title("AIHG Data Processor")

        # Button for analyzing RT_PCR data
        self.rtpcr_button = Button(master, text='Select RT-PCR file to analyze', command=self.dataprocess, width=30)
        self.rtpcr_button.pack(pady=10)

        # Button for converting Meditch to BSI (PGx formatting)
        self.bsiconvert_button = Button(master, text='Select Meditech file to convert for BSI',
                                        command=self.bsiprocess, width=30)
        self.bsiconvert_button.pack(pady=10)

        # Button for converting Meditch to BSI (COVID formatting)
        self.covidbsiconvert_button = Button(master, text='COVID - Select file to convert for BSI',
                                        command=self.covidbsiprocess, width=30)
        self.covidbsiconvert_button.pack(pady=10)

        # Button for manual antibody testing
        self.eagle_elisa_button = Button(master, text="Select ELISA file (EAGLE setup)", command=self.antibodyprocess,
                                         width=30)
        self.eagle_elisa_button.pack(pady=10)

        # self.convert_button5 = Button(master, text='Generate stats for selected files', command=self.statsprocess,
        #                               width=40)
        # self.convert_button5.pack(pady=10)

        self.dirplotbutton = Button(master, text='Plot results', command=self.dirplot, width=30)
        self.dirplotbutton.pack(pady=10)

        # self.dirstatsbutton_week = Button(master, text='Weekly results - Select "resulting_completed" directory',
        #                               command=self.dirstatsresultsweek, width=50)
        # self.dirstatsbutton_week.pack(pady=10)
        #
        # self.dirstatsbutton_month = Button(master, text='Monthly results - Select "resulting_completed" directory',
        #                               command=self.dirstatsresultsmonth, width=50)
        # self.dirstatsbutton_month.pack(pady=10)

        # Button for LIMS-friendly output
        self.lims_convert_button = Button(master, text="LIMS - Select RT-PCR file to analyze", command=self.limsprocess,
                                          width=30)
        self.lims_convert_button.pack(pady=10)

        # Button for Meditech-friendly output - will need follow up prompt for selecting metadata file from dashboard
        self.meditech_button = Button(master, text="MEDITECH - Select RT-PCR file to analyze",
                                      command=self.meditechprocess, width=40)
        self.meditech_button.pack(pady=10)

        # Help button
        self.info_button = Button(master, text="Help", command=self.info, width=10)
        self.info_button.pack(pady=10)

    def info(self):
        messages = ["1. For analysis of RT-PCR data, press 'Select RT-PCR file to analyze' "
                    "and navigate to the file of interest in the file browser. The results and "
                    "associated log files will be generated.",
                    "2. For Meditech file conversion for upload into BSI, press 'Select Meditech file to convert for "
                    "BSI' and select the Meditech file. The BSI-friendly text file will be created in the same "
                    "directory as the specified input file.",
                    "3. For COVID-friendly file conversion for upload into BSI, press 'COVID - Select file to convert "
                    "for BSI' and select output file from Meditech to BSI conversion. The COVID-friendly Excel file "
                    "will be created in the same directory as the specified input file.",
                    "4. For analysis of ELISA, press 'Select ELISA file (EAGLE setup)'. "
                    "Proceed to navigate to the appropriate ELISA results file.",
                    "5. For statistics and plots, please follow the on screen instructions.",
                    "6. For LIMS friendly output navigate to the file of interest in the file browser. The results "
                    "will appear in the results/processed/output_for_LIMS directory. A log file will be made in the "
                    "logs directory.",
                    "7. For Meditech friendly output navigate to the file of interest in the file browser. The results "
                    "will appear in the results/processed/output_for_Meditech directory. A log file will be made in "
                    "the logs directory."]

        messagebox.showinfo("Help", "\n\n".join(messages))

    def dataprocess(self):
        # Ingest input file
        # ask the user for an input read in the file selected by the user
        path = filedialog.askopenfilename()

        # Original - does not work for ViiA7
        # read in 'Results' sheet of specified file
        # df = pd.read_excel(path, sheet_name='Results', skiprows=42, header=0)

        # To accommodate either QuantStudio or ViiA7
        df_orig = pd.read_excel(path, sheet_name="Results", header=None)
        for row in range(df_orig.shape[0]):
            for col in range(df_orig.shape[1]):
                if df_orig.iat[row, col] == "Well":
                    row_start = row
                    break
        # Subset raw file for only portion below "Well" and remainder of header
        df = df_orig[row_start:]

        # Header exists in row 1, make new header
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header

        # Adding a new line to handle the 'Cт' present in the header of the output file from the 7500 instrument
        df.columns = df.columns.str.replace('Cт', 'CT')

        # Convert 'undetermined' to 'NaN' for 'CT' column
        df['CT'] = df.loc[:, 'CT'].apply(pd.to_numeric, errors='coerce')

        # Assess controls
        # Expected performance of controls
        """
        ControlType   ExternalControlName Monitors        2019nCoV_N1 2019nCOV_N2 RnaseP  ExpectedCt
        Positive      nCoVPC              Rgt Failure     +           +           +       <40
        Negative      NTC                 Contamination   -           -           -       None
        Extraction    HSC                 Extraction      -           -           +       <40

        If any of the above controls do not exhibit the expected performance as described, the assay may have been set
        up and/or executed improperly, or reagent or equipment malfunction could have occurred. Invalidate the run and
        re-test.
        """

        # TODO: DEFINE CT VALUE HERE
        ct_value = 40.00

        # Create results columns for NTC - non-template control (negative control)
        df['NTC_N1'] = None  # initial value
        df.loc[(df['Sample Name'] == 'NTC') & (df['Target Name'] == 'N1') & (df['CT'].isnull()), 'NTC_N1'] = 'passed'
        df.loc[(df['Sample Name'] == 'NTC') & (df['Target Name'] == 'N1') & (df['CT'].notnull()), 'NTC_N1'] = 'failed'
        df['NTC_N2'] = None  # initial value
        df.loc[(df['Sample Name'] == 'NTC') & (df['Target Name'] == 'N2') & (df['CT'].isnull()), 'NTC_N2'] = 'passed'
        df.loc[(df['Sample Name'] == 'NTC') & (df['Target Name'] == 'N2') & (df['CT'].notnull()), 'NTC_N2'] = 'failed'
        df['NTC_RP'] = None  # initial value
        df.loc[(df['Sample Name'] == 'NTC') & (df['Target Name'] == 'RP') & (df['CT'].isnull()), 'NTC_RP'] = 'passed'
        df.loc[(df['Sample Name'] == 'NTC') & (df['Target Name'] == 'RP') & (df['CT'].notnull()), 'NTC_RP'] = 'failed'

        # Create results columns for HSC - human specimen control (extraction control)
        # df['HSC_N1'] = None  # initial value
        # df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N1') & (df['CT'].isnull()), 'HSC_N1'] = 'passed'
        # df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N1') & (df['CT'].notnull()), 'HSC_N1'] = 'failed'
        # df['HSC_N2'] = None  # initial value
        # df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N2') & (df['CT'].isnull()), 'HSC_N2'] = 'passed'
        # df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N2') & (df['CT'].notnull()), 'HSC_N2'] = 'failed'
        # df['HSC_RP'] = None  # initial value
        # df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'RP') & (df['CT'] <= ct_value),
        #        'HSC_RP'] = 'passed'
        # df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'RP') & (df['CT'] > ct_value),
        #        'HSC_RP'] = 'failed'

        # Updated - Create results columns for HSC - human specimen control (extraction control) with full sample name
        df['HSC_N1'] = None  # initial value
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['Target Name'] == 'N1') & (df['CT'].isnull()),
               'HSC_N1'] = 'passed'
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['Target Name'] == 'N1') & (df['CT'].notnull()),
               'HSC_N1'] = 'failed'
        df['HSC_N2'] = None  # initial value
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['Target Name'] == 'N2') & (df['CT'].isnull()),
               'HSC_N2'] = 'passed'
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['Target Name'] == 'N2') & (df['CT'].notnull()),
               'HSC_N2'] = 'failed'
        df['HSC_RP'] = None  # initial value
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['Target Name'] == 'RP') & (df['CT'] <= ct_value),
               'HSC_RP'] = 'passed'
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['Target Name'] == 'RP') & (df['CT'] > ct_value),
               'HSC_RP'] = 'failed'

        # Create results columns for nCoVPC - novel Coronavirus control (positive control)
        df['nCoVPC_N1'] = None  # initial value
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['Target Name'] == 'N1') & (
                    df['CT'] <= ct_value), 'nCoVPC_N1'] = 'passed'
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['Target Name'] == 'N1') & (
                    df['CT'] > ct_value), 'nCoVPC_N1'] = 'failed'
        df['nCoVPC_N2'] = None  # initial value
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['Target Name'] == 'N2') & (
                    df['CT'] <= ct_value), 'nCoVPC_N2'] = 'passed'
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['Target Name'] == 'N2') & (
                    df['CT'] > ct_value), 'nCoVPC_N2'] = 'failed'
        df['nCoVPC_RP'] = None  # initial value
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['Target Name'] == 'RP') & (
                    df['CT'] <= ct_value), 'nCoVPC_RP'] = 'passed'
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['Target Name'] == 'RP') & (
                    df['CT'] > ct_value), 'nCoVPC_RP'] = 'failed'

        # Create column for aggregate results of NTC - negative control
        df['Negative_control'] = None
        df.loc[(df['Sample Name'] == 'NTC') & (df['NTC_N1'] == 'passed')
               | (df['Sample Name'] == 'NTC') & (df['NTC_N2'] == 'passed')
               | (df['Sample Name'] == 'NTC') & (df['NTC_RP'] == 'passed'), 'Negative_control'] = 'passed'
        df.loc[(df['Sample Name'] == 'NTC') & (df['NTC_N1'] == 'failed')
               | (df['Sample Name'] == 'NTC') & (df['NTC_N2'] == 'failed')
               | (df['Sample Name'] == 'NTC') & (df['NTC_RP'] == 'failed'), 'Negative_control'] = 'failed'

        # Create column for aggregate results of HSC - extraction control
        # df['Extraction_control'] = None
        # df.loc[(df['Sample Name'] == 'HSC') & (df['HSC_N1'] == 'passed')
        #        | (df['Sample Name'] == 'HSC') & (df['HSC_N2'] == 'passed')
        #        | (df['Sample Name'] == 'HSC') & (df['HSC_RP'] == 'passed'), 'Extraction_control'] = 'passed'
        # df.loc[(df['Sample Name'] == 'HSC') & (df['HSC_N1'] == 'failed')
        #        | (df['Sample Name'] == 'HSC') & (df['HSC_N2'] == 'failed')
        #        | (df['Sample Name'] == 'HSC') & (df['HSC_RP'] == 'failed'), 'Extraction_control'] = 'failed'

        # Updated - Create column for aggregate results of HSC -extraction control
        df['Extraction_control'] = None
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['HSC_N1'] == 'passed')
               | (df['Sample Name'].str.contains("NEG", case=False)) & (df['HSC_N2'] == 'passed')
               | (df['Sample Name'].str.contains("NEG", case=False)) & (df['HSC_RP'] == 'passed'),
               'Extraction_control'] = 'passed'
        df.loc[(df['Sample Name'].str.contains("NEG", case=False)) & (df['HSC_N1'] == 'failed')
               | (df['Sample Name'].str.contains("NEG", case=False)) & (df['HSC_N2'] == 'failed')
               | (df['Sample Name'].str.contains("NEG", case=False)) & (df['HSC_RP'] == 'failed'),
               'Extraction_control'] = 'failed'

        # Create column for aggregate results of nCoVPC - positive control
        df['Positive_control'] = None
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['nCoVPC_N1'] == 'passed')
               | (df['Sample Name'] == 'nCoVPC') & (df['nCoVPC_N2'] == 'passed')
               | (df['Sample Name'] == 'nCoVPC') & (df['nCoVPC_RP'] == 'passed'), 'Positive_control'] = 'passed'
        df.loc[(df['Sample Name'] == 'nCoVPC') & (df['nCoVPC_N1'] == 'failed')
               | (df['Sample Name'] == 'nCoVPC') & (df['nCoVPC_N2'] == 'failed')
               | (df['Sample Name'] == 'nCoVPC') & (df['nCoVPC_RP'] == 'failed'), 'Positive_control'] = 'failed'

        # Sanity checks
        # print(df.loc[df['Sample Name'] == 'NTC', ['Sample Name', 'Target Name', 'CT', 'NTC_N1', 'NTC_N2',
        #                                           'NTC_RP', 'Negative_control']])
        # print(df.loc[df['Sample Name'] == 'HSC', ['Sample Name', 'Target Name', 'CT', 'HSC_N1', 'HSC_N2',
        #                                           'HSC_RP', 'Extraction_control']])
        # print(df.loc[df['Sample Name'] == 'nCoVPC', ['Sample Name', 'Target Name', 'CT', 'nCoVPC_N1', 'nCoVPC_N2',
        #                                              'nCoVPC_RP', 'Positive_control']])

        # Filter data frame to only include controls and selected columns
        # controls_filtered = df.loc[
        #     (df['Sample Name'] == 'NTC') | (df['Sample Name'] == 'HSC') | (df['Sample Name'] == 'nCoVPC')]
        # controls = controls_filtered.loc[:, ['Sample Name', 'Target Name', 'CT', 'Negative_control',
        #                                      'Extraction_control', 'Positive_control']]

        # Updated - Filter data frame to only include controls and selected columns
        controls_filtered = df.loc[
            (df['Sample Name'] == 'NTC') | (df['Sample Name'].str.contains("NEG", case=False)) |
            (df['Sample Name'] == 'nCoVPC')]
        controls = controls_filtered.loc[:, ['Sample Name', 'Target Name', 'CT', 'Negative_control',
                                             'Extraction_control', 'Positive_control']]
        # Define list of columns to join
        cols = ['Negative_control', 'Extraction_control', 'Positive_control']
        # Join selected columns into single column - 'controls_result'
        controls['controls_result'] = controls[cols].apply(lambda x: ''.join(x.dropna()), axis=1)
        # print(controls)

        # Sort controls data frame so that controls are grouped in log output.
        controls = controls.sort_values(by=['Sample Name', 'Target Name'])

        # TODO: Raise error if controls result column contains string 'failed'. Error message below.
        """
        "One or more of the above controls does not exhibit the expected performance as described. "
        "The assay may have been set up and/or executed improperly, or reagent or equipment malfunction "
        "could have occurred. Invalidate the run and re-test."
        """

        # Results interpretation
        # Create sample results column

        # This portion will handle if the NOAMP flag is present (output from QuantStudio and ViiA7 instruments)
        # Results for N1 assay
        if 'NOAMP' in df.columns:
            df.loc[(df['Target Name'] == 'N1') & (df['CT'] > ct_value) | (df['Target Name'] == 'N1') &
                   (df['CT'].isnull()), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'N1') & (df['CT'] < ct_value) & (df['NOAMP'] == "Y"), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'N1') & (df['CT'] < ct_value) & (df['NOAMP'] == "N"), 'result'] = 'positive'
        # Results for N2 assay
            df.loc[(df['Target Name'] == 'N2') & (df['CT'] > ct_value) | (df['Target Name'] == 'N2') &
                   (df['CT'].isnull()), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'N2') & (df['CT'] < ct_value) & (df['NOAMP'] == "Y"), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'N2') & (df['CT'] < ct_value) & (df['NOAMP'] == "N"), 'result'] = 'positive'
        # Results for RP assay
            df.loc[(df['Target Name'] == 'RP') & (df['CT'] > ct_value) | (df['Target Name'] == 'RP') &
                   (df['CT'].isnull()), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'RP') & (df['CT'] < ct_value) & (df['NOAMP'] == "Y"), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'RP') & (df['CT'] < ct_value) & (df['NOAMP'] == "N"), 'result'] = 'positive'

        # This portion handles instances when NOAMP flag is absent (i.e. output from 7500 instrument)
        if 'NOAMP' not in df.columns:
            df.loc[(df['Target Name'] == 'N1') & (df['CT'] > ct_value) | (df['Target Name'] == 'N1') &
                   (df['CT'].isnull()), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'N1') & (df['CT'] < ct_value), 'result'] = 'positive'
            # Results for N2 assay
            df.loc[(df['Target Name'] == 'N2') & (df['CT'] > ct_value) | (df['Target Name'] == 'N2') &
                   (df['CT'].isnull()), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'N2') & (df['CT'] < ct_value), 'result'] = 'positive'
            # Results for RP assay
            df.loc[(df['Target Name'] == 'RP') & (df['CT'] > ct_value) | (df['Target Name'] == 'RP') &
                   (df['CT'].isnull()), 'result'] = 'negative'
            df.loc[(df['Target Name'] == 'RP') & (df['CT'] < ct_value), 'result'] = 'positive'

        # # Filter for samples (exclude controls)
        # sf = df[df['Sample Name'].apply(lambda x: x not in ['NTC', 'HSC', 'nCoVPC',
        #                                                     np.NaN])].copy(deep=True).sort_values(by=['Sample Name'])

        # Updated - Drop Sample Names that appear as NaN in 7500 output
        sf_orig = df.dropna(subset=['Sample Name'])

        # Make all Sample Name values uppercase
        # sf_orig['Sample Name'] = sf_orig['Sample Name'].str.upper()

        # Updated - Filter for samples (exclude controls)
        controls_list = ['NTC', 'NEG', 'nCoVPC']
        sf = sf_orig[~sf_orig['Sample Name'].str.contains('|'.join(controls_list), case=False)]\
            .copy(deep=True).sort_values(by=['Sample Name'])

        # Sanity check
        # print(sf.head())

        # Combine 'Sample Name' and 'Target Name' for split/pivot below.
        sf['Sample_ID'] = sf['Sample Name'] + ":" + sf['Target Name']
        # Sanity check
        # print(sf.head())

        # Split and pivot
        sf[['Sample_ID', 'assay']] = sf['Sample_ID'].str.split(':', expand=True)
        sf = sf.pivot('Sample_ID', 'assay', 'result').add_prefix('Result_')
        # Sanity check
        # print(sf)

        # 2019-nCoV rRT-PCR Diagnostic Panel Results Interpretation Guide (page 32 of reference file)
        sf.loc[(sf['Result_N1'] == 'positive') & (sf['Result_N2'] == 'positive') & (sf['Result_RP'].notnull()),
               'Result_Interpretation'] = '2019-nCoV detected'
        sf.loc[(sf['Result_N1'] == 'positive') & (sf['Result_N2'] == 'negative') & (sf['Result_RP'].notnull()),
               'Result_Interpretation'] = 'Inconclusive Result'
        sf.loc[(sf['Result_N1'] == 'negative') & (sf['Result_N2'] == 'positive') & (sf['Result_RP'].notnull()),
               'Result_Interpretation'] = 'Inconclusive Result'
        sf.loc[(sf['Result_N1'] == 'negative') & (sf['Result_N2'] == 'negative') & (sf['Result_RP'] == 'positive'),
               'Result_Interpretation'] = '2019-nCoV not detected'
        sf.loc[(sf['Result_N1'] == 'negative') & (sf['Result_N2'] == 'negative') & (sf['Result_RP'] == 'negative'),
               'Result_Interpretation'] = 'Invalid Result'

        # Create 'Results_Interpretation' column
        sf.loc[(sf['Result_Interpretation'] == '2019-nCoV detected'), 'Report'] = 'Positive 2019-nCoV'
        sf.loc[(sf['Result_Interpretation'] == 'Inconclusive Result'), 'Report'] = 'Inconclusive'
        sf.loc[(sf['Result_Interpretation'] == '2019-nCoV not detected'), 'Report'] = 'Not Detected'
        sf.loc[(sf['Result_Interpretation'] == 'Invalid Result'), 'Report'] = 'Invalid'

        # Create 'Actions' column
        sf.loc[(sf['Report'] == 'Positive 2019-nCoV'), 'Actions'] = 'Report results to CDC and sender'
        sf.loc[(sf['Report'] == 'Inconclusive'), 'Actions'] = 'Repeat testing of nucleic acid and/or re-extract and ' \
                                                              'repeat rRT-PCR. If the repeated result remains ' \
                                                              'inconclusive contact your State Public Health ' \
                                                              'Laboratory or CDC for instructions for transfer ' \
                                                              'of the specimen or further guidance.'
        sf.loc[(sf['Report'] == 'Not Detected'), 'Actions'] = 'Report results to sender. Consider testing for other ' \
                                                              'respiratory viruses.'
        sf.loc[(sf['Report'] == 'Invalid'), 'Actions'] = 'Repeat extraction and rRT-PCR. If the repeated result ' \
                                                         'remains invalid consider collecting a new specimen from ' \
                                                         'the patient.'
        # Reset index
        sf = sf.reset_index()

        # Check - Write out final results file.
        # sf.to_csv("final_results_test.csv", sep=',', index=False)

        # Prepare the outpath for the processed data using a timestamp
        timestr = time.strftime('%m_%d_%Y_%H_%M_%S')

        # This portion works for Unix systems - see section below for Windows.
        outname = os.path.split(path)
        outname1 = outname[0]
        outfilename = outname[1]
        # new_base = timestr + '_covid_results.csv'
        # # original
        # outpath = outname1 + '/' + new_base
        # sf.to_csv(outpath, sep=",", index=False)

        # For Windows-based file paths
        mypath = os.path.abspath(os.path.dirname(path))
        newpath = os.path.join(mypath, '../../processed')
        normpath = os.path.normpath(newpath)
        new_base = timestr + '_covid_results.csv'
        sf.to_csv(normpath + '\\' + new_base, sep=",", index=False)

        # Experiment details
        # runinfo = pd.read_excel(path, sheet_name='Results', skiprows=28, header=None, nrows=8)

        # To accommodate either QuantStudio or ViiA7
        info_orig = pd.read_excel(path, sheet_name="Results", header=None)
        for row2 in range(info_orig.shape[0]):
            for col2 in range(info_orig.shape[1]):
                if info_orig.iat[row2, col2] == "Experiment File Name":
                    row_start_2 = row2
                    break
        # Subset raw file for only portion below "Well" and remainder of header
        runinfo = info_orig[row_start_2:(row_start_2+9)]

        # Reset index
        runinfo.reset_index(drop=True)

        # This portion works for Unix systems - see section below for Windows.
        # Log file
        # original
        # log_base = timestr + '_covid_output.log'
        # log_filename = outname1 + '/' + log_base

        # For Windows-based file paths
        newlogpath = os.path.join(mypath, '../../processed/logs')
        normlogpath = os.path.normpath(newlogpath)
        log_base = timestr + '_covid_output.log'
        log_filename = normlogpath + '\\' + log_base
 
        # Define log file parameters
        logging.basicConfig(filename=log_filename, level=logging.DEBUG, format='%(asctime)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S')
        # Info for log file
        logging.info(' Name of input file: ' + outfilename)
        logging.info('\n')
        logging.info('Run information: ')
        logging.info('\n' + runinfo.loc[:, [0, 1]].to_string(index=False, header=False))
        logging.info('\n')
        logging.info(' Number of controls run: ' + str(len(controls['Sample Name'].unique().tolist())))
        logging.info(' Controls run: ' + str(controls['Sample Name'].unique()))
        logging.info('\n')
        logging.info(' Results of controls: ')
        logging.info('\n' + controls.to_string())
        logging.warning('\t')
        logging.warning(
            str('If any of the above controls do not exhibit the expected performance as described, the assay may '
                'have been set up and/or executed improperly, or reagent or equipment malfunction could have '
                'occurred. Invalidate the run and re-test.'))
        logging.warning('\n')
        logging.info(' Number of samples run: ' + str(len(sf['Sample_ID'].unique().tolist())))
        logging.info('Samples run: ')
        logging.info(str(sf['Sample_ID'].unique()))

        messagebox.showinfo("Complete", "Data Processing Complete!")

# Replaces Meditech to BSI R script that converts Meditech report to BSI-friendly upload file (PGx formatting)
    def bsiprocess(self):
        filepath = filedialog.askopenfilename()

        # need to obtain 'Study ID', 'Current Label', 'Account', 'Subject ID', 'Med Rec', 'Date Collected',
        # 'Date Received', 'Gender', 'DOB', 'First Name', 'Last Name', 'Specimen'
        out_list = []

        with open(filepath) as fp:
            for cnt, line in enumerate(fp):
                # print("Line {}: {}".format(cnt, line))
                if 'ACCT' in line:
                    patient_dict = {}
                    # split the file line by the ':' character, should result in a list of 5 elements
                    acct_line = line.split(':')

                    # get the name information
                    list_split = acct_line[1].split(' ')
                    lastFirst = list_split[1]
                    middle = list_split[2]
                    last, first = lastFirst.split(',')
                    # this will fix the nickname issue, will append nickname to the middle name with a space in the
                    # middle
                    if list_split[3] != '':
                        nick_name = list_split[3]
                        middle_base = middle.strip(' ')
                        middle_out = middle_base + ' ' + nick_name
                        patient_dict['middle'] = middle_out
                    else:
                        patient_dict['middle'] = middle.strip(' ')

                    patient_dict['first'] = first.strip(' ')
                    patient_dict['Last Name'] = last.strip(' ')

                    # get the account/current label information
                    acctLine = acct_line[2]
                    acctNum = acctLine.split(' ')[1]
                    patient_dict['Account'] = acctNum.strip(' ')
                    patient_dict['Current Label'] = acctNum.strip(' ')

                    # get the subject id/med rec info
                    medRecLine = acct_line[4]
                    medRec = medRecLine.strip('\n')
                    patient_dict['Subject ID'] = medRec.strip(' ')
                    patient_dict['Med Rec'] = medRec.strip(' ')

                elif 'AGE/SX' in line:
                    ageLine = line.split(':')
                    ageSex = ageLine[1].split(' ')[1]
                    sex = ageSex.split('/')[1]
                    patient_dict['Gender'] = sex.strip(' ')

                elif 'DOB' in line:
                    dobLine = line.split(':')
                    patient_dict['DOB'] = dobLine[2].split(' ')[4]

                elif 'SPEC' in line:
                    specLine = line.split(':')
                    spec1 = specLine[1]
                    spec1 = spec1.strip(' ')
                    spec2 = specLine[2].split(' ')[0]
                    spec2 = spec2.strip(' ')
                    patient_dict['Specimen'] = spec1 + ':' + spec2
                    # get the collection date
                    dob = specLine[3].split(' ')[1]
                    patient_dict['Date Collected'] = dob.strip(' ')

                elif 'RECD' in line:
                    recdLine = line.split(':')[1]
                    patient_dict['Date Received'] = recdLine.split(' ')[1]

                    out_list.append(patient_dict)

                else:
                    pass

        # make a dataframe from the output and clean it up a bit
        df_patient = pd.DataFrame(out_list)

        df_patient['FirstMiddle'] = df_patient['first'] + ' ' + df_patient['middle']
        df_patient = df_patient.drop(['first', 'middle'], axis=1)
        df_patient = df_patient.rename(columns={'FirstMiddle': 'First Name'})

        df_patient['DOB'] = pd.to_datetime(df_patient['DOB'])
        df_patient['DOB'] = df_patient['DOB'].dt.strftime('%m/%d/%Y')

        df_patient['Date Collected'] = pd.to_datetime(df_patient['Date Collected'])
        df_patient['Date Collected'] = df_patient['Date Collected'].dt.strftime('%m/%d/%Y %H:%M')

        df_patient['Date Received'] = pd.to_datetime(df_patient['Date Received'])
        df_patient['Date Received'] = df_patient['Date Received'].dt.strftime('%m/%d/%Y %H:%M')

        df_patient['Study ID'] = 'PGX'

        dfOut = df_patient[['Study ID', 'Current Label', 'Account', 'Subject ID', 'Med Rec', 'Date Collected',
                            'Date Received', 'Gender', 'DOB', 'First Name', 'Last Name', 'Specimen']]

        outname = os.path.split(filepath)
        filename = outname[1]
        filenamenoext = filename[:-4]

        # For Windows-based file paths
        mypath = os.path.abspath(os.path.dirname(filepath))

        dfOut.to_csv(mypath + '\\' + filenamenoext + '_BSIconverted.txt', sep="\t", index=False)

        messagebox.showinfo("Complete", "Data Processing Complete!")

    ## Make LIMS-friendly output
    def limsprocess(self):
        # Ingest input file
        # ask the user for an input read in the file selected by the user
        path = filedialog.askopenfilename()

        # To accommodate either QuantStudio or ViiA7
        df_orig = pd.read_excel(path, sheet_name="Results", header=None)
        for row in range(df_orig.shape[0]):
            for col in range(df_orig.shape[1]):
                if df_orig.iat[row, col] == "Well":
                    row_start = row
                    break

        # Subset raw file for only portion below "Well" and remainder of header
        df = df_orig[row_start:]

        # Header exists in row 1, make new header
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header

        # Adding a new line to handle the 'Cт' present in the header of the output file from the 7500 instrument
        df.columns = df.columns.str.replace('Cт', 'CT')

        # Convert 'undetermined' to 'NaN' for 'CT' column
        df['CT'] = df.loc[:, 'CT'].apply(pd.to_numeric, errors='coerce')

        # TODO: DEFINE CT VALUE HERE
        ct_value = 40.00

        # New code
        pt = df.pivot(index="Sample Name", columns="Target Name", values=["CT", "NOAMP"])
        new_df = pd.DataFrame(pt.to_records()).rename(columns={'Target Name': 'index'})

        newcols = {"Sample Name": "Sample_Name", "('CT', 'N1')": "N1_CT", "('CT', 'N2')": "N2_CT",
                   "('CT', 'RP')": "RP_CT", "('NOAMP', 'N1')": "N1_NOAMP", "('NOAMP', 'N2')": "N2_NOAMP",
                   "('NOAMP', 'RP')": "RP_NOAMP"}
        new_df.columns = new_df.columns.map(newcols)

        new_df['N1_Result'] = np.nan
        new_df.loc[(new_df['N1_CT'].isnull()), 'N1_Result'] = "negative"
        new_df.loc[(new_df['N1_CT'].isnull()) & (new_df['N1_NOAMP'] == "Y"), 'N1_Result'] = 'negative'
        new_df.loc[(new_df['N1_CT'] > ct_value), 'N1_Result'] = 'negative'
        new_df.loc[(new_df['N1_CT'] <= ct_value) & (new_df['N1_NOAMP'] == "Y"), 'N1_Result'] = 'negative'
        new_df.loc[(new_df['N1_CT'] <= ct_value) & (new_df['N1_NOAMP'] == "N"), 'N1_Result'] = 'positive'

        new_df['N2_Result'] = np.nan
        new_df.loc[(new_df['N2_CT'].isnull()), 'N2_Result'] = "negative"
        new_df.loc[(new_df['N2_CT'].isnull()) & (new_df['N2_NOAMP'] == "Y"), 'N2_Result'] = 'negative'
        new_df.loc[(new_df['N2_CT'] > ct_value), 'N2_Result'] = 'negative'
        new_df.loc[(new_df['N2_CT'] <= ct_value) & (new_df['N2_NOAMP'] == "Y"), 'N2_Result'] = 'negative'
        new_df.loc[(new_df['N2_CT'] <= ct_value) & (new_df['N2_NOAMP'] == "N"), 'N2_Result'] = 'positive'

        new_df['RP_Result'] = np.nan
        new_df.loc[(new_df['RP_CT'].isnull()), 'RP_Result'] = "negative"
        new_df.loc[(new_df['RP_CT'].isnull()) & (new_df['RP_NOAMP'] == "Y"), 'RP_Result'] = 'negative'
        new_df.loc[(new_df['RP_CT'] > ct_value), 'RP_Result'] = 'negative'
        new_df.loc[(new_df['RP_CT'] <= ct_value) & (new_df['RP_NOAMP'] == "Y"), 'RP_Result'] = 'negative'
        new_df.loc[(new_df['RP_CT'] <= ct_value) & (new_df['RP_NOAMP'] == "N"), 'RP_Result'] = 'positive'

        # Assess controls
        # Expected performance of controls
        """
        ControlType   ExternalControlName Monitors        2019nCoV_N1 2019nCOV_N2 RnaseP  ExpectedCt
        Positive      nCoVPC              Rgt Failure     +           +           +       <40
        Negative      NTC                 Contamination   -           -           -       None
        Extraction    HSC                 Extraction      -           -           +       <40

        If any of the above controls do not exhibit the expected performance as described, the assay may have been set
        up and/or executed improperly, or reagent or equipment malfunction could have occurred. Invalidate the run and
        re-test.
        """
        new_df['Neg_ctrl'] = np.nan
        new_df.loc[((new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N1_CT'].isnull())) & (
                    (new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N2_CT'].isnull())) & (
                               (new_df['Sample_Name'].str.contains("NTC", case=False)) & (
                           new_df['RP_CT'].isnull())), 'Neg_ctrl'] = "passed"
        new_df.loc[((new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N1_CT'].notnull())) | (
                    (new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N2_CT'].notnull())) | (
                               (new_df['Sample_Name'].str.contains("NTC", case=False)) & (
                           new_df['RP_CT'].notnull())), 'Neg_ctrl'] = "failed"

        new_df['Ext_ctrl'] = np.nan
        new_df.loc[((new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N1_CT'].isnull())) & (
                    (new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N2_CT'].isnull())) & (
                               (new_df['Sample_Name'].str.contains("NEG", case=False)) & (
                                   new_df['RP_CT'] <= ct_value)), 'Ext_ctrl'] = "passed"
        new_df.loc[((new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N1_CT'].notnull())) | (
                    (new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N2_CT'].notnull())) | (
                               (new_df['Sample_Name'].str.contains("NEG", case=False)) & (
                                   new_df['RP_CT'] > ct_value)), 'Ext_ctrl'] = "failed"

        new_df['Pos_ctrl'] = np.nan
        new_df.loc[((new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N1_CT'] <= ct_value)) & (
                    (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N2_CT'] <= ct_value)) & (
                               (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (
                                   new_df['RP_CT'] <= ct_value)), 'Pos_ctrl'] = "passed"
        new_df.loc[((new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N1_CT'] > ct_value)) | (
                    (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N2_CT'] > ct_value)) | (
                               (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (
                                   new_df['RP_CT'] > ct_value)), 'Pos_ctrl'] = "failed"

        control_cols = ['Neg_ctrl', 'Ext_ctrl', 'Pos_ctrl']
        new_df['controls_result'] = new_df[control_cols].apply(lambda x: ''.join(x.dropna()), axis=1)

        new_df['controls_result'] = new_df['controls_result'].replace(r'^\s*$', np.nan, regex=True)

        new_df = new_df.sort_values(by='Sample_Name')

        new_df = new_df.drop(['Neg_ctrl', 'Ext_ctrl', 'Pos_ctrl'], axis=1)

        # 2019-nCoV rRT-PCR Diagnostic Panel Results Interpretation Guide (page 32 of reference file)
        new_df.loc[(new_df['N1_Result'] == 'positive') & (new_df['N2_Result'] == 'positive') &
                   (new_df['RP_Result'].notnull()),
               'Result_Interpretation'] = 'Positive'
        new_df.loc[(new_df['N1_Result'] == 'positive') & (new_df['N2_Result'] == 'negative') &
                   (new_df['RP_Result'].notnull()),
               'Result_Interpretation'] = 'Inconclusive'
        new_df.loc[(new_df['N1_Result'] == 'negative') & (new_df['N2_Result'] == 'positive') &
                   (new_df['RP_Result'].notnull()),
               'Result_Interpretation'] = 'Inconclusive'
        new_df.loc[(new_df['N1_Result'] == 'negative') & (new_df['N2_Result'] == 'negative') &
                   (new_df['RP_Result'] == 'positive'),
               'Result_Interpretation'] = 'Not Detected'
        new_df.loc[(new_df['N1_Result'] == 'negative') & (new_df['N2_Result'] == 'negative') &
                   (new_df['RP_Result'] == 'negative'),
               'Result_Interpretation'] = 'Invalid'

        new_df = new_df[
            ['Sample_Name', 'N1_CT', 'N1_NOAMP', 'N1_Result', 'N2_CT', 'N2_NOAMP', 'N2_Result', 'RP_CT', 'RP_NOAMP',
             'RP_Result', 'Result_Interpretation', 'controls_result']]

        new_df['N1_CT'].fillna('Undetermined', inplace=True)
        new_df['N2_CT'].fillna('Undetermined', inplace=True)
        new_df['RP_CT'].fillna('Undetermined', inplace=True)

        # Automatically read in panel data file that is updated every 4 hours
        path2 = "J:/AIHG/AIHG_Covid/AIHG_Covid_Orders/AIHG_Covid_Orders.csv"
        paneldf = pd.read_csv(path2, header=0)

        # Merge results with panel id file
        merge = pd.merge(new_df, paneldf, left_on="Sample_Name", right_on="AccountNumber", how="left")

        merge_clean = merge[["PanelID", "Sample_Name", "N1_CT", "N1_NOAMP", "N1_Result", "N2_CT", "N2_NOAMP",
                             "N2_Result", "RP_CT", "RP_NOAMP", "RP_Result", "Result_Interpretation", "controls_result"]]

        # Prepare the outpath for the processed data using a timestamp
        timestr = time.strftime('%m_%d_%Y_%H_%M_%S')

        # Break file path/name to extract barcode from file name
        outname = os.path.split(path)
        dir_path = outname[0]
        plate_barcode = outname[1]

        # For Windows-based file paths
        mypath = os.path.abspath(os.path.dirname(path))
        newpath = os.path.join(mypath, '../../processed/output_for_LIMS')
        normpath = os.path.normpath(newpath)

        # Replace new_base with plate_barcode
        # new_base = timestr + '_covid_results.csv'

        merge_clean.to_csv(normpath + '\\' + plate_barcode + '.csv', sep=",", index=False)

        messagebox.showinfo("Complete", "Data Processing Complete!")


# Make Meditech-friendly output including panel ID.
    def meditechprocess(self):
        # Ingest input file
        # ask the user for an input read in the file selected by the user
        messagebox.showinfo("Select results file", "Select RT-PCR file to analyze")
        path = filedialog.askopenfilename()

        # To accommodate either QuantStudio or ViiA7
        df_orig = pd.read_excel(path, sheet_name="Results", header=None)
        for row in range(df_orig.shape[0]):
            for col in range(df_orig.shape[1]):
                if df_orig.iat[row, col] == "Well":
                    row_start = row
                    break

        # Subset raw file for only portion below "Well" and remainder of header
        df = df_orig[row_start:]

        # Header exists in row 1, make new header
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header

        # Adding a new line to handle the 'Cт' present in the header of the output file from the 7500 instrument
        df.columns = df.columns.str.replace('Cт', 'CT')

        # Convert 'undetermined' to 'NaN' for 'CT' column
        df['CT'] = df.loc[:, 'CT'].apply(pd.to_numeric, errors='coerce')

        # TODO: DEFINE CT VALUE HERE
        ct_value = 40.00

        # New code
        pt = df.pivot(index="Sample Name", columns="Target Name", values=["CT", "NOAMP"])
        new_df = pd.DataFrame(pt.to_records()).rename(columns={'Target Name': 'index'})

        newcols = {"Sample Name": "Sample_Name", "('CT', 'N1')": "N1_CT", "('CT', 'N2')": "N2_CT",
                   "('CT', 'RP')": "RP_CT", "('NOAMP', 'N1')": "N1_NOAMP", "('NOAMP', 'N2')": "N2_NOAMP",
                   "('NOAMP', 'RP')": "RP_NOAMP"}
        new_df.columns = new_df.columns.map(newcols)

        new_df['N1_Result'] = np.nan
        new_df.loc[(new_df['N1_CT'].isnull()), 'N1_Result'] = "negative"
        new_df.loc[(new_df['N1_CT'].isnull()) & (new_df['N1_NOAMP'] == "Y"), 'N1_Result'] = 'negative'
        new_df.loc[(new_df['N1_CT'] > ct_value), 'N1_Result'] = 'negative'
        new_df.loc[(new_df['N1_CT'] <= ct_value) & (new_df['N1_NOAMP'] == "Y"), 'N1_Result'] = 'negative'
        new_df.loc[(new_df['N1_CT'] <= ct_value) & (new_df['N1_NOAMP'] == "N"), 'N1_Result'] = 'positive'

        new_df['N2_Result'] = np.nan
        new_df.loc[(new_df['N2_CT'].isnull()), 'N2_Result'] = "negative"
        new_df.loc[(new_df['N2_CT'].isnull()) & (new_df['N2_NOAMP'] == "Y"), 'N2_Result'] = 'negative'
        new_df.loc[(new_df['N2_CT'] > ct_value), 'N2_Result'] = 'negative'
        new_df.loc[(new_df['N2_CT'] <= ct_value) & (new_df['N2_NOAMP'] == "Y"), 'N2_Result'] = 'negative'
        new_df.loc[(new_df['N2_CT'] <= ct_value) & (new_df['N2_NOAMP'] == "N"), 'N2_Result'] = 'positive'

        new_df['RP_Result'] = np.nan
        new_df.loc[(new_df['RP_CT'].isnull()), 'RP_Result'] = "negative"
        new_df.loc[(new_df['RP_CT'].isnull()) & (new_df['RP_NOAMP'] == "Y"), 'RP_Result'] = 'negative'
        new_df.loc[(new_df['RP_CT'] > ct_value), 'RP_Result'] = 'negative'
        new_df.loc[(new_df['RP_CT'] <= ct_value) & (new_df['RP_NOAMP'] == "Y"), 'RP_Result'] = 'negative'
        new_df.loc[(new_df['RP_CT'] <= ct_value) & (new_df['RP_NOAMP'] == "N"), 'RP_Result'] = 'positive'

        # Assess controls
        # Expected performance of controls
        """
        ControlType   ExternalControlName Monitors        2019nCoV_N1 2019nCOV_N2 RnaseP  ExpectedCt
        Positive      nCoVPC              Rgt Failure     +           +           +       <40
        Negative      NTC                 Contamination   -           -           -       None
        Extraction    HSC                 Extraction      -           -           +       <40

        If any of the above controls do not exhibit the expected performance as described, the assay may have been set
        up and/or executed improperly, or reagent or equipment malfunction could have occurred. Invalidate the run and
        re-test.
        """
        new_df['Neg_ctrl'] = np.nan
        new_df.loc[((new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N1_CT'].isnull())) & (
                    (new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N2_CT'].isnull())) & (
                               (new_df['Sample_Name'].str.contains("NTC", case=False)) & (
                           new_df['RP_CT'].isnull())), 'Neg_ctrl'] = "passed"
        new_df.loc[((new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N1_CT'].notnull())) | (
                    (new_df['Sample_Name'].str.contains("NTC", case=False)) & (new_df['N2_CT'].notnull())) | (
                               (new_df['Sample_Name'].str.contains("NTC", case=False)) & (
                           new_df['RP_CT'].notnull())), 'Neg_ctrl'] = "failed"

        new_df['Ext_ctrl'] = np.nan
        new_df.loc[((new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N1_CT'].isnull())) & (
                    (new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N2_CT'].isnull())) & (
                               (new_df['Sample_Name'].str.contains("NEG", case=False)) & (
                                   new_df['RP_CT'] <= ct_value)), 'Ext_ctrl'] = "passed"
        new_df.loc[((new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N1_CT'].notnull())) | (
                    (new_df['Sample_Name'].str.contains("NEG", case=False)) & (new_df['N2_CT'].notnull())) | (
                               (new_df['Sample_Name'].str.contains("NEG", case=False)) & (
                                   new_df['RP_CT'] > ct_value)), 'Ext_ctrl'] = "failed"

        new_df['Pos_ctrl'] = np.nan
        new_df.loc[((new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N1_CT'] <= ct_value)) & (
                    (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N2_CT'] <= ct_value)) & (
                               (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (
                                   new_df['RP_CT'] <= ct_value)), 'Pos_ctrl'] = "passed"
        new_df.loc[((new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N1_CT'] > ct_value)) | (
                    (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (new_df['N2_CT'] > ct_value)) | (
                               (new_df['Sample_Name'].str.contains("nCoVPC", case=False)) & (
                                   new_df['RP_CT'] > ct_value)), 'Pos_ctrl'] = "failed"

        control_cols = ['Neg_ctrl', 'Ext_ctrl', 'Pos_ctrl']
        new_df['controls_result'] = new_df[control_cols].apply(lambda x: ''.join(x.dropna()), axis=1)

        new_df['controls_result'] = new_df['controls_result'].replace(r'^\s*$', np.nan, regex=True)

        new_df = new_df.sort_values(by='Sample_Name')

        new_df = new_df.drop(['Neg_ctrl', 'Ext_ctrl', 'Pos_ctrl'], axis=1)

        # 2019-nCoV rRT-PCR Diagnostic Panel Results Interpretation Guide (page 32 of reference file)
        new_df.loc[(new_df['N1_Result'] == 'positive') & (new_df['N2_Result'] == 'positive') &
                   (new_df['RP_Result'].notnull()),
               'Result_Interpretation'] = 'Positive'
        new_df.loc[(new_df['N1_Result'] == 'positive') & (new_df['N2_Result'] == 'negative') &
                   (new_df['RP_Result'].notnull()),
               'Result_Interpretation'] = 'Inconclusive'
        new_df.loc[(new_df['N1_Result'] == 'negative') & (new_df['N2_Result'] == 'positive') &
                   (new_df['RP_Result'].notnull()),
               'Result_Interpretation'] = 'Inconclusive'
        new_df.loc[(new_df['N1_Result'] == 'negative') & (new_df['N2_Result'] == 'negative') &
                   (new_df['RP_Result'] == 'positive'),
               'Result_Interpretation'] = 'Not Detected'
        new_df.loc[(new_df['N1_Result'] == 'negative') & (new_df['N2_Result'] == 'negative') &
                   (new_df['RP_Result'] == 'negative'),
               'Result_Interpretation'] = 'Invalid'

        new_df = new_df[
            ['Sample_Name', 'N1_CT', 'N1_NOAMP', 'N1_Result', 'N2_CT', 'N2_NOAMP', 'N2_Result', 'RP_CT', 'RP_NOAMP',
             'RP_Result', 'Result_Interpretation', 'controls_result']]

        # Create a df of only samples (exclude controls)
        controls_list = ['NTC', 'NEG', 'nCoVPC']

        samples = new_df[~new_df['Sample_Name'].str.contains('|'.join(controls_list), case=False)] \
            .copy(deep=True).sort_values(by=['Sample_Name'])

        # Automatically read in panel data file that is updated every 4 hours
        path2 = "J:/AIHG/AIHG_Covid/AIHG_Covid_Orders/AIHG_Covid_Orders.csv"
        paneldf = pd.read_csv(path2, header=0)

        # Merge results with panel id file
        merge = pd.merge(samples, paneldf, left_on="Sample_Name", right_on="AccountNumber", how="left")

        # Add placeholder columns
        merge["COVID19S.P"] = ""
        merge["COVID19S.SRC"] = ""
        merge["COVID19S.SYM"] = ""

        # Select only columns of interest
        merge = merge[['PanelID', 'Sample_Name', 'N1_Result', 'N2_Result', 'RP_Result', 'COVID19S.P', 'COVID19S.SRC',
                       'COVID19S.SYM', 'Result_Interpretation']]

        # Adjust column names
        merge.rename(columns={'Sample_Name':'AccountNumber', 'N1_Result':'COVID.N1', 'N2_Result':'COVID.N2',
                              'RP_Result':"COVID.RP", 'Result_Interpretation':'COVID19S.T'}, inplace=True)

        # Capitalize negative/positive in N1/N2/RP Results fields
        merge['COVID.N1'] = merge['COVID.N1'].str.capitalize()
        merge['COVID.N2'] = merge['COVID.N2'].str.capitalize()
        merge['COVID.RP'] = merge['COVID.RP'].str.capitalize()

        # controls df for log file
        controls_filtered = new_df[new_df['Sample_Name'].str.contains('|'.join(controls_list), case=False)]\
            .copy(deep=True).sort_values(by=['Sample_Name'])

        # For output
        outname = os.path.split(path)
        outname1 = outname[0]
        outfilename = outname[1]

        # Prepare the outpath for the processed data using a timestamp
        meditech_timestr = time.strftime('%Y%m%d%H%M')

        # For Windows-based file paths
        mypath = os.path.abspath(os.path.dirname(path))
        newpath = os.path.join(mypath, '../../processed/output_for_Meditech')
        normpath = os.path.normpath(newpath)
        new_base = meditech_timestr + '_COVID19S.csv'
        merge.to_csv(normpath + '\\' + new_base, sep=",", index=False)

        info_orig = pd.read_excel(path, sheet_name="Results", header=None)
        for row2 in range(info_orig.shape[0]):
            for col2 in range(info_orig.shape[1]):
                if info_orig.iat[row2, col2] == "Experiment File Name":
                    row_start_2 = row2
                    break
        # Subset raw file for only portion below "Well" and remainder of header
        runinfo = info_orig[row_start_2:(row_start_2+9)]

        # Reset index
        runinfo.reset_index(drop=True)

        # For Windows-based file paths
        newlogpath = os.path.join(mypath, '../../processed/logs')
        normlogpath = os.path.normpath(newlogpath)
        log_base = meditech_timestr + '_Meditech_covid_output.log'
        log_filename = normlogpath + '\\' + log_base

        # Define log file parameters
        logging.basicConfig(filename=log_filename, level=logging.DEBUG, format='%(asctime)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S')
        # Info for log file
        logging.info(' Name of input file: ' + outfilename)
        logging.info('\n')
        logging.info('Run information: ')
        logging.info('\n' + runinfo.loc[:, [0, 1]].to_string(index=False, header=False))
        logging.info('\n')
        logging.info(' Number of controls run: ' + str(len(controls_filtered['Sample_Name'].unique().tolist())))
        logging.info(' Controls run: ' + str(controls_filtered['Sample_Name'].unique()))
        logging.info('\n')
        logging.info(' Results of controls: ')
        logging.info('\n' + controls_filtered.to_string())
        logging.warning('\t')
        logging.warning(
            str('If any of the above controls do not exhibit the expected performance as described, the assay may '
                'have been set up and/or executed improperly, or reagent or equipment malfunction could have '
                'occurred. Invalidate the run and re-test.'))
        logging.warning('\n')
        logging.info(' Number of samples run: ' + str(len(samples['Sample_Name'].unique().tolist())))
        logging.info('Samples run: ')
        logging.info(str(samples['Sample_Name'].unique()))

        messagebox.showinfo("Complete", "Data Processing Complete!")

## Convert Meditech to BSI file to BSI-friendly version (COVID formatting)
    def covidbsiprocess(self):
        pathbsi = filedialog.askopenfilename()
        # read in file - output from Meditech to BSI script
        current = pd.read_csv(pathbsi, sep="\t", header=0)

        # Replace 'PGX' in Study ID field with 'COVID19'
        current['Study ID'].replace('PGX', 'COVID19', inplace=True)

        # Create redundant columns (desired for BSI upload)
        current['First'] = current['First Name']
        current['Last'] = current['Last Name']
        current['Date of Birth'] = current['DOB']

        # Reorder columns
        current = current[['Study ID', 'Current Label', 'Account', 'Subject ID', 'Med Rec', 'Date Collected',
                           'Date Received', 'Gender', 'DOB', 'Date of Birth', 'First Name', 'First', 'Last Name',
                           'Last', 'Specimen']]

        # Sort by Specimen ID
        current.sort_values(by="Specimen", inplace=True)

        # Get path as string and create new base for file name
        outnamebsi = os.path.split(pathbsi)
        outname1bsi = outnamebsi[0]
        outfilenamebsi = outnamebsi[1]
        bsicleanname = outfilenamebsi[:-4]
        bsi_base = "_covid_BSI"
        newname = outname1bsi + '\\' + bsicleanname + bsi_base + ".xlsx"

        # dataframe to Excel
        writer = ExcelWriter(newname)
        current.to_excel(writer, 'Sheet1', index=False)
        writer.save()

        messagebox.showinfo("Complete", "File Successfully Converted for BSI!")

    #  TODO: Add statsprocess
    # def statsprocess(self):
    #     pathstats = filedialog.askopenfilenames()
    #     filelist = root.tk.splitlist(pathstats)
    #     files_xls = [f for f in filelist if f[-3:] == 'xls']
    #     list = []
    #     for file in filelist:
    #         list.append(os.path.split(file)[1])

    # This section works off of xls output from QuantStudio, Viia, and 7500 Fast instruments
    #     dir = filedialog.askdirectory()
    #     files_xls2 = [f for f in os.listdir(dir) if f.endswith('xls')]
    #
    #     fulldf = pd.DataFrame()
    #     for x in files_xls2:
    #         df_orig = pd.read_excel(x, sheet_name="Results", header=None)
    #         for row in range(df_orig.shape[0]):
    #             for col in range(df_orig.shape[1]):
    #                 if df_orig.iat[row, col] == "Well":
    #                     row_start = row
    #                     break
    #
    #         # Subset raw file for only portion below "Well" and remainder of header
    #         df = df_orig[row_start:]
    #
    #         # Take all but column names
    #         df = df[1:]
    #
    #         # This will not work because there will not be column names at this point.
    #         # Convert 'undetermined' to 'NaN' for 'CT' column
    #         # df['CT'] = df.loc[:, 'CT'].apply(pd.to_numeric, errors='coerce')
    #
    #         fulldf = fulldf.append(df)
    #
    #     fulldf.reset_index()

    # TODO: ADD dirplot - Plot all results (monthly and weekly)
    def dirplot(self):
        messagebox.showinfo("Select directory", "Select 'resulting_completed' directory")
        dir = filedialog.askdirectory()
        files_csv = [f for f in os.listdir(dir) if f.endswith('csv')]

        file_list = list()

        for file in files_csv:
            df = pd.read_csv(dir + "\\" + file, sep=",", header=None, skiprows=1)
            df['filename'] = file
            file_list.append(df)

        compiled_results = pd.concat(file_list, axis=0, ignore_index=True)

        # Remove rows full of NA's
        compiled_results.dropna(axis=1, how='all', inplace=True)

        compiled_results.columns = ['Sample_ID', 'Result_N1', 'Result_N2', 'Result_RP', 'Result_Interpretation',
                                    'Report', 'Actions', 'Filename']

        # Drop missings if 'Sample_ID' column is blank
        compiled_results.dropna(axis=0, subset=['Sample_ID'], inplace=True)

        # drop missings if 'Result_Interpretation' column is blank
        compiled_results.dropna(axis=0, subset=['Result_Interpretation'], inplace=True)

        exclude_list = ['LOW1', 'LOW2', 'LOW3', 'MID1', 'MID2', 'MID3', 'Mod-1_No_90', 'Mod-2_No_90', 'Mod-3_No_90',
                        'Low-1_No_90', 'Low-2_No_90', 'Low-3_No_90', 'Low_1', 'Low_2', 'Low_3', 'Mod_1', 'Mod_2',
                        'Mod_3', 'State1_04022020', 'State2_04022020', 'State3_04022020', 'State4_04022020',
                        'State5_04022020', 'State6_04022020', 'State7_04022020', 'State8_04022020', 'State_1',
                        '032420-7-M', '032420-4-M', '032420-3-M', 'Low-4_QS', 'Low-5_QS', 'Low-6_QS', 'Mod-4_QS',
                        'Mod-5_QS', 'Mod-6_QS', 'Low-1_QS', 'Low-2_QS', 'Low-3_QS', 'Mod-1_QS', 'Mod-2_QS', 'Mod-3_QS',
                        'H2O1', 'H2O2', 'H2O3', '0.16_A_Validation', '0.16_B_Validation', '0.8_A_Validation',
                        '0.8_B_Validation', '032420-5-M', '032420-8-M', '_NEG_', '20_A_Validation', '20_B_Validation',
                        '40_state3', '4_A_Validation', '4_B_Validation', 'H2O_1', 'H2O_2', 'H2O_3']

        cr1 = compiled_results[~compiled_results['Sample_ID'].str.contains('|'.join(exclude_list), case=False)].copy(
            deep=True)

        # Still some rogue values due to comments added to results files
        include_list = ['2019-nCoV not detected', '2019-nCoV detected', 'Inconclusive Result', 'Invalid Result']

        cr2 = cr1[cr1['Result_Interpretation'].str.contains('|'.join(include_list), case=False)].copy(deep=True)

        # Split filename into month, day, and year columns
        cr2['Month'], cr2['Day'], cr2['Year'], cr2['Hour'], cr2['Minute'], cr2['Seconds'], cr2['Residual'] = \
            cr2['Filename'].str.split('_', 6).str

        cr2['datetime'] = pd.to_datetime(cr2[['Month', 'Day', 'Year']])

        # Sort by ascending datetime in order to keep most recent result only
        cr2 = cr2.sort_values(by=['Month', 'Day', 'Year', 'Hour', 'Minute'], ascending=True)

        # Since all inconclusive results are retested and determined to be pos/neg, keep only the most recent result.
        cr2 = cr2.drop_duplicates(subset="Sample_ID", keep='last')

        cr2['positive'] = None
        cr2.loc[(cr2['Result_Interpretation'] == '2019-nCoV detected'), 'positive'] = 'positive'
        cr2['negative'] = None
        cr2.loc[(cr2['Result_Interpretation'] == '2019-nCoV not detected'), 'negative'] = 'negative'
        cr2['inconclusive'] = None
        cr2.loc[(cr2['Result_Interpretation'] == 'Inconclusive Result'), 'inconclusive'] = 'inconclusive'

        cols = ['positive', 'negative', 'inconclusive']
        cr2['results'] = cr2[cols].apply(lambda x: ''.join(x.dropna()), axis=1)

        week_group = cr2.groupby(cr2['datetime'].dt.week)['results'].value_counts().unstack(1)
        week_df = week_group.add_suffix("_results").reset_index().fillna(0)

        new_cols = ['positive_results', 'negative_results', 'inconclusive_results']
        week_df[new_cols] = week_df[new_cols].applymap(np.int64)

        # Prep outpath and output file name
        timestr = time.strftime('%m_%d_%Y')

        # This portion works for Unix systems - see section below for Windows.
        outname = os.path.split(dir)
        outname1 = outname[0]
        outfilename = outname[1]

        # For Windows-based file paths
        mypath = os.path.abspath(os.path.dirname(dir))
        newpath = os.path.join(mypath, './statistics_and_plots')
        normpath = os.path.normpath(newpath)
        new_base_week = timestr + '_AIHG_2019-nCoVRT-PCR_weekly_results.png'

        # Plotting
        dates = np.arange(len(week_df))
        width = 0.3
        opacity = 0.4

        plt.figure(figsize=(10, 12))

        fig, ax = plt.subplots()

        ax.barh(dates, week_df['negative_results'], width, alpha=opacity, color="blue", label="Negative")
        ax.barh(dates + width, week_df['positive_results'], width, alpha=opacity, color="red", label="Positive")
        ax.barh(dates + (width * 2), week_df['inconclusive_results'], width, alpha=opacity, color="green",
                label="Inconclusive")
        ax.set(yticks=dates + width, yticklabels=week_df['datetime'], ylim=[2 * width - 1, len(week_df)])
        ax.legend()
        ax.set_ylabel("2020 Week Number")
        ax.set_xlabel("Count")
        ax.set_title("AIHG 2019-nCoV RT-PCR Weekly Results")

        for i, v in enumerate(week_df['negative_results']):
            ax.text(v + 4, i, str(v), color="blue", va="center")

        for o, b in enumerate(week_df['positive_results']):
            ax.text(b + 4, o + 0.3, str(b), color="red", va="center")

        for p, n in enumerate(week_df['inconclusive_results']):
            ax.text(n + 4, p + 0.6, str(n), color="green", va="center")

        fig.tight_layout()

        fig.subplots_adjust(right=1.75)

        fig.savefig(normpath + '\\' + new_base_week, dpi=300, bbox_inches="tight")

        # Clear current figure prior to plotting monthly results
        plt.clf()

        month_group = cr2.groupby(cr2['datetime'].dt.month)['results'].value_counts().unstack(1)
        month_df = month_group.add_suffix("_results").reset_index().fillna(0)

        new_cols = ['positive_results', 'negative_results', 'inconclusive_results']
        month_df[new_cols] = month_df[new_cols].applymap(np.int64)

        # Prep outpath and output file name
        timestr = time.strftime('%m_%d_%Y')

        # This portion works for Unix systems - see section below for Windows.
        outname = os.path.split(dir)
        outname1 = outname[0]
        outfilename = outname[1]

        # For Windows-based file paths
        mypath = os.path.abspath(os.path.dirname(dir))
        newpath = os.path.join(mypath, './statistics_and_plots')
        normpath = os.path.normpath(newpath)
        new_base_month = timestr + '_AIHG_2019-nCoVRT-PCR_monthly_results.png'

        # Plotting
        dates = np.arange(len(month_df))
        width = 0.3
        opacity = 0.4

        plt.figure(figsize=(10, 12))

        fig, ax = plt.subplots()

        ax.barh(dates, month_df['negative_results'], width, alpha=opacity, color="blue", label="Negative")
        ax.barh(dates + width, month_df['positive_results'], width, alpha=opacity, color="red", label="Positive")
        ax.barh(dates + (width * 2), month_df['inconclusive_results'], width, alpha=opacity, color="green",
                label="Inconclusive")
        ax.set(yticks=dates + width, yticklabels=month_df['datetime'], ylim=[2 * width - 1, len(month_df)])
        ax.legend()
        ax.set_ylabel("2020 Month Number")
        ax.set_xlabel("Count")
        ax.set_title("AIHG 2019-nCoV RT-PCR Monthly Results")

        for i, v in enumerate(month_df['negative_results']):
            ax.text(v + 4, i, str(v), color="blue", va="center")

        for o, b in enumerate(month_df['positive_results']):
            ax.text(b + 4, o + 0.3, str(b), color="red", va="center")

        for p, n in enumerate(month_df['inconclusive_results']):
            ax.text(n + 4, p + 0.6, str(n), color="green", va="center")

        fig.tight_layout()

        fig.subplots_adjust(right=1.75)

        fig.savefig(normpath + '\\' + new_base_month, dpi=300, bbox_inches="tight")

        messagebox.showinfo("Complete", "Plotting Successful!")

        # TODO: ADD dirstatsresultsmonth - Monthly results
    def dirstatsresultsmonth(self):
        dir = filedialog.askdirectory()
        files_csv = [f for f in os.listdir(dir) if f.endswith('csv')]

        file_list = list()

        for file in files_csv:
            df = pd.read_csv(dir + "\\" + file, sep=",", header=None, skiprows=1)
            df['filename'] = file
            file_list.append(df)

        compiled_results = pd.concat(file_list, axis=0, ignore_index=True)

        compiled_results.dropna(axis=1, how='all', inplace=True)

        compiled_results.columns = ['Sample_ID', 'Result_N1', 'Result_N2', 'Result_RP', 'Result_Interpretation',
                                        'Report', 'Actions', 'Filename']

        compiled_results.dropna(axis=0, subset=['Sample_ID'], inplace=True)

        exclude_list = ['LOW1', 'LOW2', 'LOW3', 'MID1', 'MID2', 'MID3', 'Mod-1_No_90', 'Mod-2_No_90', 'Mod-3_No_90',
                            'Low-1_No_90', 'Low-2_No_90', 'Low-3_No_90', 'Low_1', 'Low_2', 'Low_3', 'Mod_1', 'Mod_2',
                            'Mod_3', 'State1_04022020', 'State2_04022020', 'State3_04022020', 'State4_04022020',
                            'State5_04022020', 'State6_04022020', 'State7_04022020', 'State8_04022020', 'State_1',
                            '032420-7-M', '032420-4-M', '032420-3-M', 'Low-4_QS', 'Low-5_QS', 'Low-6_QS', 'Mod-4_QS',
                            'Mod-5_QS', 'Mod-6_QS', 'Low-1_QS', 'Low-2_QS', 'Low-3_QS', 'Mod-1_QS', 'Mod-2_QS',
                            'Mod-3_QS',
                            'H2O1', 'H2O2', 'H2O3', '0.16_A_Validation', '0.16_B_Validation', '0.8_A_Validation',
                            '0.8_B_Validation', '032420-5-M', '032420-8-M', '_NEG_', '20_A_Validation',
                            '20_B_Validation',
                            '40_state3', '4_A_Validation', '4_B_Validation']

        cr2 = compiled_results[~compiled_results['Sample_ID'].str.contains('|'.join(exclude_list), case=False)].copy(
            deep=True)

        cr2['Month'], cr2['Day'], cr2['Year'], cr2['Residual'] = cr2['Filename'].str.split('_', 3).str

        cr2['datetime'] = pd.to_datetime(cr2[['Month', 'Day', 'Year']])

        cr2 = cr2.sort_values('datetime', ascending=True)

        cr2['positive'] = None
        cr2.loc[(cr2['Result_Interpretation'] == '2019-nCoV detected'), 'positive'] = 'positive'
        cr2['negative'] = None
        cr2.loc[(cr2['Result_Interpretation'] == '2019-nCoV not detected'), 'negative'] = 'negative'
        cr2['inconclusive'] = None
        cr2.loc[(cr2['Result_Interpretation'] == 'Inconclusive Result'), 'inconclusive'] = 'inconclusive'

        cols = ['positive', 'negative', 'inconclusive']
        cr2['results'] = cr2[cols].apply(lambda x: ''.join(x.dropna()), axis=1)




    # Manual antibodyprocess
    def antibodyprocess(self):
        abpath = filedialog.askopenfilename()
        # Read in file - encoding is important
        abdf = pd.read_csv(abpath, sep='\t', encoding='utf-16', skiprows=2, skipfooter=4, engine='python')

        # Simple replacements for spaces included in controls
        abdf = abdf.replace('Neg Ctrl', 'Neg_Ctrl')
        abdf = abdf.replace('Pos Ctrl', 'Pos_Ctrl')

        # Replace empty Sample names with np.nan in order fo forward fill
        abdf = abdf.replace(r'\s', np.nan, regex=True)

        # Forward fill Sample names
        abdf['Sample'] = abdf['Sample'].fillna(method='ffill')

        # Control counts for log
        neg_count = len(abdf[abdf['Sample'].str.contains('Neg_Ctrl')])
        pos_count = len(abdf[abdf['Sample'].str.contains('Pos_Ctrl')])

        # Updated - Drop Sample Names that appear as NaN
        abdf_clean = abdf.dropna(subset=['Sample'])

        # Updated - Filter for samples (exclude controls)
        controls_list = ['Neg_Ctrl', 'Pos_Ctrl']
        samples = abdf_clean[~abdf_clean['Sample'].str.contains('|'.join(controls_list))]\
            .copy(deep=True).sort_values(by=['Sample'])

        # Make a dataframe of mean optical density (OD) values
        meanod_df = abdf.groupby('Sample', as_index=False)['OD'].mean().set_index('Sample').rename(columns={'OD':"mean_OD"})

        # Obtain the absorbance of the positive control
        xPC = meanod_df.loc['Pos_Ctrl', 'mean_OD'].round(5)

        # Calculate the average value of the absorbance of the negative control
        xNC = meanod_df.loc['Neg_Ctrl', 'mean_OD'].round(5)

        # Quality control
        # The average value of the absorbance of the negative control is than 0.25
        # The absorbance of the positive control is NOT less than 0.30
        neg_ctrl_avg_value_threshold = 0.25
        pos_ctrl_value_threshold = 0.30

        rules = [xNC < neg_ctrl_avg_value_threshold,
                 xPC > pos_ctrl_value_threshold]

        try:
            if all(rules):
                positive_cutoff = 1.1 * (xNC + 0.18)
                print("Positive cutoff: ", positive_cutoff.round(5))
                negative_cutoff = 0.9 * (xNC + 0.18)
                print("Negative cutoff: ", negative_cutoff.round(5))
        #   elif xNC > neg_ctrl_avg_value_threshold:
        #       print("WARNING: The average absorbance of negative control exceeds the threshold of 0.25.")
        #   elif xPC < pos_ctrl_value_threshold:
        #       print("WARNING: The absorbance of the positive control is less than the threshold of 0.30.")

                # Make results table
                sampledf = meanod_df.copy(deep=True)

                # Interpretation of results
                sampledf.loc[sampledf['mean_OD'] <= negative_cutoff, "Interpretation"] = "Negative"
                sampledf.loc[sampledf['mean_OD'] >= positive_cutoff, "Interpretation"] = "Positive"
                sampledf.loc[(sampledf['mean_OD'] > negative_cutoff) & (sampledf['mean_OD'] < positive_cutoff),
                             "Interpretation"] = "Borderline"

                sampledf.loc[sampledf['Interpretation'] == "Negative", "Results"] = \
                    "The sample does not contain the new coronavirus (COVID-19) IgG-related antibody"
                sampledf.loc[sampledf['Interpretation'] == "Positive", "Results"] = \
                    "The sample contains novel coronavirus (COVID-19) and IgG-associated antibodies"
                sampledf.loc[sampledf['Interpretation'] == "Borderline", "Results"] = \
                    "Retest the sample in conjunction with other clinical tests"

                # Reset index
                sampledf = sampledf.reset_index()

                # Prepare the outpath for the processed data using a timestamp
                timestr = time.strftime('%m_%d_%Y_%H_%M_%S')

                # This portion works for Unix systems - see section below for Windows.
                outname = os.path.split(abpath)
                outname1 = outname[0]
                outfilename = outname[1]

                # For Windows-based file paths
                mypath = os.path.abspath(os.path.dirname(abpath))
                newpath = os.path.join(mypath, '../../processed')
                normpath = os.path.normpath(newpath)
                new_base = timestr + '_ELISA_results.csv'
                sampledf.to_csv(normpath + '\\' + new_base, sep=",", index=False)

                # For logging
                info_orig = pd.read_csv(abpath, sep='\t', encoding='utf-16', skiprows=2, engine='python')

                # Take only last line (filename information) and reset index
                bottom = info_orig.tail(1).reset_index(drop=True)

                # Obtain run info
                runinfo = (bottom.iloc[0, 0])

                # For Windows-based file paths
                newlogpath = os.path.join(mypath, '../../processed/logs')
                normlogpath = os.path.normpath(newlogpath)
                log_base = timestr + '_covid_ELISA_output.log'
                log_filename = normlogpath + '\\' + log_base

                # Define log file parameters
                logging.basicConfig(filename=log_filename, level=logging.DEBUG,
                                    format='%(asctime)s %(levelname)s %(message)s',
                                    datefmt='%H:%M:%S')
                # Info for log file
                logging.info(' Name of input file: ' + outfilename)
                logging.info('\n')
                logging.info(' Run information: ')
                logging.info('\n' + str(runinfo))
                logging.info('\n')
                logging.info(' Number of positive controls run: ' + str(pos_count))
                logging.info(' Number of negative controls run: ' + str(neg_count))
                logging.info('\n')
                logging.info(' Absorbance of positive control: ' + str(xPC))
                logging.info(' Average absorbance of negative control(s): ' + str(xNC))
                logging.info('\n')
                logging.info('Quality Control: ')
                logging.info(' Absorbance of positive control greater than 0.30? ' + str(xPC > pos_ctrl_value_threshold))
                logging.info(
                    ' Average absorbance of negative control less than 0.25? ' + str(xNC < neg_ctrl_avg_value_threshold))
                logging.info('\n')
                logging.info(' Cutoffs (as determined by absorbance of the negative control): ')
                logging.info(' Positive cutoff: ' + str(positive_cutoff.round(5)))
                logging.info(' Negative cutoff: ' + str(negative_cutoff.round(5)))
                logging.info('\n')
                logging.info(' Number of samples run: ' + str(len(samples['Sample'].unique().tolist())))
                logging.info('Samples run: ')
                logging.info(str(samples['Sample'].unique()))

            elif xNC > neg_ctrl_avg_value_threshold:
                raise ValueError("ERROR: The average absorbance of negative control exceeds the threshold of 0.25.")

            elif xPC < pos_ctrl_value_threshold:
                raise ValueError("ERROR: The absorbance of the positive control is less than the threshold of 0.30.")

        except Exception as e:
                s = getattr(e, "Could not interpret results because one or more controls are out of bounds.", repr(e))
                # print(s)
                messagebox.showinfo("ERROR", s)

        messagebox.showinfo("Complete", "ELISA Data Processing Complete!")


my_gui = AIHGdataprocessor(root)
root.update()
root.mainloop()
