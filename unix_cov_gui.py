"""
By: Jeff Beck and Casey Finnicum
Date of inception: March 16, 2020
Program for determining results of 2019-nCoV testing at the Avera Institute for Human Genetics.
Ingests files from RT-qPCR assay and creates summarized results for upload.
Reference: CDC-006-00019, Revision: 02
"""

from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
import os
import ntpath
import time
import logging
from PIL import ImageTk, Image

root = Tk()
root.configure(bg='white')

img = ImageTk.PhotoImage(Image.open("./misc/aihg.gif"))
panel = Label(root, image=img)
panel.pack(side="bottom", fill="both", expand="yes")


class COV:
    def __init__(self, master):
        master.minsize(width=200, height=100)
        self.master = master
        master.title("COVID-19 Data Processor")

        self.convert_button = Button(master, text="Select input file",
                                     command=self.dataprocess, width=13)
        self.convert_button.pack(pady=10)

        # self.convert_button = Button(master, text="COVID-19 RT-qPCR Data Processing",
        #                              command=self.dataprocess, width=45)
        # self.convert_button.grid(row=1, column=1)

    # def secondProcess(self):
        # print('out')

    def dataprocess(self):
        # Ingest input file
        # ask the user for an input read in the file selected by the user
        path = filedialog.askopenfilename()
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
        df['HSC_N1'] = None  # initial value
        df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N1') & (df['CT'].isnull()), 'HSC_N1'] = 'passed'
        df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N1') & (df['CT'].notnull()), 'HSC_N1'] = 'failed'
        df['HSC_N2'] = None  # initial value
        df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N2') & (df['CT'].isnull()), 'HSC_N2'] = 'passed'
        df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'N2') & (df['CT'].notnull()), 'HSC_N2'] = 'failed'
        df['HSC_RP'] = None  # initial value
        df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'RP') & (df['CT'] <= ct_value), 'HSC_RP'] = 'passed'
        df.loc[(df['Sample Name'] == 'HSC') & (df['Target Name'] == 'RP') & (df['CT'] > ct_value), 'HSC_RP'] = 'failed'

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
        df['Extraction_control'] = None
        df.loc[(df['Sample Name'] == 'HSC') & (df['HSC_N1'] == 'passed')
               | (df['Sample Name'] == 'HSC') & (df['HSC_N2'] == 'passed')
               | (df['Sample Name'] == 'HSC') & (df['HSC_RP'] == 'passed'), 'Extraction_control'] = 'passed'
        df.loc[(df['Sample Name'] == 'HSC') & (df['HSC_N1'] == 'failed')
               | (df['Sample Name'] == 'HSC') & (df['HSC_N2'] == 'failed')
               | (df['Sample Name'] == 'HSC') & (df['HSC_RP'] == 'failed'), 'Extraction_control'] = 'failed'

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
        controls_filtered = df.loc[
            (df['Sample Name'] == 'NTC') | (df['Sample Name'] == 'HSC') | (df['Sample Name'] == 'nCoVPC')]
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
        # Results for N1 assay
        df.loc[(df['Target Name'] == 'N1') & (df['CT'] > ct_value) | (df['Target Name'] == 'N1') & (df['CT'].isnull()),
               'result'] = 'negative'
        df.loc[(df['Target Name'] == 'N1') & (df['CT'] < ct_value) & (df['NOAMP'] == "Y"), 'result'] = 'negative'
        df.loc[(df['Target Name'] == 'N1') & (df['CT'] < ct_value) & (df['NOAMP'] == "N"), 'result'] = 'positive'
        # Results for N2 assay
        df.loc[(df['Target Name'] == 'N2') & (df['CT'] > ct_value) | (df['Target Name'] == 'N2') & (df['CT'].isnull()),
               'result'] = 'negative'
        df.loc[(df['Target Name'] == 'N2') & (df['CT'] < ct_value) & (df['NOAMP'] == "Y"), 'result'] = 'negative'
        df.loc[(df['Target Name'] == 'N2') & (df['CT'] < ct_value) & (df['NOAMP'] == "N"), 'result'] = 'positive'
        # Results for RP assay
        df.loc[(df['Target Name'] == 'RP') & (df['CT'] > ct_value) | (df['Target Name'] == 'RP') & (df['CT'].isnull()),
               'result'] = 'negative'
        df.loc[(df['Target Name'] == 'RP') & (df['CT'] < ct_value) & (df['NOAMP'] == "Y"), 'result'] = 'negative'
        df.loc[(df['Target Name'] == 'RP') & (df['CT'] < ct_value) & (df['NOAMP'] == "N"), 'result'] = 'positive'

        # Filter for samples (exclude controls)
        sf = df[df['Sample Name'].apply(lambda x: x not in ['NTC', 'HSC', 'nCoVPC'])].copy(deep=True).sort_values(
            by=['Sample Name'])
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
        outname = os.path.split(path)
        outname1 = outname[0]
        outfilename = outname[1]
        new_base = timestr + '_covid_results.csv'
        # original
        outpath = outname1 + '/' + new_base
        sf.to_csv(outpath, sep=",", index=False)

        # new
        # print(path)
        # path_parts = os.path.split(os.path.sep)
        # print(path_parts)
        # temp_path_parts = path_parts[:-3]
        # print(temp_path_parts)
        # new_path = os.path.join('/', *temp_path_parts)
        # print(new_path)
        # processedpath = '/processed/'
        # sf.to_csv(new_path + processedpath + new_base, sep=",", index=False)

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

        # Log file
        # Prepare path for the log file
        # logpath = '/processed/logs/'
        # original
        log_base = timestr + '_covid_output.log'
        log_filename = outname1 + '/' + log_base
        # new
        # log_filename = new_path + logpath + log_base

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


my_gui = COV(root)
root.update()
root.mainloop()
