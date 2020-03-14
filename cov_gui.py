from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
import os
import json
# import boto3
import time

root = Tk()
root.configure(bg='gray')


# def datauploads3():
#     filepath = filedialog.askopenfilename()
#     name = os.path.basename(filepath)
#     s3 = boto3.client('s3')
#     with open(filepath, "rb") as f:
#         s3.upload_fileobj(f, 'cfinnpgx', name)
#     messagebox.showinfo("Complete", "Data has been uploaded to S3 bucket")


class COV:

    def __init__(self, master):
        master.minsize(width=400, height=200)
        self.master = master
        master.title("COVID-19 Data Processor")

        # self.convert_button2 = Button(master, text="Data Upload S3", command=secondProcess, width=25)
        # self.convert_button2.grid(row=1, column=1)

        self.convert_button3 = Button(master, text="COVID-19 RT-qPCR Data Processing", command=self.labProcess, width=25)
        self.convert_button3.grid(row=2, column=1)

    def secondProcess(self):
        # TODO: Need to correct phenotype for 2d6 based on CNV
        print('out')

    def dataprocess(self):
        # ask the user for an input read in the file selected by the user
        path = filedialog.askopenfilename()
        df = pd.read_csv(path)

        # prepare the outpath for the processed data using a timestamp
        timestr = time.strftime('%m_%d_%Y')
        outname = os.path.split(path)
        outname1 = outname[0]
        new_base = timestr + '_covid.csv'
        outPath = outname1 + '/' + new_base

        # prepare path for the log file
        logName = timestr + '_output_log.txt'
        outPathLog = outname1 + '/' + logName


        # use this for
        #     with open(outPath, 'w') as file:
        #         for key, value in out_dict.items():
        #             file.write(key + "\t" + value + "\n")
        #
        # log_out = ['There where ' + str(sum([invenioLen, averaLen])) + ' patients uploaded' + '\n',
        #            'There were ' + str(averaLen) +
        #            ' Avera patients in the uploaded data, the data for these patients can be found in ' + new_baseAvera + '\n',
        #            'There were ' + str(invenioLen) +
        #            ' Invenio patients in the uploaded data, the data for these patients can be found in ' + new_baseInvenio]
        #
        # with open(outPathLog, 'w') as file:
        #     for i in log_out:
        #         file.write(i)

        # TODO: write out the log file that explains the number of patients and if there were any missing data

        messagebox.showinfo("Complete", "Data Processing Complete!")

    # TODO: possibly implement a finer timestamp so reruns would not  be over wrote. Could also check for same name


my_gui = COV(root)
root.mainloop()
