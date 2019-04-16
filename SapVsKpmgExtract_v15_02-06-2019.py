#!/usr/bin/env python
# coding: utf-8

# # Import

# In[9]:


import pandas as pd
import numpy as np
import xlsxwriter
import os
import glob
import dask.dataframe as dd
import zipfile
import shutil
import win32com.client
import PySimpleGUI as sg
import datetime
import time
import warnings
import networkx as nx
import cx_Oracle
warnings.filterwarnings('ignore')
pd.set_option('display.float_format', lambda x: '%.2f' % x)


# # SAP Pull

# In[10]:


def sap_pull(num_retry=3):
    global filename
    
    complete = False
    filename = datetime.datetime.now().strftime('%m-%d-%Y')
    filename = r'ZSE16H_{}.xlsx'.format(str(filename))
    
    window.Read(timeout=0)
    window.Element('task').Update("Pulling Data From FAGLEFLEXT....")
    window.FindElement('progbar').UpdateBar(10)
    window.Read(timeout=0)
    time.sleep(.1)
    
    if any(File.endswith(".xlsx") and File.startswith("ZSE16H") for File in os.listdir(filepath)):
        use = sg.PopupYesNo('ZSE16H extract file is already available in this folder. Do you want to continue using this file?')
        if use == 'Yes':
            complete = True
        else:
            complete = False
    
    
    
    for attempt_no in range(num_retry):
        while complete is False:  
            try:
                SapGuiAuto = win32com.client.GetObject("SAPGui")
                if not type(SapGuiAuto) == win32com.client.CDispatch:
                    return

                application = SapGuiAuto.GetScriptingEngine
                if not type(application) == win32com.client.CDispatch:
                    SapGuiAuto = None
                    return

                connection = application.Children(0)
                if not type(connection) == win32com.client.CDispatch:
                    application = None
                    SapGuiAuto = None
                    return

                session = connection.Children(0)
                if not type(session) == win32com.client.CDispatch:
                    connection = None
                    application = None
                    SapGuiAuto = None
                    return

                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16h"
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/ctxtGD-TAB").text = "FAGLFLEXT"
                session.findById("wnd[0]/usr/ctxtGD-TAB").setFocus()
                session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 9
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]"
                ).text = "2018"
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]"
                ).setFocus()
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]"
                ).caretPosition = 4
                session.findById("wnd[0]/tbar[1]/btn[18]").press()
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,1]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-GROUP_BY[8,1]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,24]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-GROUP_BY[8,24]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,18]"
                ).text = "0l"
                # session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,24]").setFocus()
                # session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,24]").press()
                # session.findById("wnd[1]/tbar[0]/btn[24]").press()
                # session.findById("wnd[1]/tbar[0]/btn[8]").press()
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,22]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-GROUP_BY[8,22]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-GROUP_BY[8,22]"
                ).setFocus()
                session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC"
                                 ).verticalScrollbar.position = 65
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,11]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,11]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,12]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,13]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,14]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,15]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,16]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,17]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,18]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,19]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,20]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,21]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,22]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,23]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,24]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,25]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,26]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[6,27]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,12]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,13]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,14]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,15]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,16]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,17]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,18]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,19]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,20]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,21]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,22]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,23]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,24]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,25]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,26]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,27]"
                ).selected = True
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-SUM_UP[7,27]"
                ).setFocus()
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell"
                                 ).pressToolbarContextButton("&MB_EXPORT")
                session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell"
                                 ).selectContextMenuItem("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = filepath
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
                session.findById("wnd[1]/tbar[0]/btn[0]").press()


                session = None
                connection = None
                application = None
                SapGuiAuto = None

                
                complete = True
                
            
            except :

                sg.Popup("Failed","Unable to Retrieve Data from SAP. Make sure SAP is opened to the Production Database(R3P)")
                break
    

    if complete is True:
        window.Read(timeout=0)
        window.FindElement('task').Update("File Saved")
        window.FindElement('progbar').UpdateBar(30)
        window.Read(timeout=0)
        time.sleep(.2)
    else:
        
#         #filename = input(
#             "Extract Failed. Please Manually Extract SAP Data, Place in Folder, and provide file name"
#         )
        filename = sg.PopupGetFile("Extract Failed. Please Manually Extract SAP Data, Place in Folder, and provide file name",'SAP Vs Kpmg Reconciliation')

    return 


# # Kpmg Aggregation

# In[11]:


def kpmg_sum(num_retry=4):
    
    if any(File.endswith(".txt") and File.startswith("kpmg") for File in os.listdir(filepath)):
        use = sg.PopupYesNo('"kpmg_calc.txt" file is already available in this folder. Do you want to continue using this file?')
        if use == 'Yes':
            complete = True
        else:
            complete = False
    
    if complete == False: 
        # Try loop incase the first attempt fails. Look for folder that contains several zipped files.
        for attempt_no in range(num_retry):

            # Search through user path for zipped files and prints grabbed files
            kpmg_files = glob.glob(os.path.join(filepath, "FAGLFLEXT*.zip"))
            foldername = datetime.datetime.now().strftime('%m-%d-%Y')
            os.mkdir(os.path.join(filepath,foldername))

            for f in kpmg_files:
                shutil.move(f,os.path.join(filepath,foldername))

            kpmg_files = glob.glob(
                os.path.join(filepath, "*\\*FAGLFLEXT*.zip"), recursive=True)

            if kpmg_files == []:
                sg.PopupOk('Couldnt find any zip files! Please Check file path!')
                found = False
            else:
                found = True
                break

        #Update Comments in Gui Window 
        window.Read(timeout=0)
        window.FindElement('progbar').UpdateBar(10)
        window.FindElement('task').Update("Extracting From Zipped Files")
        window.Read(timeout=0)
        time.sleep(0.2)

        while found is True:

            # Loops through zipped folder to extract txt file to a temp folder and prints out file from zipped folder
            tempf_path = os.path.join(filepath, "temp")
            for x in kpmg_files:
                zip_ref = zipfile.ZipFile(x)
                print("Files Extracted", zip_ref.namelist())
                if not os.path.exists(tempf_path):
                    os.makedirs(tempf_path)
                zip_ref.extractall(tempf_path)
                zip_ref.close()


            #Update Comments in Gui Window  
            window.Read(timeout=0)
            window.FindElement('progbar').UpdateBar(40)
            window.FindElement('task').Update("Files Extracted")
            window.Read(timeout=0)
            time.sleep(0.2)


            #Reads in each txt file and combines them to a dask dataframe
            cl = ['#RBUKRS#','#RACCT#', '#KSLVT#', '#KSL01#', '#KSL02#', '#KSL03#', '#KSL04#', '#KSL05#',
                  '#KSL06#', '#KSL07#', '#KSL08#', '#KSL09#','#KSL10#', '#KSL11#', '#KSL12#', '#KSL13#','#KSL14#', 
                  '#KSL15#', '#KSL16#', '#RYEAR#']

            kpmgdf = dd.read_csv(os.path.join(tempf_path,"*.txt"), sep='|', header=0,
                                 dtype= {'#ACCT#':'object', '#RBURKS#':'object'})

            kpmgdf = kpmgdf[cl]



            #Strip out '#' signs
            kpmgdf = kpmgdf.rename(columns=lambda x: x.strip('#'))
            kpmgdf = kpmgdf.apply(lambda x: x.str.strip('#'), axis = 1)
            kpmgdf = kpmgdf[kpmgdf.RYEAR == '2018']

            # Update Datatype
            kpmgdf['KSLVT'] = kpmgdf['KSLVT'].astype(np.float64)
            kpmgdf['KSL01'] = kpmgdf['KSL01'].astype(np.float64)
            kpmgdf['KSL02'] = kpmgdf['KSL02'].astype(np.float64)
            kpmgdf['KSL03'] = kpmgdf['KSL03'].astype(np.float64)
            kpmgdf['KSL04'] = kpmgdf['KSL04'].astype(np.float64)
            kpmgdf['KSL05'] = kpmgdf['KSL05'].astype(np.float64)
            kpmgdf['KSL06'] = kpmgdf['KSL06'].astype(np.float64)
            kpmgdf['KSL07'] = kpmgdf['KSL07'].astype(np.float64)
            kpmgdf['KSL08'] = kpmgdf['KSL08'].astype(np.float64)
            kpmgdf['KSL09'] = kpmgdf['KSL09'].astype(np.float64)
            kpmgdf['KSL10'] = kpmgdf['KSL10'].astype(np.float64)
            kpmgdf['KSL11'] = kpmgdf['KSL11'].astype(np.float64)
            kpmgdf['KSL12'] = kpmgdf['KSL12'].astype(np.float64)
            kpmgdf['KSL13'] = kpmgdf['KSL13'].astype(np.float64)
            kpmgdf['KSL14'] = kpmgdf['KSL14'].astype(np.float64)
            kpmgdf['KSL15'] = kpmgdf['KSL15'].astype(np.float64)
            kpmgdf['KSL16'] = kpmgdf['KSL16'].astype(np.float64)
            kpmgdf['RYEAR'] = kpmgdf['RYEAR'].astype(np.int64)
            kpmgdf = kpmgdf.drop(['RYEAR'], axis=1)


            #Update Comments in Gui Window  
            window.FindElement('progbar').UpdateBar(40)
            window.FindElement('task').Update("Summerizing Files...")
            print("This part can up to an 1HR. Please Remain Patient.")
            window.Read(timeout=0)
            time.sleep(0.2)


    #        Group by "RBURKS" and "RACCT" then generate to CSV file

            kpmg_calc = kpmgdf.groupby(['RBUKRS', 'RACCT']).sum().compute()

    #         # Group by "RBURKS" and "RACCT" then generate to CSV file
    #         kpmg_calc = kpmgdf.groupby(['RBUKRS', 'RACCT']).sum().compute()

            kpmg_calc.to_csv(os.path.join(filepath, r"kpmg_calc.txt"), float_format='%f')

        #Delete Temp file containing extracted text file
        shutil.rmtree(tempf_path)
        window.FindElement('task').Update("KPMG file Summerized")
        window.FindElement('progbar').UpdateBar(50)
        print("Kpmg_calc Saved to",filepath)
        window.Read(timeout=0)
        time.sleep(0.1)
        found = False

        


# # Join, Compare and To Excel

# In[12]:


def to_excel(NodesAcct):
  
    # Read in both excel files
    df2 = pd.read_csv(
        os.path.join(filepath, r"kpmg_calc.txt"),
        dtype={
            'RACCT': 'object',
            'RBUKRS': 'object'
        })  #kpmg summary
    
    df1 = pd.read_excel(
        glob.glob(os.path.join(filepath, "ZSE16H*.xlsx"))[0],
        'Sheet1',
        dtype={
            'Company Code': 'object',
            'Account Number': 'object'
        })  #ZSE16H file 
    
    window.Read(0)
    window.FindElement('task').Update("Merging Data...")
    window.FindElement('progbar').UpdateBar(70)
    window.Read(timeout=0)
    time.sleep(0.2)
    
    # Clean df2 for Join
    df2 = df2.sort_values(['RACCT', 'RBUKRS'])
    df2['RBUKRS'] = df2['RBUKRS'].str.pad(4, 'left', '0')
    df2['KPMG_Total'] = df2.sum(axis=1)
    df2['  '] = ' '
    
    df2 = df2.rename(columns={'RBUKRS': 'Company Code', 'RACCT': 'Account Number'})
    df2['Account Number'] = df2['Account Number'].str.pad(10, 'left', '0')
    df2 = df2.set_index(['Company Code', 'Account Number'])
    
    # Clean df1 for join
    df1 = df1.drop(['Fiscal Year', 'Number of Entries'], axis=1)
    df1.reset_index(inplace=True)   
    
    df1['Account Number'] = df1['Account Number'].str.pad(10, 'left', '0')


    df1 = df1.drop('index', 1)
    df1 = df1.rename(columns={'RBUKRS': 'Company Code', 'RACCT': 'Account Number'})

    df1 = df1.set_index(['Company Code', 'Account Number'])

    df1['SAP_Total'] = df1.sum(axis=1)
    df1[''] = ''
    

    # Joins Kpmg and Sap tables
    Combdf = df2.join(df1, on=['Company Code', 'Account Number'], sort=True, how='outer')
    #Combdf = Combdf.fillna(0.00)
    Combdf['ESS Extracted Manually VS KPMG Extracted '] = Combdf[
        'KPMG_Total'] - Combdf['SAP_Total']
    Combdf[['', '  ']] = ''
    Combdf2 = Combdf.drop(['', '  '],axis=1)
    
    
    # Nodes to Accounts Mapping:    
    recon = Combdf2
    nodes = NodesAcct
    
    recon = recon.reset_index()
    recon = recon.rename(columns={'level_0': 'Company Code', 'level_1': 'Account Number'})
    recon = recon[['Account Number','Company Code','KPMG_Total','SAP_Total','ESS Extracted Manually VS KPMG Extracted ']]

    df = nodes.merge(recon,on='Account Number',how='right')
    df = df.round({'KPMG_Total': 2, 'SAP_Total': 2, 'ESS Extracted Manually VS KPMG Extracted' : 2})
    
    
    
    

    #create Ranges for cell formatting
    range2 = 'C3:S{}'.format(len(Combdf))
    range3 = 'V3:AM{}'.format(len(Combdf))
    range4 = 'AO3:AO{}'.format(len(Combdf))
    range5 = 'C2:J2'
    range6 = 'L2:S2'
    range7 = 'T3:T{}'.format(len(Combdf))
    range8 = 'AM3:AM{}'.format(len(Combdf))
    range9 = 'K1:K{}'.format(len(Combdf))
    range10 = 'T1:T{}'.format(len(Combdf))
    
    
    
    window.FindElement('task').Update("Writing Data to Excel File")
    window.FindElement('progbar').UpdateBar(80)
    window.Read(timeout=0)
    time.sleep(0.2)
    
    
    #create Excel Object
    writer = pd.ExcelWriter(
        os.path.join(filepath,
                     'KPMG_File_Reconciliation{}.xlsx'.format(datetime.datetime.now().strftime('%m-%d-%Y'))),
        engine='xlsxwriter')

    #create Excel Sheet
    Combdf.to_excel(writer, sheet_name='Detail', startrow=1)
    df.to_excel(writer,sheet_name='PivotSourceData', startrow=0)
    
    worksheet = writer.sheets['Detail']
    workbook = writer.book

    #Define Formats
    red_format = workbook.add_format({
        'bg_color': '#FFC7CE',
        'font_color': '#9C0006',
        'border': 1
    })  #  ess VS kpmg column values

    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',  # column headers
        'fg_color': '#c2bcbc',
        'border': 1
    })

    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',  #Extract headers
        'valign': 'vcenter',
        'fg_color': '#c2bcbc',
        'size': 30
    })
    merge_format2 = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',  #index format(Company codes and Account Number)
        'valign': 'vcenter',
        'fg_color': '#c2bcbc',
        'size': 15
    })

    money = workbook.add_format({
        'num_format': '$#,##0.00',
        'border': 1
    })  #values format

    totals = workbook.add_format({
        'fg_color': '#AED6F1',
        'num_format': '$#,##0.00',
        'border': 1
    })  # Totals Column
    empty = workbook.add_format({
        'fg_color': '#F6DDCC',
        'border': 1
    })  # format for empty rows
    gtotals = workbook.add_format({
        'fg_color': '#F9E79F',
        'num_format': '$#,##0.00',
        'border': 1
    })  # Difference column
    format1 = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

    worksheet.set_row(0, 30)

    # formating Headers
    worksheet.set_column(range5, 18, header_format)
    worksheet.set_column(range6, 18, header_format)
    
    
    #write data to excel
    for col_num, value in enumerate(Combdf.columns.values):
        worksheet.write(1, col_num + 2, value, header_format)
        
    

    #formating values
    worksheet.set_column(range2, 20, money)
    worksheet.set_column(range3, 20, money)
    worksheet.set_column(range4, 20, money)
    #Total Columns Background
    worksheet.set_column(range7, 20, totals)
    worksheet.set_column(range8, 20, totals)
    worksheet.set_column('U:U', 5, empty)
    worksheet.set_column('U:U', 5, empty)
    worksheet.set_column(range4, 45, gtotals)
    worksheet.set_column('A:A', 19)
    worksheet.set_column('B:B', 21.14)

    #Merged Cells formating(Sheet1)
    worksheet.merge_range('C1:T1', 'KPMG Extracted File - Table FAGLFLEXT',
                          merge_format)
    worksheet.merge_range('V1:AM1', 'Summary of Table FAGLFLEXT', merge_format)
    worksheet.merge_range('A1:A2', 'Company Code', merge_format2)
    worksheet.merge_range('B1:B2', 'Account Number', merge_format2)
    worksheet.merge_range('AO1:AO2', 'ESS Extracted Manually VS KPMG Extracted ',
                          header_format)
    worksheet.merge_range('U1:U20000', '', empty)
    worksheet.merge_range('AN1:AN20000', '', empty)
    
    
    
    # write formulas to Totals columns in excel sheet:
    for row_num in range(2, len(Combdf)):#sum of KSL1 - KSL16
        cell = xlsxwriter.utility.xl_rowcol_to_cell(row_num, 19)
        start = '$C{}'.format(row_num + 1)
        stop = '$S{}'.format(row_num + 1)
        row = start + ':' + stop
        formula = '=Round(SUM({}),2)'.format(row)
        worksheet.write_formula(cell, formula)

    for row_num in range(2, len(Combdf)):#sum of 'total trans. in GC' - 'total trans. in GC 15'
        cell = xlsxwriter.utility.xl_rowcol_to_cell(row_num, 38)
        start = '$V{}'.format(row_num + 1)
        stop = '$AL{}'.format(row_num + 1)
        row = start + ':' + stop
        formula = '=Round(SUM({}),2)'.format(row)
        worksheet.write_formula(cell, formula)

    for row_num in range(2, len(Combdf)+2):#diff column
        cell = xlsxwriter.utility.xl_rowcol_to_cell(row_num, 40)
        start = '$T{}'.format(row_num + 1)
        stop = '$AM{}'.format(row_num + 1)
        row = start + '-' + stop
        formula = '={}'.format(row)
        worksheet.write_formula(cell, formula)

    # Highlight rows where sums do not match:
    number_rows = len(Combdf.index) + 10

    worksheet.conditional_format(
        "$C$3:$AM$%d" % (number_rows), {
            "type": "formula",
            "criteria": '=INDIRECT("AO"&ROW())<>0.00',
            "format": format1
        })

 
    
    
    writer.save()
    
    window.FindElement('task').Update("Creating Pivot Table")
    window.FindElement('progbar').UpdateBar(90)
    print('KPMG Extracted FAGLFLEXT File Reconciliation_testv2.xlsx saved at',filepath)
    window.Read(timeout=0)
    time.sleep(1)
    window.Close()


# # Get Path

# In[13]:


def GetPath():
    global filepath
    global filename
    filepath =  sg.PopupGetFolder('Please enter a folder name','SAP Vs Kpmg Reconciliation')
    
    filepath = r"{}".format(filepath)
    sg.Popup('Alert', 'All files will be saved to this Location:', filepath)


# # Add Pivot

# In[14]:



def addPivot_Table():

    win32c = win32com.client.constants

    #Start Excel Application 
    app = win32com.client.gencache.EnsureDispatch('Excel.Application')
    app.Visible = True

    #open exsisting Wb
    wb = app.Workbooks.Open(os.path.join(filepath,
                     'KPMG_File_Reconciliation{}.xlsx'.format(datetime.datetime.now().strftime('%m-%d-%Y'))))


    # Specify Source Data Ranges
    pivotrng = wb.Sheets('PivotSourceData').Range("$B$1:$U$16600")

    # Specify Name:
    PivotTableName = 'ReportPivotTable'

    Sheet1 = wb.Worksheets("Detail")
    sheet2 = wb.Worksheets.Add(After=Sheet1)

    s4 = wb.Worksheets("Sheet1")

    # Output Cells
    cl3 = s4.Cells(4,1)
    cl4 = s4.Cells(4,2)

    # Specify Pivot Table Output
    PivotTargetRange = s4.Cells(4,1)


    #Create Pivot Table:
    PivotCache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=pivotrng, Version=win32c.xlPivotTableVersion14)

    PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)

    #Specify Pivot Table Rows 
    PivotTable.PivotFields('0').Orientation = win32c.xlRowField 
    PivotTable.PivotFields('0').Position = 1
    PivotTable.PivotFields('1').Orientation = win32c.xlRowField
    PivotTable.PivotFields('1').Position = 2
    PivotTable.PivotFields('2').Orientation = win32c.xlRowField
    PivotTable.PivotFields('2').Position = 3
    PivotTable.PivotFields('3').Orientation = win32c.xlRowField
    PivotTable.PivotFields('3').Position = 4
    # PivotTable.PivotFields('Account Number').Orientation = win32c.xlRowField
    # PivotTable.PivotFields('Account Number').Position = 5

    #Specify Pivot Table Values
    DataField1 = PivotTable.AddDataField(PivotTable.PivotFields('KPMG_Total'), "Sum Of KPMG_Total", win32c.xlSum)
    DataField1.NumberFormat = "$#,##0.00_);($#,##0.00)"
    DataField2 = PivotTable.AddDataField(PivotTable.PivotFields('SAP_Total'), "Sum Of SAP_Total", win32c.xlSum)
    DataField2.NumberFormat = "$#,##0.00_);($#,##0.00)"
    DataField3 = PivotTable.AddDataField(PivotTable.PivotFields('ESS Extracted Manually VS KPMG Extracted '),'ESS Extracted Manually VS KPMG Extracted', win32c.xlSum)
    DataField3.NumberFormat = "$#,##0.00_);($#,##0.00)"
    
    wb.Worksheets("Sheet1").Name = "BalanceSheet"
    

    wb.Save()


# # Node Pull

# In[15]:


def GetNodes():
    
    window.FindElement('task').Update("Pulling Nodes Data")
    window.FindElement('progbar').UpdateBar(55)
    window.Read(timeout=0)
    time.sleep(0.2)
    
    
    start_time = time.clock()

    con = cx_Oracle.connect('ato_read/restricted1#@sapatodb:1521/ATO')

    cur = con.cursor()

    query1 = """
            Select SAPR3P.SETNODE.* From SAPR3P.SETNODE
            """
    cur.execute(query1)

    SetNodeAll = pd.read_sql(query1, con)

    cur.close()
    con.close()
    
    df = SetNodeAll[['SETNAME', 'SUBSETNAME']]

    df.columns = ['parent', 'child']

    g2 = nx.from_pandas_edgelist(df, 'parent', 'child', create_using = nx.DiGraph())

    def find_all_paths(graph, start, path=[]):
        path = path + [start]
        yield path
        if start not in graph:
            return
        for node in graph[start]:
            if node not in path:
                yield from find_all_paths(graph, node, path)

    def f(x):
        if x.last_valid_index() is None:
            return np.nan
        else:
            return x[x.last_valid_index()]


    solution = find_all_paths(g2, 'ZF_ALLACCOUNTS')

    final = pd.DataFrame(solution)
    #print(graph)
    final['leaflist'] = final.apply(f, axis=1)

    con = cx_Oracle.connect('ato_read/restricted1#@sapatodb:1521/ATO')

    cur = con.cursor()

    query1 = "Select SAPR3P.SETLEAF.SETNAME,SAPR3P.SETLEAF.VALFROM From SAPR3P.SETLEAF WHERE SETNAME IN {}".format(tuple(final['leaflist']))

    cur.execute(query1)

    Lf = pd.read_sql(query1, con)
    Lf.columns = ['leaflist','Account Number']

    cur.close()
    con.close()


    NodeAcct = final.merge(Lf, on='leaflist').sort_values('Account Number')
    
    NodeAcct['Account Number'] = NodeAcct['Account Number'].str.pad(10, 'left', '0')
    
    window.FindElement('task').Update("Node Data Prepped")
    window.FindElement('progbar').UpdateBar(55)
    window.Read(timeout=0)
    time.sleep(0.2)

    return NodeAcct


# # Main

# In[16]:


def Main():
    cont = 'Yes'
    GetPath()
    sap_pull()
    cont = sg.PopupYesNo('SAP Vs Kpmg Reconciliation',"Click ""Yes"" to continue processing files.", "Click No to end program")
    if cont == 'Yes':
        kpmg_sum()
        NodesAcct = GetNodes()
        to_excel(NodesAcct)
        addPivot_Table()


# # Script

# In[17]:


if __name__ == "__main__":
    
    # create the Window
    layout = [[sg.Text('')],
          [sg.ProgressBar(100, orientation='h', size=(500, 20), key='progbar')],
          [sg.Text('', size=(100, 2), justification='center', key='task')],
          [sg.Output(size=(80, 50))]]
          

    window = sg.Window('SAP Vs Kpmg Reconciliation', size=(500, 500)).Layout(layout)
    
    
    #Run Program
    pd.set_option('display.float_format', lambda x: '%.2f' % x)
    Main()
    sg.Popup('SAP Vs Kpmg Reconciliation', 'Process Is Complete! All files will be saved to this Location:', filepath)
    window.Close()


# In[ ]:




