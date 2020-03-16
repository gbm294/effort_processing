# -*- coding: utf-8 -*-
"""
Created on Fri Sep  6 14:30:08 2019

Make a csv file with all the report spec docs so the IRB and PI numbers
can be scraped from those docs.
@author: gbm294
"""



import os#,sys
import re
from datetime import datetime
#from shutil import copyfile
import pandas as pd
from datetime import datetime

import tkinter
from tkinter import filedialog #GUI


##############################################################################
## Get list of files in directory and all subdirectories
def listPaths(indir):
    fList = []
    for path, subdirs, files in os.walk(indir):
        for name in files:
            if name[-5:] == '.docx':
                fList.append(os.path.join(path, name))
    return fList
############################################################################## 
############################################################################## 
def get_dir(what_file,parent_dir):
    #get input file
    root = tkinter.Tk()  #access windowing system
    user_dir = filedialog.askdirectory(initialdir = parent_dir,title=what_file)
    #user_dir = os.path.dirname(effort_file[0])
    #os.chdir(user_dir)  #specify directory
    #user_dir_name = os.path.basename(os.path.dirname(effort_file[0]))
    #file_path = effort_file[0]
    root.destroy() #kill Tkinter
    
    if len(user_dir) <1 :
        print("you probably didn't choose any files")
        root.destroy() #kill Tkinter
        sys.exit()
    
    return user_dir
############################################################################## 

def search_files():
    #Folder to start searching in
    start_main_Time = datetime.now()

    
    inFolders = ['U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/Requests_DATA/'
                 ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/Requests_COUNTS/'
                 ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/CRDS_Partners_Requests_DATA']
    
    
    SN_ritm_file_dir = get_dir('Get Manual_effort_calcs_gm todays date folder' ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Manual_effort_calcs_gm')    
    #out_file = 'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Effort_tracking_REDCap/2019_12_17/QD_path_list.csv'

    os.chdir(SN_ritm_file_dir)
    working_dir = '%s/02_QD_Path_List_Output' % os.getcwd()
    os.chdir(working_dir)
    out_file = 'QD_path_list.csv'
    out_questionable_file = 'QD_questionable_path_list.csv'


    #create list of files in chosen directory #crawls through subdirectories
    fileList = []
    for f in inFolders:
        l = listPaths(f)
        fileList.extend(l)

    ##output RITMS
    rs_list = []
    questionable_list = []
    #Loop through files in directory
    for f in fileList:
        if f[-5:] == '.docx':  #only looking through word files, so will miss text files
            inFile = f.replace('\\','\\\\')  #for windows file paths need to modify backslashes
            
            (inFilePath, inFileName) = os.path.split(inFile)
            
            match_name = re.search(r'.*report.*spec.*', inFileName, re.IGNORECASE)
            match_name2 = re.search(r'.*query.*design.*', inFileName, re.IGNORECASE)
            match_archive = re.search(r'.*archive.*', inFilePath, re.IGNORECASE)
            
            if (match_name or match_name2) and not match_archive:
                s = inFile
                s = s.replace('\\', '/').replace('//', '/')
                mod_time = None
                mod_time = os.path.getmtime(s)
                mod_time = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
                
                
                ritm = ''
                match_ritm = re.search(r'RITM\d\d\d\d\d\d\d', inFileName)
                if match_ritm:
                    #print(match_ritm.group())
                    ritm = match_ritm.group()
                else:
                    questionable_list.append((s,mod_time))
                
                rs_list.append((s,mod_time,ritm))  ##Problem when there is multiple RITMS in name

    out_df = pd.DataFrame(rs_list)
    out_df.columns = ['path','mod_date','ritm']
    #out_df_max = out_df.groupby(['ritm'])['mod_date'].max()
    
    ## Get the most recent report spec file
    out_df_max = out_df.sort_values('mod_date').groupby('ritm').tail(1)
    out_df_max.to_csv(out_file, index=False)
    
    ## output questionable files
    out_questionable_df = pd.DataFrame(questionable_list)
    out_questionable_df.columns = ['path','mod_date']
    
    out_questionable_df.to_csv(out_questionable_file, index=False)
    
    print('\n\nDone!')
    t = str(datetime.now() - start_main_Time)
    print('took %s' % t)
                         
                         
    pass





if __name__ == "__main__":
    search_files()