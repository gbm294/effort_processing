# -*- coding: utf-8 -*-
"""
Created on Wed Sep  4 16:11:25 2019

Use reports exported from ServiceNow and REDCap
Compare RITMs, then output a csv file with a list of RITMs that don't exist in REDCap.
@author: gbm294
"""

import tkinter
from tkinter import filedialog #GUI
import pandas as pd
import sys, os
from datetime import date



def process_csv(in_path):
    df = pd.read_csv(in_path, sep=',', escapechar='\\', quotechar='"', encoding = "ISO-8859-1", keep_default_na =False)  #encoding = "ISO-8859-1"
    return df


def write_file(out_df, out_path):
    out_df.to_csv(out_path, index=False)
    return


def return_ritm(prj_name):
    if prj_name[:4].upper() == 'RITM':    
        return prj_name[:11]
    else:
        return 'not_RITM'


def get_file_list(what_file,parent_dir):
    #get input file
    root = tkinter.Tk()  #access windowing system
    effort_file = filedialog.askopenfilenames(initialdir = parent_dir,title=what_file)
    user_dir = os.path.dirname(effort_file[0])
    #os.chdir(user_dir)  #specify directory
    user_dir_name = os.path.basename(os.path.dirname(effort_file[0]))
    file_path = effort_file[0]
    root.destroy() #kill Tkinter
    
    if len(effort_file) <1 :
        print("you probably didn't choose any files")
        root.destroy() #kill Tkinter
        sys.exit()
    
    return file_path, user_dir


def main():
    print("Let's go!")
    
    ##specify input csv files
    ## Go to RITM file that was just downloaded
    #SN_ritm_file_path = 'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Effort_tracking_REDCap/2019_09_04/all_ritms_2019_09_23.csv'
    SN_ritm_file_path, SN_ritm_file_dir = get_file_list('Get ServiceNow RITM file' ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Manual_effort_calcs_gm')
    #print(SN_ritm_file_path)
    
    
    ## Go to REDCap projects file that was just downloaded
    #REDCap_projects_file_path = 'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Effort_tracking_REDCap/2019_09_04/PRODUCTIONBMICProjec_DATA_LABELS_2019-09-23_1258.csv'
    REDCap_projects_file_path, working_dir = get_file_list('Get REDCap PRJ file' ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Manual_effort_calcs_gm')
    #print(REDCap_projects_file_path)
    
    os.chdir(working_dir)
    os.chdir("..")
    working_dir = '%s/01_Find_RITMs_not_in_REDCap_from_SN_Output' % os.getcwd()
    os.chdir(working_dir)

    today = date.today().strftime("%Y-%m-%d")
    ##print(os.getcwd())

    #out_new_ritms = 'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Effort_tracking_REDCap/2019_09_04/new_ritms_2019_09_23_python.csv'
    out_new_ritms = 'new_ritms_%s_python.csv' % today
    
    #out_rc_prj_recordid = 'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Effort_tracking_REDCap/2019_09_04/RC_crosswalk_2019_09_23_python.csv'
    out_rc_prj_recordid = 'RC_crosswalk_%s_python.csv' % today
    
        
    ##bring csv files in as dataframes
    sn_df = process_csv(SN_ritm_file_path)
    rc_df = process_csv(REDCap_projects_file_path)
    
    
    ##Get RITM from REDCap file
    rc_df['RITM'] = rc_df['Project Short Name/RITM Number'].apply(return_ritm)
    
    
    #matching_ritms = pd.merge(sn_df, rc_df, left_on='number', right_on='RITM')['RITM']
    new_ritms_df = pd.merge(sn_df, rc_df, left_on='number', right_on='RITM', how='left')  ##Joining SN and RC files
    new_ritms_df = new_ritms_df[new_ritms_df['RITM'].isnull()].reset_index(drop=True).rename(columns={"number": "SN_RITM"})  ##Where RITM in RC is null  ##new_ritms[new_ritms['RITM'].notnull()]
    new_ritms_df = new_ritms_df[new_ritms_df['sys_created_on'] >= '2019-09-01']

    cols = [9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32]
    new_ritms_df = new_ritms_df.drop(new_ritms_df.columns[cols], axis=1)

    print(os.getcwd())    
    write_file(new_ritms_df, out_new_ritms)


    ##Get Existing Record_IDs
    rids_df = rc_df[['RITM','Record ID']]
    write_file(rids_df, out_rc_prj_recordid)
    
    
    print('Dude')
    
    
    
    
    
if __name__ == "__main__":
    main()