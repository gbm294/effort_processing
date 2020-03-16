# -*- coding: utf-8 -*-
"""
Created on Fri Sep 13 12:41:22 2019

Use ServiceNow time tracking report to compile and format
hours of effort for REDCap
@author: GBM294
"""


import tkinter
from tkinter import filedialog #GUI
import sys, os

import pandas as pd
pd.set_option("display.max_colwidth", 10000)  ##prevents strings from being truncated.
pd.set_option("display.max_columns", 100)  ##so all columns will show in print
from datetime import timedelta
from datetime import date





def get_file_list(what_file,parent_dir):
    #get input file
    root = tkinter.Tk()  #access windowing system
    effort_file = filedialog.askopenfilename(initialdir = parent_dir,title=what_file)
    user_dir = os.path.dirname(effort_file)
    #user_dir_name = os.path.basename(os.path.dirname(effort_file))
    file_path = effort_file #[0]
    root.destroy() #kill Tkinter
    if len(effort_file) <1 :
        print("you probably didn't choose any files")
        root.destroy() #kill Tkinter
        sys.exit()
    
    return file_path , user_dir


#Days of week for calculating date
DOW = {'sunday': 0, 'monday': 1, 'tuesday': 2, 'wednesday': 3, 'thursday': 4, 'friday': 5, 'saturday': 6}

# - Type of Service - Lookup table
TOS = {
       'Query': 2
       ,'Analysis': 3
       ,'Tools Development': 6
       ,'Teaching/Training': 7
       ,'Infrastructure/Operations/Support': 8
       ,'Learning': 9
       ,'Service': 10
       ,'Review': 11
       ,'QA': 12
       ,'Meeting': 13
       ,'Design': 14
       ,'Consult': 15
       ,'Other (Please Describe)': 16
       ,'Data': 18
       ,'Write': 19
       ,'Qlikview': 20 
       }
# - DAGs - Lookup table
DAG = {
       'Gabriel McMahan': 866
       ,'Sweta Singh': 1650
       ,'Thomas Callaci': 1624
       ,'Clark Xu': 1908
       ,'Oliver Eng': 1949
       ,'Mankee Wong':1947
       ,'Yonghe Yan':2097
       }



##Tie user to REDCap code and output in REDCap effor format.

def process_csv(in_path):
    df = pd.read_csv(in_path, sep=',', escapechar='\\', quotechar='"', encoding = "ISO-8859-1", keep_default_na =False)  #encoding = "ISO-8859-1"
    return df

def write_file(out_df, out_path):
    out_df.to_csv(out_path, index=False)
    return


def process_row(in_row, bmi_lookup):
    #print(in_row)
    DAG = {
       'Gabriel McMahan': ['gabe' , 866]
       ,'Sweta Singh': ['sweta' , 1650]
       ,'Thomas Callaci': ['tom_callaci' , 1624]
       ,'Clark Xu': ['clark_xu' , 1908]
       ,'Oliver Eng': ['oliver_eng', 1949]
       ,'Mankee Wong':['mankee_wong', 1947]
       ,'Yonghe Yan':['yonghe', 2097]
       }
    
    
    row_list = []
    days = [9,10,11,12,13,14,15]
    #print(in_row.iloc[0,1])
    try:
        ritm_id = bmi_lookup[bmi_lookup['Associated RITM #'] == in_row.iloc[0,2]].iloc[0,0]
    except:
        print(in_row.iloc[0,2] + ' not in REDCap')
        pass
    
    for i in days:
        #x = in_row.iloc[0,i]
        
        if in_row.iloc[0,i] > 0:
            out_line = {}
            try:
                dt = in_row.iloc[0,i + 7].strftime("%Y%m%d")
            except:
                print('problem')
            #print(in_row.iloc[0,i])
            
            out_line['record_id'] = str(str(DAG[in_row.iloc[0,0]][1]) +'-' + dt + '-' + str(in_row.iloc[0,3][4:]))  ##need DAG Lookup str(str(DAG[geffort_df['user'][r]])
            out_line['redcap_data_access_group'] = DAG[in_row.iloc[0,0]][0] #in_row['user'][i]
            out_line['date'] = in_row.iloc[0,i + 7].strftime("%m/%d/%Y") #in_row['monday_dt'][i].strftime("%m/%d/%Y")
            out_line['hours'] = in_row.iloc[0,i] # Hours for date
            try:
                out_line['customer'] = ritm_id
            except:
                out_line['customer'] = 'NOTFOUND'
                pass
            if DAG[in_row.iloc[0,0]][0] == 'tom_callaci':
                out_line['typeofservice'] = 11  #Service time of Review
            else:
                out_line['typeofservice'] = 2  #Service type of Query
            out_line['note'] = ''
            out_line['effort_complete'] = 0
            
            row_list.append(out_line)
        
    return row_list


def main():
    print("Let's go!")
    
    #parent_dir = "U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Effort_Tracking_SN"
    effort_file, parent_dir = get_file_list('SN Effort File' ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Manual_effort_calcs_gm')
    BMI_PRJ_File, BMI_PRJ_parent_dir = get_file_list('Updated REDCap Projects' ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Manual_effort_calcs_gm')
    
    os.chdir(parent_dir)
    os.chdir("..")
    working_dir = '%s/05_crunch_SN_effort_hrs_Output' % os.getcwd()
    os.chdir(working_dir)
    
    today = date.today().strftime("%Y-%m-%d")
    
    ## Output file
    effort_out_file = 'SN_to_RC_time_card_%s.csv' % today
    
    ## Read lookup
    bmi_prj_df = process_csv(BMI_PRJ_File)
    #print(bmi_prj_df.head(10))
    
    ## read effort file
    eff_df = process_csv(effort_file)
    #print(eff_df.head(10))
    print(eff_df.dtypes)
    
    #eff_df['week_starts_on'] = pd.to_datetime(eff_df['week_starts_on'], format='%Y-%m-%d') ##excel flips this when you save #format='%m/%d/%Y'
    eff_df['week_starts_on'] = pd.to_datetime(eff_df['week_starts_on'], format='%m/%d/%Y') ##format='%Y-%m-%d' #format='%m/%d/%Y'
    

    geffort_df = eff_df[eff_df['user'] != '']  ##make sure we have a user
    
    if geffort_df.shape[0] == 0:
        print('No Rows')
    
    geffort_df['sunday_dt'] =  geffort_df['week_starts_on'] + timedelta(days=0)
    geffort_df['monday_dt'] =  geffort_df['week_starts_on'] + timedelta(days=1)
    geffort_df['tuesday_dt'] =  geffort_df['week_starts_on'] + timedelta(days=2)
    geffort_df['wednesday_dt'] =  geffort_df['week_starts_on'] + timedelta(days=3)
    geffort_df['thursday_dt'] =  geffort_df['week_starts_on'] + timedelta(days=4)
    geffort_df['friday_dt'] =  geffort_df['week_starts_on'] + timedelta(days=5)
    geffort_df['saturday_dt'] =  geffort_df['week_starts_on'] + timedelta(days=6)

    # to hold output
    effort_rows = []
    
    ##Need to send all days of week to function     
    for r in geffort_df.index:
        if geffort_df['category'][r] == 'Request' and geffort_df['total'][r] > 0:
            #print(geffort_df.iloc[[r]])
            rl = process_row(geffort_df.iloc[[r]], bmi_prj_df)
            effort_rows.extend(rl)

    
    
    #process_row(in_row)
    effort_rows_df = pd.DataFrame(effort_rows)
    effort_rows_df['date_dt'] = pd.to_datetime(effort_rows_df['date']) ##adding for filtering

    effort_rows_df = effort_rows_df[effort_rows_df['date_dt']> '2019-09-01'] ##Only lool at records after a certain date
    effort_rows_df.drop(columns=['date_dt'], inplace = True)
    
    effort_rows_df = effort_rows_df.sort_values(by=['date']).reset_index(drop=True)
    write_file(effort_rows_df, effort_out_file)
    
    print('Done!')
    
    


if __name__ == "__main__":
    main()