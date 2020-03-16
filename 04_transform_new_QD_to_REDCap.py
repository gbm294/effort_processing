# -*- coding: utf-8 -*-
"""
Created on Fri Mar 13 11:17:12 2020
@author: gbm294

Need to transform QD_details scraped from U: drive to REDCap Projects format
"""

import sys, os
import tkinter
from tkinter import filedialog #GUI

import pandas as pd
pd.set_option("display.max_colwidth", 10000)  ##prevents strings from being truncated.

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

def process_csv(in_path):
    df = pd.read_csv(in_path, sep=',', escapechar='\\', quotechar='"', encoding = "ISO-8859-1", keep_default_na =False)  #encoding = "ISO-8859-1"
    return df



def sn_req_type(in_value):  ##'Type of request (count, data, recurring, etc)'
    d_type_of_consult = {
                    'ICTR - Count Request' : '1'
                    ,'ICTR - Data Request' : '2'
                    ,'ICTR - Data Request - Recurring' : '3'
                    ,'ICTR - Consult Request' : '4'
                    ,'REDCap database development' : '5'
                    ,'ICTR - Honest Broker Data Processing Request' : '6'
                    }
    if in_value == '':
        out_value = in_value
    else:
        out_value = d_type_of_consult[in_value]
    return out_value

def main():
    print("Let's go!")

    qd_path_list, qd_path_list_dir = get_file_list('QD Details List' ,'U:/UWHealth/EA/SpecialShares/DM/CRDS/AdHocQueries/ICTR Metrics Dashboard/Manual_effort_calcs_gm')
    
    os.chdir(qd_path_list_dir)
    os.chdir("..")
    working_dir = '%s/04_Format_details_for_REDCap_Output' % os.getcwd()
    os.chdir(working_dir)
    rc_format = 'New_RITMS_to_RC_format.csv'

    new_qd_in = process_csv(qd_path_list)
    
    ## For row in file, need to to output to REDCap format
        ##Switch headers
    new_qd_in.rename(columns={'ritm' : 'ritm'
                           ,'ritm_type' : 'count_data_recurring'
                           ,'created_date' : 'start'
                           ,'requested_for' : 'contact'
                           ,'short_description' : 'description'
                           ,'Project_Official_Title' : 'project_title'
                           ,'PI_or_Primary_Contact' : 'pi'
                           ,'Project_is_currently_Active' : 'active'
                           ,'irb' : 'irb'
                           }, inplace =True)
    ##new_qd_in.insert(loc, column, value)
    new_qd_in['record_id'] = 0
    new_qd_in['project'] = ''
    new_qd_in['end'] = ''
    new_qd_in['active'] = ''
    new_qd_in['type_request'] = ''
    new_qd_in['shared'] = ''
    new_qd_in['pcornet'] = ''
    new_qd_in['consult_only'] = ''
    new_qd_in['protocolid'] = ''
    new_qd_in['protocoltitle'] = ''
    new_qd_in['webcamppifirst'] = ''
    new_qd_in['webcamppilast'] = ''
    new_qd_in['projectfunding'] = ''
    new_qd_in['notes'] = ''
    new_qd_in['projects_complete'] = ''
    
    ## Reorder columns
    cols = ["record_id","project","project_title","contact","pi","description"
            ,"end","start","active","type_request","shared","pcornet"
            ,"consult_only","count_data_recurring","irb","ritm","protocolid"
            ,"protocoltitle","webcamppifirst","webcamppilast","projectfunding","notes","projects_complete"]
    new_qd_in = new_qd_in.reindex(columns=cols)
    
    new_qd_in['project'] = new_qd_in['ritm'] + ' - ' + new_qd_in['pi'] + ' - ' + new_qd_in['description'].str[0:45]
    new_qd_in['type_request'] = '4'
    new_qd_in['active'] = '1'
    new_qd_in['projects_complete'] = '2'
    new_qd_in['count_data_recurring'] = new_qd_in['count_data_recurring'].apply(sn_req_type)
    
    
    today = date.today().strftime("%Y%m%d")
    
    new_qd_in['record_id'] = today + new_qd_in.index.astype(str).str.zfill(5)



    print('Done!')
    new_qd_in.to_csv(rc_format, index=False)


    
if __name__ == "__main__":
    main()