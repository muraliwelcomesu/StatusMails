import Mail_Utils as mail
import os,openpyxl
import pandas as pd
import shutil
import Globals
from datetime import datetime

def get_curr_time():
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    return dt_string

def get_Sheet_Lists(ExcelName):
    wb = openpyxl.load_workbook(ExcelName)
    lst_sheets = wb.sheetnames
    print(lst_sheets)
    wb.close()
    return lst_sheets   
           
def create_Status_Excel(OutFileName):
    wb1 = openpyxl.Workbook()
    sheet1 = wb1.create_sheet()
    sheet1.title = 'TestCases-Summary'
    sheet = wb1.create_sheet()
    sheet.title = 'IssueDetails'
    wb1.remove(wb1['Sheet'])
    #wb1.remove_sheet(wb1.get_sheet_by_name('Sheet'))
    wb1.save(OutFileName)  

def prepare_Issues(ExcelName,OutFileName,p_status):
    print('Start of preparing {} Issues..'.format(p_status))
    lst_sheets = get_Sheet_Lists(ExcelName)
    l_cnt = 0
    SheetName = p_status + ' Issues'
    for sheet in lst_sheets:
        if sheet.startswith('#'):
            continue
        if l_cnt < 1:
            df = pd.read_excel(ExcelName,sheet_name =sheet,engine='openpyxl')
            df_open_issue = df[df['STATUS']==p_status]
            df_open_issue_1 =  df_open_issue[['NO','SFRNO','ISSUE_DESCRIPTION','SPC','CRITICAL']].sort_values('NO')
            #df_open_issue_1['NO'] = df_open_issue_1['NO'].apply(lambda x: "{}{}".format('SFR_', x))
        else:
            df_tmp = pd.read_excel(ExcelName,sheet_name =sheet)
            df_open_issue_tmp = df_tmp[df_tmp['STATUS']==p_status]
            df_open_issue_1_tmp =  df_open_issue_tmp[['NO','SFRNO','ISSUE_DESCRIPTION','SPC','CRITICAL']].sort_values('NO')
            #df_open_issue_1_tmp['NO'] = df_open_issue_1_tmp['NO'].apply(lambda x: "{}{}".format('SFR_', x))
            df_open_issue_1 = pd.concat([df_open_issue_1, df_open_issue_1_tmp], ignore_index=True)
        l_cnt += 1
    print('Completed...')
    return SheetName, df_open_issue_1
        
    
def prepare_status_excel():
    print('starting')
    shutil.copy(Globals.Issues_Excel_Path, Globals.Work_dir_path)
    os.chdir(Globals.Work_dir_path)
    create_Status_Excel(Globals.Status_Excel_Name)
    l_dict_sheets = {}
    l_sheetname, l_data_frame = prepare_Issues(Globals.Issues_Excel_Name,Globals.Status_Excel_Name,'Open')
    if len(l_data_frame) > 0:
        l_dict_sheets[l_sheetname] = l_data_frame    
    l_sheetname, l_data_frame = prepare_Issues(Globals.Issues_Excel_Name,Globals.Status_Excel_Name,'Fixed')
    if len(l_data_frame) > 0:
        l_dict_sheets[l_sheetname] = l_data_frame
    writer = pd.ExcelWriter(Globals.Status_Excel_Name, engine='xlsxwriter')
    for sheet,dataframe in l_dict_sheets.items():
        dataframe.to_excel(writer, sheet_name=sheet,index=False)
    writer.save()      
    print('Completed')
    
def issues_Status_Mail():
    prepare_status_excel()
    MsgBody = Globals.Issues_MsgBody
    p_sheetList= Globals.Issues_sheetList   
    html = Globals.Issues_html
    p_cols_reqd = Globals.Issues_cols_reqd
    subject = '{}   {} '.format(Globals.Issues_subject,get_curr_time())
    mail.conv_Excel_html_Issues(Globals.Status_Excel_Name,MsgBody,html,p_sheetList,p_cols_reqd)
    #mail.send_html_gmail(subject,html)
    sender = 'dailytestingstatusTD@noreply.oracle.com'
    Globals.pr_sendMail_Plsql(sender,subject,os.path.join(Globals.Work_dir_path,Globals.Issues_html),Globals.Issues_Excel_Path,Globals.Issues_To_List)
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    print('status sending completed... at {}'.format(dt_string))

if __name__ == "__main__":
    issues_Status_Mail()
    
