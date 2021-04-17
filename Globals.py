import cx_Oracle,os
import Cipher
import sys,traceback

Key = 'QOAVLMCBKSWITZGYUFNPXRHJDE'
Work_dir_path = 'C:\\Murali\\Status'
conn_str =  'MA143NPV/MA143NPVEQU12HNJ@VID1800_MA144'
'''############### Issue Status ######################################### '''
Issues_Excel_Path = 'C:\\ChakraTeam-Share\\Testing_Share\\TD_Batch\\Issues.xlsx'
Status_Excel_Name = 'Status.xlsx' 
Issues_Excel_Name = 'Issues.xlsx'
Issues_html  = 'issues_status.html'
Issues_sheetList= ['Open Issues','Fixed Issues']
Issues_MsgBody = 'Please find the latest status of TD Java Testing.'
Issues_subject = 'TD-Java Testing Status at '
Footer_text = 'Please Update the Fix Details :  ChakraTeam-Share\Testing_Share\TD_Batch\issues.xlsx '
Issues_cols_reqd = [1,2,3,4,5]
#Issues_To_List = 'muralidharan.rengarajan@oracle.com'#,shenbagavel.g@oracle.com,shishira.shiannashetty@oracle.com,thirumoorthy.subramaniam@oracle.com,vivek.chari@oracle.com,susanta.k.patra@oracle.com,santosh.k.patel@oracle.com,rashmi.kundu@oracle.com,subhankar.chatterjee@oracle.com,justine.george@oracle.com,vidhya.athmakuri@oracle.com'
Issues_To_List = 'd.chakradhar@oracle.com,jeevit.stanley@oracle.com,muralidharan.rengarajan@oracle.com,lijo.james@oracle.com,rajamanickam.palani@oracle.com,umakant.naik@oracle.com,vidhya.athmakuri@oracle.com'

def print_log(str1, str2 = ''):

    print(str1)

def pr_sendMail_Plsql(p_sender,p_subject,p_Status_html,p_attchment_path,To_List):
    try:
        p_recipents = To_List
        fp  = open(p_Status_html)
        p_html_body = fp.read()
        fp.close()
        print_log('Start of pr_sendMail_Plsql')
        print_log('p_recipents',p_recipents)
        print_log('p_subject',p_subject)
        print_log('p_html_body length ',str(len(p_html_body)))
        p_attchment_path = '*'
        l_conn_str = Cipher.translateMessage(Key,conn_str,'D')
        connection1 = cx_Oracle.connect(l_conn_str)
        cur1 = connection1.cursor()
        cur1.callproc('pkg_send_mails.pr_send_mails',(p_sender,p_recipents, p_subject,p_html_body, p_attchment_path))
        cur1.close()
        connection1.close()
        print_log('completed pr_sendMail_Plsql')
    except:
        print_log('********* Failed in pr_sendMail_Plsql  **********')
        print_log('Unexpected error : {0}'.format(sys.exc_info()[0]))
        traceback.print_exc()
