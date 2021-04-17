import os,openpyxl
import smtplib 
import socket
from email.mime.text import MIMEText
from datetime import datetime
import Globals
import schedule
import shutil
import pandas as pd
import Cipher as cp

def recvline(sock):
    """Receives a line."""
    stop = 0
    line = ''
    while True:
        i = sock.recv(1)
        if i.decode('UTF-8') == '\n': stop = 1
        line += i.decode('UTF-8')
        if stop == 1:
            print('Stop reached.')
            break
    print('Received line: %s' % line)
    return line


class ProxySMTP(smtplib.SMTP):
    """Connects to a SMTP server through a HTTP proxy."""

    def __init__(self, host='', port=0, p_address='',p_port=0, local_hostname=None,
             timeout=socket._GLOBAL_DEFAULT_TIMEOUT):
        """Initialize a new instance.

        If specified, `host' is the name of the remote host to which to
        connect.  If specified, `port' specifies the port to which to connect.
        By default, smtplib.SMTP_PORT is used.  An SMTPConnectError is raised
        if the specified `host' doesn't respond correctly.  If specified,
        `local_hostname` is used as the FQDN of the local host.  By default,
        the local hostname is found using socket.getfqdn().

        """
        self.p_address = p_address
        self.p_port = p_port

        self.timeout = timeout
        self.esmtp_features = {}
        self.default_port = smtplib.SMTP_PORT

        if host:
            (code, msg) = self.connect(host, port)
            if code != 220:
                raise IOError(code, msg)

        if local_hostname is not None:
            self.local_hostname = local_hostname
        else:
            # RFC 2821 says we should use the fqdn in the EHLO/HELO verb, and
            # if that can't be calculated, that we should use a domain literal
            # instead (essentially an encoded IP address like [A.B.C.D]).
            fqdn = socket.getfqdn()

            if '.' in fqdn:
                self.local_hostname = fqdn
            else:
                # We can't find an fqdn hostname, so use a domain literal
                addr = '127.0.0.1'

                try:
                    addr = socket.gethostbyname(socket.gethostname())
                except socket.gaierror:
                    pass
                self.local_hostname = '[%s]' % addr

        smtplib.SMTP.__init__(self)

    def _get_socket(self, port, host, timeout):
        # This makes it simpler for SMTP to use the SMTP connect code
        # and just alter the socket connection bit.
        print('Will connect to:', (host, port))
        print('Connect to proxy.')
        new_socket = socket.create_connection((self.p_address,self.p_port), timeout)

        s = "CONNECT %s:%s HTTP/1.1\r\n\r\n" % (port,host)
        s = s.encode('UTF-8')
        new_socket.sendall(s)

        print('Sent CONNECT. Receiving lines.')
        for x in range(2): recvline(new_socket)

        print('Connected.')
        return new_socket
    
def conv_Excel_Dict(ExcelName,SheetName,p_col_list):
    wb = openpyxl.load_workbook(ExcelName,data_only=True)
    sheet = wb.get_sheet_by_name(SheetName)
    lst_row = []
    dict_sheet = {}
    l_rownum = 0
    for row in sheet.rows:
        l_rownum = l_rownum + 1
        l_col_cnt = 0
        for cell in row:
            l_col_cnt = l_col_cnt + 1
            if l_col_cnt in p_col_list:
                if cell.value is None:
                    cell.value = ' '
                lst_row.append(cell.value)
        dict_sheet[l_rownum] = lst_row
        lst_row = []
    return dict_sheet

def build_header(p_dict_out,p_header):
    l_html_str = '<span style = "color:black">Hi,<br> {} <br> <br> \n'.format(p_header)
    p_dict_out['Header'] = l_html_str

def build_footer(p_dict_out,p_name):
    
    l_html_str = '<span style = "color:black"> <br> <b> {} </b> <br><br>\n'.format(Globals.Footer_text)
    l_html_str = l_html_str + '<span style = "color:black">Thanks and Regards<br> {} <br><br>\n'.format(p_name)
    p_dict_out['Footer'] = l_html_str

def write_dict_htlmfile(p_dict,htmlFile):
    fp = open(htmlFile,'w')
    for key,value in p_dict.items():
        fp.write(value + '\n')
    fp.close()
          
def Conv_Dict_HTMLDict_Issues(p_dict,Title,p_dict_out):
    l_html_str = '<span style = "color:black"><b>{}</b><br>\n'.format(Title)
    l_html_str = l_html_str + '<br>'
    l_html_str = l_html_str + '<html><table border = 1>\n'
    #----------------------------------------------------------------------------------
    for key,values in p_dict.items():
        if str(key)=='1':
            #print('first row')
            l_col_str = ''
            l_cntr = 0
            for value in values:
                
                if l_cntr != (len(values) - 1):
                    l_col_str = l_col_str + '<td><b><span style="color:black" >{}</b></td>\n'.format(value)
                l_cntr = l_cntr + 1
            #print(l_col_str)
            l_html_str = l_html_str + '<tr style="background-color:red">{}</tr>\n'.format(l_col_str)
        else:
            l_col_str = ''
            if values[-1] == 'Y':
                l_cntr = 0 
                for value in values:
                    
                    if l_cntr != (len(values) - 1):
                        l_col_str = l_col_str + '<td align = "left"><b><span style = "color:black">{}</b></td>\n'.format(value)
                    l_cntr = l_cntr + 1
            else:
                l_cntr = 0 
                for value in values:
                    
                    if l_cntr != (len(values) - 1):
                        l_col_str = l_col_str + '<td align = "left"><span style = "color:black">{}</td>\n'.format(value)
                    l_cntr = l_cntr + 1
            #print('else part ' + l_col_str)
            l_html_str = l_html_str +  '<tr style="background-color:white">{}</tr>\n'.format(l_col_str)
    #----------------------------------------------------------------------------------
    l_html_str = l_html_str + '</table></html><br>\n'
    p_dict_out[Title] = l_html_str


def Conv_Dict_HTMLDict(p_dict,Title,p_dict_out):
    l_html_str = '<span style = "color:black"><b>{}</b><br>\n'.format(Title)
    l_html_str = l_html_str + '<br>'
    l_html_str = l_html_str + '<html><table border = 1>\n'
    #----------------------------------------------------------------------------------
    for key,values in p_dict.items():
        if str(key)=='1':
            #print('first row')
            l_col_str = ''
            for value in values:
                    l_col_str = l_col_str + '<td><b><span style="color:black" >{}</b></td>\n'.format(value)
            #print(l_col_str)
            l_html_str = l_html_str + '<tr style="background-color:red">{}</tr>\n'.format(l_col_str)
        else:
            l_col_str = ''
            if values[-1] == 'Y':
                for value in values:
                        l_col_str = l_col_str + '<td align = "left"><b><span style = "color:black">{}</b></td>\n'.format(value)
            else:
                for value in values:
                        l_col_str = l_col_str + '<td align = "left"><span style = "color:black">{}</td>\n'.format(value)
            #print('else part ' + l_col_str)
            l_html_str = l_html_str +  '<tr style="background-color:white">{}</tr>\n'.format(l_col_str)
    #----------------------------------------------------------------------------------
    l_html_str = l_html_str + '</table></html><br>\n'
    p_dict_out[Title] = l_html_str
    
def conv_Excel_html(ExcelName,subject,htmlFile,p_sheetList,p_cols_reqd):
    Globals.print_log('starting preparing html file')     
    os.chdir(Globals.Work_dir_path)  
    l_html_dict = {}
    build_header(l_html_dict,subject)
    wb = openpyxl.load_workbook(ExcelName,data_only=True)
    for sheet in wb.sheetnames:
        if sheet in p_sheetList:
            try:
                l_cols_reqd = p_cols_reqd[sheet]
            except:
                l_cols_reqd = [1,2,3,4]
                
            dict_sheet = conv_Excel_Dict(ExcelName,sheet,l_cols_reqd)
            Conv_Dict_HTMLDict(dict_sheet,sheet,l_html_dict)
    build_footer(l_html_dict,'Muralidharan R')
    write_dict_htlmfile(l_html_dict,htmlFile)
    Globals.print_log('Completed preparing html file')

def conv_Excel_html_Issues(ExcelName,subject,htmlFile,p_sheetList,p_cols_reqd):
    Globals.print_log('starting preparing html file')     
    os.chdir(Globals.Work_dir_path)  
    l_html_dict = {}
    build_header(l_html_dict,subject)
    wb = openpyxl.load_workbook(ExcelName,data_only=True)
    for sheet in wb.sheetnames:
        if sheet in p_sheetList:
            dict_sheet = conv_Excel_Dict(ExcelName,sheet,p_cols_reqd)
            Conv_Dict_HTMLDict_Issues(dict_sheet,sheet,l_html_dict)
    build_footer(l_html_dict,'Muralidharan R')
    write_dict_htlmfile(l_html_dict,htmlFile)
    Globals.print_log('Completed preparing html file')
    
def send_html_gmail(subject,htmlFileName):
    os.chdir(Globals.Work_dir_path)
    l_from = cp.translateMessage(Globals.Key,Globals.email_id, 'D')
    l_passwd = cp.translateMessage(Globals.Key,Globals.email_passwd, 'D')
    with open(htmlFileName,mode = "rb") as message:    #open report html for reading
        msg = MIMEText(message.read(),'html','html') # create html message
    msg['Subject'] = subject
    msg['From'] = l_from
    msg['To'] = Globals.To_List
    proxy_host = 'www-proxy-idc.in.oracle.com'
    proxy_port = 80
    smtpObj = ProxySMTP(host='smtp.gmail.com', port=587,p_address=proxy_host, p_port=proxy_port)
    smtpObj.starttls()
    smtpObj.login(l_from,l_passwd )
    smtpObj.send_message(msg)
    smtpObj.close()

              
def schedule_job():
    print('\n \n \n \n')
    print('Please check with Murali before closing this...')
    def job():
        print('Start of  sending status...')
        issues_Status_Mail()
        print('\n \n *********************************************** \n \n ')
        
    #schedule.every().day.at("22:00").do(job)
    job()
    
    '''schedule.every(10).minutes.do(job)
    schedule.every().hour.do(job)
    schedule.every().day.at("10:30").do(job)
    schedule.every().monday.do(job)
    schedule.every().wednesday.at("13:15").do(job)'''
    
    '''while True:
        schedule.run_pending()
        time.sleep(60)'''
               
if __name__ == "__main__":
    schedule_job()
