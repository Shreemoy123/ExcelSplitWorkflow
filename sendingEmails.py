#!/usr/bin/env python
# coding: utf-8

# In[1]:


import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

import openpyxl as xl;
import logging
from FileDetails import FileDetailsUtility 
import os, uuid, sys



# In[2]:


class EmailSendingUtility:
    def getEmailBody(self,senderIdPwd,vendorId,vendorIndex,excelWorkbook):
        self.toaddr = str(vendorId[vendorIndex])
        msg = MIMEMultipart()    
        msg['From'] =  str(senderIdPwd[0])
        msg['To'] = self.toaddr 
        msg['Subject'] = "Subject of the Mail"
        body = "Body_of_the_mail"
        msg.attach(MIMEText(body, 'plain')) 
        #print(self.fileNameNewWorkbook)
        logging.debug('File Attachment Name : ',excelWorkbook)
        self.filename = excelWorkbook
        attachment = open(self.filename, "rb") 

        p = MIMEBase('application', 'octet-stream') 
        p.set_payload((attachment).read()) 
        encoders.encode_base64(p) 
        p.add_header('Content-Disposition', "attachment; filename= report_file.xlsx") 
        msg.attach(p)
        text = msg.as_string()
        return text

