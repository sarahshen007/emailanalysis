# program to complete email analysis
#
# should have the following features:
#   add on to 'CS Feedback' log with the following informational categories
#       date, issue summary, product pod, name, email, comment, ip, session
#   should also produce daily report
# 
#  
#   once you're done with the automation aspect,
#       do analysis of data



import os
import sqlite3
import pywin32
import tkinter
from email.message import _HeaderType
from email.message import Message

folder_path = os.path.normpath(askdirectory(title='Select Folder'))
email_list = [file for file in os.listdir(folder_path) if file.endswith(".msg")]


outlook = win32com.client.Dispatch("Outlook.Application").getNamespace("MAPI")

for i, _ in enumerate(email_list):
   # Create variable storing info from current email being parsed
   msg = outlook.OpenSharedItem(os.path.join(folder_path, email_list[i]))
   # Search email HTML for body text
   regex = re.search(r"<body([\s\S]*)</body>", msg.HTMLBody)
   body = regex.group()
   print(body)
