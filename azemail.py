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
import re
import azsort
import summarize
import time
import win32com
import openpyxl
from win32com import client
from tkinter import filedialog
from bs4 import BeautifulSoup
from colorama import Fore
from colorama import Style

# Message to user
print(f"{Fore.LIGHTBLUE_EX}===\n===")
print(f"{Fore.LIGHTBLUE_EX}Welcome to the {Fore.RED}AZ Email Analysis {Fore.LIGHTBLUE_EX}Program!")
print("This program will: \n\t1. read all the emails (.msg files) in your local documents folder\n\t2. process the data\n\t3. add entries to the end of a spreadsheet.\nPlease select the folder containing all your emails.")
print("===\n===\n")
time.sleep(1)

# Create an folder input dialog with tkinter
folder_path = os.path.normpath(filedialog.askdirectory(title='Select Folder'))
print("Nice! Creating db...\n")

# Initialise & Populate list of emails
email_list = [file for file in os.listdir(folder_path) if file.endswith(".msg")]

# Connect to Outlook with MAPI
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Log keeping track of email objects
emails_log = []

# Iterate through every email
for i, _ in enumerate(email_list):

   # Create variable storing info from current email being parsed
   msg = outlook.OpenSharedItem(os.path.join(folder_path, email_list[i]))


   # Dictionary to keep track of info from current email
   info = {}
   info['date'] = str(msg.SentOn).split(' ')[0] # Date email was received
   
   regex = msg.HTMLBody.replace('\r', '').replace('\n', '') # Remove unecessary characters from msg html
   soup = BeautifulSoup(regex, "html.parser") # Parse into html using soup
    

    # Create list of category + values
   texts = str(soup.find_all('font')[0].encode_contents()).strip('b').strip('\'').strip('\"').replace('<br/>', '\n').strip().split('\n') 
   texts = list(filter(None, texts))

   # Create list of pairs to populate info dictionary
   pairs = []
   
   # Edit list for unwanted extra elements caused by extra break elements
   lastKey = ""
   for data in texts:
      pair = data.split(':', 1)
      if len(pair) == 1:
         info[lastKey] = info[lastKey] + pair[0]
      elif len(pair) == 2: 
         lastKey = pair[0].strip()
         info[lastKey] = pair[1].strip()


   # Generate summary of comment
   summary = info['Comment Value']
   summary = summarize.summarize(summary, 0.3)
   info['Issue Summary'] = summary
     

   newEmail = azsort.emailCreator(info)

   # Add email object to emails log
   emails_log.append(newEmail)
   

try:
   print("===\n===")
   print("Please choose the spreadsheet where you keep logs.")
   print("===\n===\n")
   time.sleep(1)
   
   excelPath = os.path.normpath(filedialog.askopenfilename(title='Select File'))
   print("Nice! Adding to spreadsheet now...\n")
   wb = openpyxl.load_workbook(excelPath) 
   
   sheet = wb.active 
   
   data = []

   for email in emails_log:
      row = (email.date, email.issueSummary, email.product, email.name, email.customerEmail, email.comment, email.ipAddress, email.browser, email.cookies, email.followup)
      data.append(row)
   
   for row in data:
      sheet.append(row)
   
   wb.save(excelPath)
   wb.close()
except:
   print("Error adding to spreadsheet. Please check that you chose a valid file.\n")


print("===\n===\n===")
print("All done! Thank you for using AZ Email Analysis!")
print("===\n===\n===\n")

