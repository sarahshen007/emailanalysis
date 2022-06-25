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
import azsort
import azsummary
import time
import win32com
from win32com import client
import openpyxl
from tkinter import filedialog
from bs4 import BeautifulSoup
from colorama import Fore

# Loops through all Outlook folders to find given folder_name
def get_folder_by_name(folder_name, root_folder):

    for folder in root_folder.Folders: 

        if folder.Name == folder_name:

            print ('Awesome, ' + folder_name + ' was found!')

            found_folder = folder

    return found_folder

# Message to user
print(f"{Fore.LIGHTBLUE_EX}=========")
print(f"{Fore.LIGHTBLUE_EX}Welcome to the {Fore.RED}AZ Email Analysis {Fore.LIGHTBLUE_EX}Program!")
print("=========\n")
time.sleep(1)
print("This program will: \n\t1. read all the emails (.msg files) in your local documents folder\n\t2. process the data\n\t3. add entries to the end of a spreadsheet.\n")
time.sleep(1)
print("Please select the folder containing all your emails...")
time.sleep(1)

# Connect to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts

account = accounts[0]
print("Welcome,", account.DisplayName)
root_folder = outlook.Folders(account.DisplayName)
emails_folder = get_folder_by_name("CS EMAILS", root_folder)

messages = emails_folder.Items

print("Please choose the spreadsheet where you keep logs...")
excelPath = os.path.normpath(filedialog.askopenfilename(title='Select File'))

print("Thank you for choosing your spreadsheet file.\n")
time.sleep(1)
print("Parsing spreadsheet...")

# Initialise & Populate list of emails
email_list = [file for file in messages]

# Log keeping track of email objects
emails_log = []

print("Generating data...")
# Data for prediction of issue
prev_data = azsummary.generateData(excelPath)

print("Finished parsing spreadsheet.")
print()
print("Parsing emails...")
for x in prev_data.keys():
   print(x)
   for y in prev_data[x].keys():
      print("\t",y, ":", prev_data[x][y])

# Iterate through every email
for i, _ in enumerate(email_list):

   # Create variable storing info from current email being parsed
   msg = outlook.OpenSharedItem(os.path.join(folder_path, email_list[i]))


   # Dictionary to keep track of info from current email
   info = {}
   date = str(msg.SentOn).split(' ')[0].split('-')
   dateString = date[1].strip('0') + "/" + date[2] + "/" + date[0]
   info['date'] = dateString # Date email was received
   
   regex = msg.HTMLBody.replace('\r', '').replace('\n', '') # Remove unecessary characters from msg html
   soup = BeautifulSoup(regex, "html.parser") # Parse into html using soup
    

    # Create list of category + values
   texts = str(soup.find_all('font')[0].encode_contents(encoding='utf-8')).strip('b').strip('\'').strip('\"').replace('<br/>', '\n')
   texts = azsort.replaceCharacters(texts)
   texts = texts.strip().split('\n')
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

   predictedIssue = azsummary.generateIssueSummary(summary, prev_data)
   print(predictedIssue)

   info['Issue Summary'] = predictedIssue[0]
   info['Product'] = predictedIssue[1]
   
   # Make new email object with info
   newEmail = azsort.emailCreator(info)

   # Add email object to emails log
   emails_log.append(newEmail)

print("Finished parsing emails.\n")
try:
   print("Adding to log to spreadsheet...\n")
   wb = openpyxl.load_workbook(excelPath) 
   
   sheet = wb.active 
   
   data = []

   for email in emails_log:
      row = (email.date, email.issueSummary, email.product, email.name, email.customerEmail, email.comment, email.ipAddress, email.cookies, email.followup)
      data.append(row)
   
   for row in data:
      sheet.append(row)
   
   wb.save(excelPath)
   wb.close()

except:
   print("Error adding to spreadsheet. Please check that you chose a valid file and that it isn't currently open.\n")


print("=========")
print("All done! Thank you for using AZ Email Analysis!")
print("=========")
