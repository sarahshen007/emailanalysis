# Module for email class

class Email: 
    def __init__(self, date="", issueSummary="", product="", name="", customerEmail="", comment="", ipAddress="", browser="", cookies="", followup=False):
        self.date = date
        self.issueSummary = issueSummary
        self.product = product
        self.name = name
        self.customerEmail = customerEmail
        self.comment = comment
        self.ipAddress = ipAddress
        self.browser = browser
        self.cookies = cookies
        self.followup = followup  
    

def emailCreator(info):
    date = info['date']
    name = info['First Name'] + ' ' + info['Last Name']
    name = name.title()
    issueSummary = info['Issue Summary']
    customerEmail = info['E-mail']
    comment = info['Comment Value']
    ipAddress = info['Extracted IP Address']
    browser = info['Extracted Browser/OS']
    cookies = info['Cookies']
    

    newEmail = Email(date, issueSummary, '', name, customerEmail, comment, ipAddress, browser, cookies, False)
    return newEmail


   
