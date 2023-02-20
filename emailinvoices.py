# TO DO
# # figure out how to automatically switch from field
# # create way to skip to a given line in case the program exits early
# # expand handling of rows with missing data

import win32com.client
import csv
import re

html_body = """
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Hello,<br>&nbsp;<br>&nbsp;Please see the attached invoice for your order. If you have any questions please contact us at your earliest convenience.<br>&nbsp;<br>&nbsp;We appreciate your business!</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Sincerely,<br>&nbsp;The KANE Accounting Team<br>&nbsp;<a href="mailto:ar@kanegraphical.com"><span style="color:#0563C1;">ar@kanegraphical.com</span></a>&nbsp;</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>Kane Graphical Corporation&nbsp;</span></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>2255 W. Logan Blvd. | Chicago, IL 60647-2114</span></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><a href="http://www.kanegraphical.com/"><span style='font-size:12px;font-family:"Arial",sans-serif;color:#F58025;'>www.kanegraphical.com</span></a></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>&nbsp;</span></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>773-384-1207 fax | 773-384-1200 main | 800-992-2921 toll free</span></p>
"""

plural_html_body = """
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Hello,<br>&nbsp;<br>&nbsp;Please see the attached invoices for your recent orders. If you have any questions please contact us at your earliest convenience.<br>&nbsp;<br>&nbsp;We appreciate your business!</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Sincerely,<br>&nbsp;The KANE Accounting Team<br>&nbsp;<a href="mailto:ar@kanegraphical.com"><span style="color:#0563C1;">ar@kanegraphical.com</span></a>&nbsp;</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>Kane Graphical Corporation&nbsp;</span></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>2255 W. Logan Blvd. | Chicago, IL 60647-2114</span></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><a href="http://www.kanegraphical.com/"><span style='font-size:12px;font-family:"Arial",sans-serif;color:#F58025;'>www.kanegraphical.com</span></a></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>&nbsp;</span></p>
    <p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:12px;font-family:"Arial",sans-serif;color:#595959;'>773-384-1207 fax | 773-384-1200 main | 800-992-2921 toll free</span></p>
"""

def email_finder(text):
    all_lowercase = text.lower()
    if "travis.powers@bancfirst.bank" in all_lowercase:
        return "travis.powers@bancfirst.bank" # I don't know why regex breaks :'(
    else:
        emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", all_lowercase)
        email_string = "; ".join(emails)
        return email_string

def email_sender(key, value):
     # create new email
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)

    # defining email fields
    emailsubject = "Kane Graphical - " + ", ".join(value)

    newmail.Subject = emailsubject
    newmail.To = key
    if len(value) > 1:
        newmail.HTMLBody = plural_html_body
    else:
        newmail.HTMLBody = html_body

    # attachments
    for invoice in value: 
        attachmentpath = 'C:/Users/Annika - Accounting/Email Automation/Invoices/' + invoice + '.pdf'
        mailattachment = attachmentpath.encode('unicode-escape').decode()
        try:
            newmail.Attachments.Add(mailattachment)
        except Exception as inst:
            print(type(inst))
            print(inst.args)
    
    # try displaying the email
    try:
        newmail.Display()
    except Exception as inst:
        print(type(inst))
        print(inst.args)
    # if input("Next? Type n to exit. Press enter to continue. ") == "n":
    #        break

ol = win32com.client.Dispatch('Outlook.Application')

# ope global variable
manyemail_dict = {}

# open invoicing csv
with open(r"C:\Users\Annika - Accounting\Email Automation\INVOICING.csv", "r") as file:
    csvreader = csv.reader(file)
    next(csvreader)
    # create dictionary of emails and invoice numbers
    for row in csvreader:
        if any(row):
            if row[6] != "1414": # excepting mechanics bank
                email = email_finder(row[14])
                invoiceno = row[12]
                if email in manyemail_dict.keys():
                    # append invoice no to email key's value list
                    manyemail_dict[email].append(invoiceno)
                else:
                    # create list of invoice numbers
                    invoicelist = []
                    invoicelist.append(invoiceno)
                    # add email key with invoice number list value
                    manyemail_dict.update({email : invoicelist}) # should work even with multiple emails, as email_finder returns a string
    # for each email address, create an email with appropriate subject (citing invoice numbers), body (plural or not), and attached invoice pdfs
    for key, value in manyemail_dict.items():
        email_sender(key, value)

# follow-up manual process: change from field and send email