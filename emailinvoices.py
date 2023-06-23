# currently pauses if it attachment fails: 
# # # usually means invoice pdf can't be found -> put invoice in the folder and it can try again and continue
# # # ideally if you wanted you could also just continue without replacing the invoice, and go back to that one manually

# sometimes bmo doesn't need special handling

# TO DO
# # figure out how to automatically switch from field
# # -> currently skips known funny customers; could throw something to let the user know that external actions have to be taken
# # -> more explicit reporting on anything with a missing to field

# pull from onedrive and save as csv?
# make a guide to inputs
# organize: split functions between files / html


import win32com.client
import csv
import re
import sys

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

# input addressee, invno_list
def email_sender(key, value):
    successfulattach_count = 0
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
            print("added: " + key + " " + invoice)
            successfulattach_count += 1
        except Exception as inst:
            print("Error adding " + invoice)
            print(type(inst))
            print(inst.args)
            # try again
            if input("Fix it! Type n to exit. Press enter to continue. ") != "n":
                try:
                    newmail.Attachments.Add(mailattachment)
                    print("added: " + key + " " + invoice)
                    successfulattach_count += 1
                except Exception as inst:
                    print(type(inst))
                    print(inst.args)
                    print(invoice)
                    if input("It didn't work. Type n to exit. Press enter to continue without fixing. ") == "n":
                        sys.exit()   
                # will this keep working if you don't fix it? ideally
            else:
                sys.exit()

    return newmail, successfulattach_count

ol = win32com.client.Dispatch('Outlook.Application')

# ope global variable
manyemail_dict = {}
extrahandling = []

# open invoicing csv
with open(r"C:\Users\Annika - Accounting\Email Automation\INVOICING.csv", "r") as file:
    csvreader = csv.reader(file)
    next(csvreader)
    # create dictionary of emails and invoice numbers
    for row in csvreader:
        if any(row):
            if row[6] not in ["1414", "9289", "1955", "1178", "8923"]: # excepting mechanics bank, associated bank, td bank, fifth third

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
            else:
                extrahandling += row[6] + ": extra handling required for " + row[12] + " : " + row[0] +"\n"
    
    # start counter to make sure every email is sent
    returnedcount = 0
    mail_object_list = []
    # for each email address, create an email with appropriate subject (citing invoice numbers), body (plural or not), and attached invoice pdfs
    for key, value in manyemail_dict.items():
        try:
            mail_object, attachment_count = email_sender(key, value)
            mail_object_list.append(mail_object)
            returnedcount += attachment_count # the count is the number of invoices that have been successfully attached
        except Exception as inst:
            print(type(inst))
            print(inst.args)
    print("----------------------------")
    maildisplaycount = 0
    for mail in mail_object_list:
        # try displaying the email
        try:
            mail.Display()
            maildisplaycount += 1
            print("sent!")
        except Exception as inst:
            print(type(inst))
            print(inst.args)
        # if input("Next? Type n to exit. Press enter to continue. ") == "n":
        #        sys.exit()
        
    print("----------------------------")
    print("total invoices sent")
    print(returnedcount)
    print("total emails sent")
    print(maildisplaycount)
    print("----------------------------")
    print("".join(extrahandling))

# follow-up manual process: change from field and send email