#  Import packages
import re
import csv
import smtplib
import pandas as pd

from optparse import OptionParser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

#######
# Sources: https://gist.github.com/vjo/4119185
#          https://docs.python.org/3/library/smtplib.html
#          https://pandas.pydata.org/pandas-docs/stable/generated/pandas.read_csv.html
#
# NOTE: To enable this script to send mail you need to change the settings or
#       your gmail account. To do that:
#       
#       1) Click on your Google account icon (top right)
#       2) Click on My Account
#       3) From the Sign-In and Security group select "Apps with account access"
#       4) Enable "Allow less secure apps: ON"
#######

#######
# FUNCTION createMsg(row,o)
#######
def createMsg(row,o):
    # no hard conding of source file names
    html = open(o.text, 'r').read()
    ae_name_lastname = row['Opportunity Owner'].split()
    if pd.isnull(row['SE Opportunity Owner']):
        html = re.sub('#AE# #SE#', ae_name_lastname[0] + ',', html)
    else:
        se_name_lastname = row['SE Opportunity Owner'].split()
        html = re.sub('#AE#',ae_name_lastname[0], html)
        html = re.sub('#SE#','and ' + se_name_lastname[0] + ',', html)
    html = re.sub('#AN#', row['Account Name'], html)

    return (html)

#######
# FUNCTION sendMail(h, row)
#######
def sendMail(html,row, o):
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = 'Oracle at ' + row['Account Name']
    msgRoot['From'] = o.usr
    recipients = []
    if pd.notnull(row['Opportunity SE email']):
        recipients.append(row['Opportunity SE email'])
    recipients.append(row['Opportunity Owner Email'])
    msgRoot['To'] = ", ".join(recipients)
    msgRoot['Bcc'] = "ainfanzon@purestorage.com"

# Create message container.

    if o.image is not False:
        img = open(o.image, 'rb').read()
        msgImg = MIMEImage(img, 'png')
        msgImg.add_header('Content-ID', '<image1>')
        msgImg.add_header('Content-Disposition', 'inline', filename=o.image)

    msgHtml = MIMEText(html, 'html')
    msgRoot.attach(msgHtml)
    msgRoot.attach(msgImg)

# Record the MIME types.
# Send the message via local SMTP server.
    s = smtplib.SMTP('smtp.gmail.com:25')
    s.starttls()
    s.login(o.usr,o.passwd)    
# sendmail function takes 3 arguments: sender's address, recipient's address
# and message to send - here it is sent as one string.
    s.sendmail(o.usr, recipients, msgRoot.as_string())
    s.quit()

#<img src="cid:image1">

#######
# FUNCTION readFile()
#######
def readFile(o):
    df = pd.read_csv(o.csvfile, header=0, encoding='latin1')
    for index, row in df.iterrows():
        html = createMsg(row,o)
        if o.s:
            sendMail(html,row,o)
        else:
            fn = 'msg_f' + str(index) + '.html'
            text_file = open(fn, "w")
            text_file.write(html)
            text_file.close()

#######
# FUNCTION main()
######
def main():

    # Get command line argument values:
    parser = OptionParser()
    parser.add_option("-i", "--image", dest="image", help="image attachment", default=False)
    parser.add_option("-u", "--user", dest="usr", help="User ID", default=False)
    parser.add_option("-p", "--password", dest="passwd", help="Password", default=False)
    parser.add_option("-t", "--text", dest="text", help="HTML or plain text stored in a local file name", default=False)
    parser.add_option("-c", "--csv", dest="csvfile", help="CSV file", default=False)
    parser.add_option("-s", action="store_true", default=False)

    (options, args) = parser.parse_args()

    # just pass all the options to read file 
    #readFile(options.csvfile,options)
    readFile(options)

#######
# MAIN SECTION
######
if __name__== "__main__":
    main()
