#!/usr/bin/python3

import sys
import getpass
import smtplib
import base64

print 'Office365 Email Sender Program\n'
print 'Please enter your credentials:\n'

user = raw_input("Username: ")
pwdEnc = base64.b64encode(getpass.getpass())
sendto = raw_input("Recipient: ")
subject = raw_input("Subject: ")
body = raw_input("Message: ")

try:
    ## Connecting to Office 365
    smtpsrv = "smtp.office365.com"
    smtpserver = smtplib.SMTP(smtpsrv,587)
    smtpserver.ehlo()
    smtpserver.starttls()
    smtpserver.ehlo
    smtpserver.login(user, base64.b64decode(pwdEnc))

except Exception as error:
    print ('\nERROR: ', error)

finally:
    try:
        ## Composing the email header and body.
        header = 'From: ' + user + '\r\n' + 'To: ' + sendto + '\r\n' + 'Subject: ' + subject + ' \n'

        ## Display how the email will be sent
        print '\nVerify your email:\n' + header + 'Message:\n' + body

        msgbody = header + '\n' + body + '\n\n'
        smtpserver.sendmail(user, sendto, msgbody)

    except Exception as error:
        print ('\nERROR: ', error)
        smtpserver.close()

    finally:
        print ('\nDone!')
        smtpserver.close()
