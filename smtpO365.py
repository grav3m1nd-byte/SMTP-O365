#!/usr/bin/python3

import base64
import getpass
import smtplib


def main():
    print("Office365 Email Sender Program\n")
    print("\nPlease enter your credentials:\n")

    smtpADDR = "smtp.office365.com"

    try:
        user = check_email_contains(input("Username: "))
        pwdEnc = base64.b64encode(getpass.getpass().encode('UTF-8')).decode('ascii')
        sendto = check_rcpt(int(input("\nNumber of Recipient(s): ")))
        subject = input("\nSubject: ")
        body = input("Message: ")

    except Exception as error:
        print("\nERROR: ")
        print(error)
        print("\nExiting!")

    else:
        try:
            # Connecting to Office 365
            smtpServer = smtplib.SMTP(smtpADDR, 587)

            smtpServer.ehlo()
            smtpServer.starttls()
            smtpServer.ehlo
            smtpServer.login(user, base64.b64decode(pwdEnc.encode('ascii')).decode('UTF-8'))

        except smtplib.SMTPException as error:
            print("\nERROR: ")
            print(error)
            print("\nExiting!")

        else:
            try:
                # Composing the email header and body.
                header = "From: " + user + "\r\nTo: " + sendto + "\r\nSubject: " + subject + "\n"

                # Display how the email will be sent
                print("\nVerify your email:\n" + header + "Message:\n" + body)

                msgbody = header + "\n" + body + "\n\n"
                smtpServer.sendmail(user, sendto, msgbody)

            except smtplib.SMTPException as error:
                print("\nERROR: ")
                print(error)
                print("\nExiting!")

            else:
                print("\nDone!")
                smtpServer.close()


def check_email_contains(email_address="test@domain.com", min_length=6):
    CHARACTERS = "@."
    while True:
        for character in CHARACTERS:
            if character not in email_address:
                email_address = input("\nERROR: Your input must have '{}' in it,"
                                      "\n\tfollowing the email address formatting."
                                      "\n\nPlease write the email address again: ".format(character))
                continue
            if len(email_address) <= min_length:
                email_address = input("ERROR: Your input as email address is too short."
                                      "\n\nPlease write the email address again: ")
                continue
            return email_address


def check_rcpt(max_rcpt=1):
    rcp_list = [check_email_contains(input("Provide recipient: "))]
    for i in range(1, max_rcpt):  # for(i = 1, i < max_rcpt, i += 1)
        rcp_list.append(check_email_contains(input("Provide recipient: ")))
    rcpts = str(",".join(rcp_list))
    return rcpts


if __name__ == "__main__":
    main()

