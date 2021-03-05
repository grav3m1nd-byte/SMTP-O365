#!/usr/bin/python3

import base64
import getpass
import smtplib


def main():
    print("Office365 Email Sender Program\n")
    print("\nPlease enter your credentials:\n")

    smtp_addr = "smtp.office365.com"

    error_header = "ERROR: "
    exit_msg = "Exiting!"

    try:
        user = check_email_contains(input("Username: "))
        pwd_enc = base64.b64encode(getpass.getpass().encode('UTF-8')).decode('ascii')
        sendto = check_rcpt(int(input("\nNumber of Recipient(s): ")))
        subject = input("\nSubject: ")
        body = input("Message: ")

    except Exception as error:
        print("\n" + error_header + error + "\n" + exit_msg)

    else:
        try:
            # Connecting to Office 365
            smtp_server = smtplib.SMTP(smtp_addr, 587)

            smtp_server.ehlo()
            smtp_server.starttls()
            smtp_server.ehlo
            smtp_server.login(user, base64.b64decode(pwd_enc.encode('ascii')).decode('UTF-8'))

        except smtplib.SMTPException as error:
            print("\n" + error_header + error + "\n" + exit_msg)

        else:
            try:
                # Composing the email header and body.
                header = "From: " + user + "\r\nTo: " + sendto + "\r\nSubject: " + subject + "\n"

                # Display how the email will be sent
                print("\nVerify your email:\n" + header + "Message:\n" + body)

                msgbody = header + "\n" + body + "\n\n"
                smtp_server.sendmail(user, sendto, msgbody)

            except smtplib.SMTPException as error:
                print("\n" + error_header + error + "\n" + exit_msg)

            else:
                print("\nDone!")
                smtp_server.close()


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
    i = 0
    while i < max_rcpt:
        rcp_list.append(check_email_contains(input("Provide recipient: ")))
        i += 1
    rcpts = str(",".join(rcp_list))
    return rcpts


if __name__ == "__main__":
    main()

