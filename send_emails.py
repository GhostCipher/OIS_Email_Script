import pandas as pd
import re
from time import sleep
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import argparse

# This script requires the following packages:
# pandas, openpyxl

parser = argparse.ArgumentParser()
parser.add_argument('-m', metavar='Max Age', action='store', default=7, type=int, help='Skips rows older than this (Default 7)')
parser.add_argument('-r', dest='reporter', type=str, metavar='Reporter', default='rsalazar15', action='store', help='Jira reporter')
parser.add_argument('-a', dest='assignee', type=str, metavar='Assignee', default='rsalazar15', action='store', help='Jira assignee')
parser.add_argument('-d', dest='delay', type=int, metavar='Delay', default=2, action='store', help='Delay between emails in seconds (Default 2)')
group = parser.add_mutually_exclusive_group()
group.add_argument('-v', '--verbose', action='store_true')
group.add_argument('-q', '--quiet', action='store_true')
parser.add_argument('files', nargs='*', type=str, help="List of Execl files")


def fix_name(name):
    temp = name.split(", ", 1)
    return f"{temp[1]} {temp[0]}"


def send_mail(arg):
    ois_email = "Office-Of-Information-Security@tamucc.edu"
    jira_email = "jira@jira.tamucc.edu"
    service_desk_email = "servicedesk@tamucc.edu"

    for n in range(0, len(arg.files)):
        if arg.verbose:
            print("Reading from file %i" % n)
        # Read data from the spreadsheet
        in_file = arg.files[n]
        try:
            data = pd.read_excel(in_file)
        except FileNotFoundError:
            print("%s not Found" % arg.files[n])
            exit(1)

        # Iterate through each row of the spreadsheet
        size = data.shape[0]
        for row in range(0, size):
            # skip rows older than 7 days by default
            if int(data.iloc[row, 10]) > arg.m:
                if arg.verbose:
                    print("Skipping row %i" % row)
                continue

            # Check if student or Faculty
            if data.iloc[row, 5] == "Student":
                template_file = "Student.html"
            else:
                template_file = "Faculty.html"

            # Determine Course Title and acronym
            course_title = data.iloc[row, 8]
            if course_title == "Information Security Awareness":
                course_acronym = "ISA"
            elif course_title == "FERPA":
                course_acronym = "FERPA"
            elif course_title == "Digital Accessibility Awareness - TAMUCC":
                course_title = "Digital Accessibility Awareness"
                course_acronym = "DAA"
            else:
                raise Exception("Unknown Training type on row %s" % row)

            # Read the template into the email body
            template = open(template_file)
            name = fix_name(data.iloc[row, 0])
            body = template.read()
            template.close()

            # Replace placeholders with data from spreadsheet
            body = re.sub("NAME", name, body)
            body = re.sub("REPTR", arg.reporter, body)
            body = re.sub("ASIGN", arg.assignee, body)
            body = re.sub("UID", str(data.iloc[row, 1]), body)
            body = re.sub("ESD", str(data.iloc[row, 2]), body)
            body = re.sub("LPD", str(data.iloc[row, 3]), body)
            body = re.sub("TCD", str(data.iloc[row, 4]), body)
            body = re.sub("SSFD", str(data.iloc[row, 5]), body)
            body = re.sub("EMAIL", str(data.iloc[row, 6]), body)
            body = re.sub("PCID", str(data.iloc[row, 7]), body)
            body = re.sub("CT", str(data.iloc[row, 8]), body)
            body = re.sub("TDD", str(data.iloc[row, 9]), body)
            body = re.sub("DPD", str(data.iloc[row, 10]), body)
            body = re.sub("PDR", str(data.iloc[row, 11]), body)
            body = re.sub("DESC", str(data.iloc[row, 12]), body)
            body = re.sub("ADLOC", str(data.iloc[row, 13]), body)
            body = re.sub("SLOC", str(data.iloc[row, 14]), body)
            body = re.sub("SUPER", str(data.iloc[row, 15]), body)
            body = re.sub("SMAIL", str(data.iloc[row, 16]), body)
            body = re.sub("CRSTITLE", course_title, body)
            body = re.sub("CRSA", course_acronym, body)

            # Prepare Message
            msg = MIMEMultipart()
            body = MIMEText(body, 'html')
            msg.attach(body)
            msg['From'] = ois_email
            msg['Subject'] = course_title + " training for " + name
            msg['To'] = str(data.iloc[row, 6])
            msg['CC'] = "%s; %s" % (str(data.iloc[row, 16]), service_desk_email)
            msg['BCC'] = ois_email

            # Send Email
            s = smtplib.SMTP('smtp.tamucc.edu')
            s.send_message(msg)
            s.quit()

            if not arg.quiet:
                print("Sent %s email to: %s" % (course_acronym, name))

            # Avoid Spam Filter
            sleep(arg.delay)

    print("DONE")


if __name__ == '__main__':
    args = parser.parse_args()
    if args.files:
        send_mail(args)
    else:
        print("No files specified.")
