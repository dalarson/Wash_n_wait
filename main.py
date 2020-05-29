import csv
import icalendar
import datetime as dt
import pytz
import tempfile, os
import smtplib
import random
import uuid
import getpass
import json
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.encoders import encode_base64

def open_csv(filename):
    f = open(filename, newline='')
    csvfile = csv.reader(f, delimiter=',', quotechar='|')
    print("Local file", filename, 'loaded.')
    return csvfile

def getUniqueId():
    return uuid.uuid4() #Generate UUID according to RFC 4122

def pull_csv():
    print("Attempting to sign into google drive...")
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)

    file_id = '1zmuryV61xolxB4ixMt4lfvEf_ER3uW7sIJoYRFLizoI'
    f = drive.CreateFile({"id": file_id})
    print('Downloading file %s from Google Drive' % f['title'])
    filename = 'washnwait.csv'
    f.GetContentFile(filename, mimetype='text/csv')  # Save Drive file as a local file
    print(filename, "has been downloaded from google drive.")
    return filename

def get_emails(names, emails_dict):
    emails = []
    for name in names:
        try:
            emails.append(emails_dict[name])
        except Exception:
            continue
    return emails


def parse_csv(f, emails_dict): #expects a csv file
    data = list(f)
    emails =  get_emails([data[1][1].lower(), data[2][1].lower()], emails_dict)
    sendAppointment(emails, next_monday(dt.datetime.combine(dt.datetime.now(), dt.time(11))))
    # print(data[1][2], data[2][2]) #11:00 on tuesday
    # print(data[3][1], data[4][1]) # 2:00 on Monday
    # print(data[12][1]) #before 5 on Saturday
    # print(data[12][2]) # before 5 on Sunday

def prompt_creds():
    username = input("Enter WPI email:")
    password = getpass.getpass()
    return username, password

def next_monday(d):
    days_ahead = 0 - d.weekday()
    if days_ahead <= 0: # Target day already happened this week
        days_ahead += 7
    return d + dt.timedelta(days_ahead)

def read_json(filename):
    f = open(filename)
    data = json.load(f)
    emails = data['brothers']
    return emails

def sendAppointment(attendees, dtstart):
    #Login to SMTP server
    s = smtplib.SMTP('smtp.office365.com', 587)
    s.connect("smtp.office365.com", 587)
    s.ehlo()
    s.starttls()
    s.ehlo()
    username, password = prompt_creds()
    try:
        s.login(username, password)
    except smtplib.SMTPAuthenticationError:
        print("Invalid credentials. Please try again.")
        quit()

    # Timezone to use for our dates - change as needed
    tz = pytz.timezone("US/Eastern")
    reminderHours = 1
    description = "wash dishes"
    start = tz.localize(dtstart)
    cal = icalendar.Calendar()
    cal.add('prodid', '-//My calendar application//example.com//')
    cal.add('version', '2.0')
    cal.add('method', "REQUEST")
    event = icalendar.Event()
    for attendee in attendees:
        event.add('attendee', attendee)
    event.add('organizer', username)
    event.add('status', "confirmed")
    event.add('category', "Event")
    event.add('summary', "wash n wait")
    event.add('description', description)
    event.add('location', "SigEp Kitchen")
    event.add('dtstart', start)
    event.add('dtend', start + dt.timedelta(hours=1))
    event.add('dtstamp', tz.localize(dt.datetime.now()))
    event['uid'] = getUniqueId() # Generate some unique ID
    event.add('priority', 5)
    event.add('sequence', 1)
    event.add('created', tz.localize(dt.datetime.now()))     
    alarm = icalendar.Alarm()
    alarm.add("action", "DISPLAY")
    alarm.add('description', "Reminder")
    alarm.add("trigger", dt.timedelta(hours=-reminderHours))
    # The only way to convince Outlook to do it correctly
    alarm.add("TRIGGER;RELATED=START", "-PT{0}H".format(reminderHours))
    event.add_component(alarm)
    cal.add_component(event)
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Wash 'n' Wait"
    msg["From"] = username
    msg['To'] = " ".join(attendees)
    msg["Content-class"] = "urn:content-classes:calendarmessage"
    msg.attach(MIMEText(description))
    filename = "invite.ics"
    part = MIMEBase('text', "calendar", method="REQUEST", name=filename)
    part.set_payload( cal.to_ical() )
    encode_base64(part)
    part.add_header('Content-Description', filename)
    part.add_header("Content-class", "urn:content-classes:calendarmessage")
    part.add_header("Filename", filename)
    part.add_header("Path", filename)
    msg.attach(part)
    send = input("Are you sure you want to send invitations? (type \"yes\" to send) ")
    
    if send.lower() != "yes":
        quit()
    try:
        s.sendmail(msg["From"], msg["To"], msg.as_string()) 
        # print("would send mail HERE")
    except smtplib.SMTPDataError:
        print("SMTP failed to send the invite. SMTPDataError Thrown. You cannot send a calendar invite to the account you have logged in with.") 
    except Exception as e:
        print("SMTP failed to send the invite:", e)
    else: 
        print("Outlook invitation successfully sent to", msg['To'])
    s.quit()


def main():
    filename = 'washnwait.csv'
    pull = input("Do you want to pull wash n wait schedule from google drive? (type \"yes\" to pull down) ")
    if pull.lower() == "yes":
        filename = pull_csv()
    f = open_csv(filename)
    emails_dict = read_json('emails.json')
    parse_csv(f, emails_dict)

    # sendAppointment('fuck', dt.date.today())
    

if __name__ == '__main__':
    main()

    


    