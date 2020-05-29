import csv
import icalendar
import datetime as dt
import pytz
import tempfile, os
import smtplib
import random
import uuid
import getpass
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.encoders import encode_base64

def open_csv(filename):
    f = open(filename, newline='')
    csvfile = csv.reader(f, delimiter=',', quotechar='|')
    print(filename, 'loaded.')
    return csvfile

def getUniqueId():
    return uuid.uuid4() #Generate UUID according to RFC 4122

def pull_csv():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)

    file_id = '1zmuryV61xolxB4ixMt4lfvEf_ER3uW7sIJoYRFLizoI'
    f = drive.CreateFile({"id": file_id})
    print('Downloading file %s from Google Drive' % f['title'])
    filename = 'washnwait.csv'
    f.GetContentFile(filename, mimetype='text/csv')  # Save Drive file as a local file
    return filename

def parse_csv(f): #expects a csv file
    data = list(f)

def prompt_creds():
    username = input("Enter outlook username:")
    password = getpass.getpass()
    return username, password

def sendAppointment(attendees, dtstart):
    #Login to SMTP server
    s = smtplib.SMTP('smtp.office365.com', 587)
    s.connect("smtp.office365.com", 587)
    s.ehlo()
    s.starttls()
    s.ehlo()
    username, password = prompt_creds()
    s.login(username, password)


    # Timezone to use for our dates - change as needed
    tz = pytz.timezone("US/Eastern")
    reminderHours = 1
    start = tz.localize(dtstart)
    cal = icalendar.Calendar()
    cal.add('prodid', '-//My calendar application//example.com//')
    cal.add('version', '2.0')
    cal.add('method', "REQUEST")
    event = icalendar.Event()
    # event.add('attendee', 'enschneider@wpi.edu') 
    # event.add('organizer', "dalarson@wpi.edu")
    event.add('status', "confirmed")
    event.add('category', "Event")
    event.add('summary', subj)
    event.add('description', description)
    event.add('location', "SigEp Kitchen")
    event.add('dtstart', start)
    event.add('dtend', tz.localize(dt.datetime.combine(dt.date.today(), dt.time(startHour + 1, 0, 0))))
    event.add('dtstamp', tz.localize(dt.datetime.combine(dt.date.today(), dt.time(6, 0, 0))))
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
    msg["Subject"] = subj
    msg["From"] = "dalarson@wpi.edu"
    msg['To'] = ", ".join(['enschneider@wpi.edu'])
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
    
    try:
        # s.sendmail(msg["From"], msg["To"], msg.as_string()) 
        pass
    except smtplib.SMTPDataError:
        print("SMTP failed to send the invite. SMTPDataError Thrown. You cannot send a calendar invite to the account you have logged in with.") 
    except Exception:
        print("SMTP failed to send the invite.")
    else: 
        print("Outlook invitation successfully sent to", msg['To'])
    s.quit()


def main():
    # filename = pull_csv()
    filename = 'washnwait.csv'
    f = open_csv(filename)
    parse_csv(f)
    # sendAppointment('fuck', dt.date.today())
    

if __name__ == '__main__':
    main()

    


    