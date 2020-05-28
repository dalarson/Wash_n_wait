import csv
import icalendar
import datetime as dt
import pytz
import tempfile, os
import smtplib
import random
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.encoders import encode_base64

def open_csv(filename):
    with open(filename, newline='') as f:    
        csvfile = csv.reader(f, delimiter=',', quotechar='|')
        csvfile = list(csvfile)
        print(filename, 'loaded.')
        return csvfile

def getUniqueId():
    return 'UID:5FC53010-1267-4F8E-BC28-1D7AE55A7C99'

def sendAppointment(subj, description):
    # Timezone to use for our dates - change as needed
    tz = pytz.timezone("US/Eastern")
    reminderHours = 1
    startHour = 19
    start = tz.localize(dt.datetime.combine(dt.date.today(), dt.time(startHour, 0, 0)))
    cal = icalendar.Calendar()
    cal.add('prodid', '-//My calendar application//example.com//')
    cal.add('version', '2.0')
    cal.add('method', "REQUEST")
    event = icalendar.Event()
    event.add('attendee', 'dalarson@wpi.edu')
    event.add('organizer', "dalarson@wpi.edu")
    event.add('status', "confirmed")
    event.add('category', "Event")
    event.add('summary', subj)
    event.add('description', description)
    event.add('location', "SigEp Kitchen")
    event.add('dtstart', start)
    event.add('dtend', tz.localize(dt.datetime.combine(dt.date.today(), dt.time(startHour + 1, 0, 0))))
    event.add('dtstamp', tz.localize(dt.datetime.combine(dt.date.today(), dt.time(6, 0, 0))))
    # event['uid'] = getUniqueId() # Generate some unique ID
    event.add('priority', 5)
    event.add('sequence', 1)
    event.add('created', tz.localize(dt.datetime.now()))     
    alarm = icalendar.Alarm()
    alarm.add("action", "DISPLAY")
    alarm.add('description', "Reminder")
    #alarm.add("trigger", dt.timedelta(hours=-reminderHours))
    # The only way to convince Outlook to do it correctly
    alarm.add("TRIGGER;RELATED=START", "-PT{0}H".format(reminderHours))
    event.add_component(alarm)
    cal.add_component(event)
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subj
    msg["From"] = "dalarson@wpi.edu"
    msg['To'] = ", ".join(['dalarson@wpi.edu'])
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
    s = smtplib.SMTP('smtp.office365.com', 587)
    s.connect("smtp.office365.com", 587)
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login('dalarson@wpi.edu', 'Chgt44u44')
    s.sendmail(msg["From"], ['dalarson@wpi.edu'], msg.as_string())
    s.quit()

if __name__ == '__main__':
    sendAppointment("fuck you", "bitch")

    


    