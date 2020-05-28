import csv
import icalendar
from datetime import datetime
import pytz
import tempfile, os

def open_csv(filename):
    with open(filename, newline='') as f:    
        csvfile = csv.reader(f, delimiter=',', quotechar='|')
        csvfile = list(csvfile)
        print(filename, 'loaded.')
        return csvfile

if __name__ == '__main__':
    # open_csv('washnwait.csv')
    cal = icalendar.Calendar()
    cal.add('prodid', '-//My calendar product//mxm.dk//')
    cal.add('version', '2.0')
    event = icalendar.Event()
    event.add('summary', 'yunk')
    event.add('dtend', datetime(2020,6,1,10,0,0,tzinfo=pytz.utc))
    event.add('dtstamp', datetime(2020,6,1,0,10,0,tzinfo=pytz.utc))
    organizer = icalendar.vCalAddress('MAILTO:dalarson@wpi.edu')
    attendee = icalendar.vCalAddress('MAILTO:dalarson@wpi.edu')
    event.add('organizer', organizer)
    event.add('attendee', attendee)
    cal.add_component(event)
    print(cal.to_ical())
    with open('invite.ics', 'wb') as f:
        f.write(cal.to_ical())

    


    