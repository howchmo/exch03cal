import imaplib
imaplib._MAXLINE = 40000
import email
from icalendar import Calendar, Event
import json

client_secrets = {}
with open("client_secrets.json") as f:
	client_secrets = json.load(f)

mail = imaplib.IMAP4_SSL(client_secrets.imap_server)
mail.login(client_secrets.login, client_secrets.password)
# mail.list()
# Out: list of "folders" aka labels in gmail.
mail.select("calendar") # connect to inbox.

# Initialize a calendar to hold the merged data
merged_calendar = Calendar()
merged_calendar.add('prodid', '-//exchange//ecsorl.com//')
merged_calendar.add('calscale', 'GREGORIAN')

import datetime
# date = (datetime.date.today() - datetime.timedelta(timedelta)).strftime("%d-%b-%Y")
date = "06-Mar-2015"
result, uids = mail.uid('search', None, '(SENTSINCE {date})'.format(date=date))
try:
	for uid in uids[0].split():
#		print( uid )
		result, data = mail.uid('fetch', uid, '(RFC822)')
		raw_email = data[0][1]
	
		msg = email.message_from_string(raw_email)
		for part in msg.walk():
			if part.get_content_type() == 'text/calendar':
				ics_text = part.get_payload(decode=1)
#				print( "--------------------------------------------" )
#				print( ics_text )
#				print( "--------------------------------------------" )
				try:
					importing = Calendar.from_ical(ics_text)
					for event in importing.subcomponents:
						if event.name != 'VEVENT':
							continue
						merged_calendar.add_component(event)
				except:
					print("WTF!")
finally:
	# Disconnect from the IMAP server
	if mail.state != 'AUTH':
		mail.close()
	mail.logout()

# take this stuff and stick in google calendar
#print( "===== Put This Into Google Calendar =====" )

import gflags
import httplib2

from apiclient.discovery import build
from oauth2client.file import Storage
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run

FLAGS = gflags.FLAGS

FLOW = OAuth2WebServerFlow(
    client_id=client_secrets.installed.client_id,
    client_secret=client_secrets.installed.client_secret,
    scope='https://www.googleapis.com/auth/calendar',
    user_agent='exch2003imap2gcal/1')

storage = Storage('calendar.dat')
credentials = storage.get()
if credentials is None or credentials.invalid == True:
  credentials = run(FLOW, storage)

http = httplib2.Http()
http = credentials.authorize(http)

service = build(serviceName='calendar', version='v3', http=http,
       developerKey=client_secrets.developerKey)


oldiCalUIDs = {}
with open("icaluids.json") as f:
	oldiCalUIDs = json.load(f)
newiCalUIDs = {}

#print(oldiCalUIDs)

for e in merged_calendar.walk('vevent'):
#	print( "" )
#	print( "============================================" )
#	for k in e.keys():
#		print( k )
#	print( "============================================" )
#	print( "" )
	attendees = []
	if( "ATTENDEE" in e.keys() ):
		for a in e["ATTENDEE"]:
			attendees.append({'email':a[7:]})
#	else:
#		print("could not find 'ATTENDEE'")
	event = {
		'iCalUID': e['UID'],
	  'summary': e['SUMMARY'],
	  'start': {
			'dateTime': e['DTSTART'].dt.strftime('%Y-%m-%dT%H:%M:%S.000-05:00'),
			'timeZone':'America/New_York'
		},
	  'end': {
			'dateTime': e['DTEND'].dt.strftime('%Y-%m-%dT%H:%M:%S.000-05:00'),
			'timeZone':'America/New_York'
		},
		'organizer': {"email": e['ORGANIZER'][7:]},
		'attendees': attendees,
		'status': e['STATUS'].lower(),
	  'created': e['CREATED'].dt.strftime('%Y-%m-%dT%H:%M:%S.000-05:00'),
	  'updated': e['LAST-MODIFIED'].dt.strftime('%Y-%m-%dT%H:%M:%S.000-05:00'),
	}
	if( 'LOCATION' in e.keys() ):
		event['location'] =  e['LOCATION']
#	else:
#		print("could not find 'LOCATION'")
	if( 'DESCRIPTION' in e.keys() ):
		event['description'] =  e['DESCRIPTION']
#	else:
#		print("could not find 'DESCRIPTION'")
	if( e.has_key('RRULE') ):
		event['recurrence'] = ['RRULE:'+e['RRULE'].to_ical()]
#	else:
#		print("could not find 'RRULE'")
	if event['iCalUID'] not in oldiCalUIDs:
#		try:
			created_event = service.events().insert(calendarId='primary', body=event).execute()
			newiCalUIDs[event['iCalUID']] = created_event['id']
#		except:
#			print("There was a problem inserting this one !!!")
	else:
#		print( event['iCalUID'] + " has already been inserted." )
		newiCalUIDs[event['iCalUID']] = oldiCalUIDs[event['iCalUID']]
		updated_event = service.events().update(calendarId='primary', eventId=newiCalUIDs[event['iCalUID']], body=event).execute()
		del oldiCalUIDs[event['iCalUID']]

for icaluid, uid in oldiCalUIDs:
	service.events().delete(calendarId='primary', eventId=uid).execute()

#print( "write out all tihe new iCalUIDs to a file" )
with open("icaluids.json","w") as f:
	json.dump(newiCalUIDs,f)
