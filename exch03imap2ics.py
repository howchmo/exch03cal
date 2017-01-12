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
date = (datetime.date.today() - datetime.timedelta(1)).strftime("%d-%b-%Y")
result, uids = mail.uid('search', None, '(SENTSINCE {date})'.format(date=date))
try:
	for uid in uids[0].split():
		result, data = mail.uid('fetch', uid, '(RFC822)')
		raw_email = data[0][1]
	
		msg = email.message_from_string(raw_email)
		# print( msg['To'] )
		# print( email.utils.parseaddr(msg['From']) )
		# print( email_message.items() )
		for part in msg.walk():
			if part.get_content_type() == 'text/calendar':
				ics_text = part.get_payload(decode=1)
				importing = Calendar.from_ical(ics_text)
				for event in importing.subcomponents:
					print(event.name)
					if event.name != 'VEVENT':
						continue
					merged_calendar.add_component(event)
finally:
	# Disconnect from the IMAP server
	if mail.state != 'AUTH':
		mail.close()
	mail.logout()

output = open( 'test.ics', 'wt')
try:
	output.write(merged_calendar.to_ical()) #.decode('utf-8'))
finally:
	output.close()
