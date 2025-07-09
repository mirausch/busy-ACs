from ics import Calendar, Event
import requests

url = "https://calendar.google.com/calendar/ical/michellemrausch%40gmail.com/public/basic.ics"
c = Calendar(requests.get(url).text)
e = Event()
e.name = "Bob's meeting"
e.begin = '2025-09-07 00:00:00'
c.events.add(e)
c.events
print(c.events)
