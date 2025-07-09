from ics import Calendar, Event
import requests

url = "https://outlook.office365.com/owa/calendar/af46911eeef94613a79a8f1f7332574b@uzh.ch/5b2660ac410147e6b50e1cd330712bdd4038683141011928974/calendar.ics"
c = Calendar(requests.get(url).text)
e = Event()
e.name = "Bob's meeting"
e.begin = '2025-09-07 00:00:00'
c.events.add(e)
c.events

g = Event()
g.name = "Meeting 2"
g.begin = '2025-09-08 00:00:00'

print(e)
