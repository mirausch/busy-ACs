from ics import Calendar, Event
import requests
from datetime import datetime, timedelta
import time

# current time (adkjusted for timezone)
now = datetime.now() + timedelta(hours=2)

calendar_url = "https://outlook.office365.com/owa/calendar/af46911eeef94613a79a8f1f7332574b@uzh.ch/5b2660ac410147e6b50e1cd330712bdd4038683141011928974/calendar.ics"

def check_calendar():
  c = Calendar(requests.get(calendar_url).text)
  e = Event()
  e.name = "Michelle Rausch"
  e.begin = datetime.fromisoformat('2025-07-09T11:55:00')
  e.end = datetime.fromisoformat('2025-07-09T12:15:00')

  is_busy = False
  for event in c.events:
    event_start = event.begin.datetime
    event_end = event.end.datetime
  
    if event_start.tzinfo is not None:
      event_start = event_start.replace(tzinfo=None)
    if event_end.tzinfo is not None:
      event_end = event_end.replace(tzinfo=None)
    
    if event_start <= now <= event_end:
      is_busy = True
      print(f"Status: {event.name}")
      print(f"Besetzt bis: {event_end:%H:%M}")
      break

# run calendar check every 5 min
while True:
  check_calendar()
  print("Waiting 5 minutes...")
  time.sleep(300)