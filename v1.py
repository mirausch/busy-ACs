from ics import Calendar, Event
import requests
from datetime import datetime, timedelta
import time
import os
import sys

calendar_url = "https://outlook.office365.com/owa/calendar/af46911eeef94613a79a8f1f7332574b@uzh.ch/5b2660ac410147e6b50e1cd330712bdd4038683141011928974/calendar.ics"

# SHELLY PLUGS // REPLACE W ACTUAL IPs
# shelly_plug_1 = ["192.168.1.100"] 

# def control_shelly_plug(ip_address, turn_on=True):
#    try:
      #if turn_on
          #url = f"http://{ip_address}/relay/0?turn=on"
          #action = "ON"
      #else:
          #url = f"http://{ip_address}/relay/0?turn=off"
          #action = "OFF"


def check_calendar():
  # current time (adkjusted for timezone) NEEDS TO BE IN LOOP!!
  now = datetime.now() + timedelta(hours=2)
  c = Calendar(requests.get(calendar_url, headers={'Cache-Control': 'no-cache'}).text)
  #e = Event()
  #e.name = "Michelle Rausch"
  #e.begin = datetime.fromisoformat('2025-07-09T11:55:00')
  #e.end = datetime.fromisoformat('2025-07-09T12:15:00')

  is_busy = False
  for event in c.events:
    event_start = event.begin.datetime
    event_end = event.end.datetime

      # remove timezone info
    if event_start.tzinfo is not None:
      event_start = event_start.replace(tzinfo=None)
    if event_end.tzinfo is not None:
      event_end = event_end.replace(tzinfo=None)
    
    if event_start <= now <= event_end:
      is_busy = True
      print(f"Status: {event.name}") # outlook event name is always "busy"
      print(f"Besetzt bis: {event_end:%H:%M}")
      # http on-request here? ADD BUFFER TIME (if meeting starts in 5 min, turn on)
          # also need to adjust to length of meeting!! "Besetzt bis"
      break
      
  if not is_busy:
    print("Status: Frei")
    # http off-request here?

# run calendar check every 5 min
while True:
  check_calendar()

  last_update = datetime.now() + timedelta(hours=2)
  print(f"Last update: {last_update:%H:%M}. Next update in 5 min") # add time of next meeting?
  print("-----")
  time.sleep(300)

os.execv(sys.executable, ['python'] + sys.argv)
