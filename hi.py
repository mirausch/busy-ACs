from ics import Calendar, Event
import requests
from datetime import datetime
import pytz

calendar_url = "https://outlook.office365.com/owa/calendar/af46911eeef94613a79a8f1f7332574b@uzh.ch/5b2660ac410147e6b50e1cd330712bdd4038683141011928974/calendar.ics"

# Load calendar
c = Calendar(requests.get(calendar_url).text)

# Add your test event
e = Event()
e.name = "Bob's meeting"
e.begin = datetime.fromisoformat('2025-07-09T11:20:00')
e.end = datetime.fromisoformat('2025-07-09T11:50:00')
c.events.add(e)

# Get current time
now = datetime.now()
print(f"Current time: {now}")

# Check if there's currently a meeting happening
is_busy = False
for event in c.events:
    # Convert event times to naive datetime if they're timezone-aware
    event_start = event.begin.datetime
    event_end = event.end.datetime

    # Make sure we're comparing like with like (naive vs aware datetimes)
    if event_start.tzinfo is not None:
        event_start = event_start.replace(tzinfo=None)
    if event_end.tzinfo is not None:
        event_end = event_end.replace(tzinfo=None)

    # Debug: Print the comparison
    print(f"Event: {event.name}")
    print(f"  Start: {event_start}")
    print(f"  End: {event_end}")
    print(f"  Now: {now}")
    print(f"  Is {event_start} <= {now} <= {event_end}? {event_start <= now <= event_end}")

    # Check if current time is within the event time range
    if event_start <= now <= event_end:
        is_busy = True
        print(f"Currently in meeting: {event.name}")
        break

if is_busy:
    print('Busy')
else:
    print('Free')

# Optional: Print all events for debugging
print("\nAll events:")
for event in c.events:
    print(f"- {event.name}: {event.begin.datetime} to {event.end.datetime}")