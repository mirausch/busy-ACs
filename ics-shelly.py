from ics import Calendar
import requests
from datetime import datetime, timedelta
import time

calendar_url = "https://outlook.office365.com/owa/calendar/af46911eeef94613a79a8f1f7332574b@uzh.ch/abffc4c067fe4b328df80c853ad641305800255049284926518/calendar.ics"

# SHELLY PLUGS (andere noch adden)
shelly_OG3_VT = "192.168.50.209"
# shelly_Beletage_VT = "192.168.XX.XXX"

# test: requests.get(f"http://{shelly_OG3_VT}/relay/0?turn=on")

def control_shelly_plug(ip_address, turn_on=True):
    """Control Shelly plug via http req. - turn on/off
    Args:
    ip_address (str): IP address of the Shelly plug
    turn_on (bool): True to turn on, False to turn off"""
    try:
        if turn_on:
            url = f"http://{ip_address}/relay/0?turn=on"
            action = "ON"
        else:
            url = f"http://{ip_address}/relay/0?turn=off"
            action = "OFF"
        
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            print(f"Shelly plug {action} successfully")
        else:
            print(f"Failed to turn {action} Shelly plug: {response.status_code}")
    except Exception as e:
        print(f"Error controlling Shelly plug: {e}")

def check_calendar():
    # current time (adjusted for timezone UTC+2)
    now = datetime.utcnow() + timedelta(hours=2)

    # fetch and parse calendar
    c = Calendar(requests.get(calendar_url, headers={'Cache-Control': 'no-cache'}).text)

    is_busy = False
    next_sitzung = None

    for event in c.events:
        event_start_utc = event.begin.datetime.astimezone(tz=None).replace(tzinfo=None) - timedelta(hours=2)
        event_end_utc = event.end.datetime.astimezone(tz=None).replace(tzinfo=None) - timedelta(hours=2)

        if event_start_utc <= now <= event_end_utc:
            is_busy = True
            print(f"Status: BESETZT (bis {event_end_utc:%H:%M})")
            control_shelly_plug(shelly_OG3_VT, turn_on=True)
            break

        if event_start_utc > now and (next_sitzung is None or event_start_utc < next_sitzung):
            next_sitzung = event_start_utc

    if not is_busy:
        print("Status: FREI")
        control_shelly_plug(shelly_OG3_VT, turn_on=False)

    return next_sitzung

# run calendar check every 5 min CHANGE TO 5 MIN
while True:
    next_sitzung = check_calendar()

    last_update = datetime.utcnow() + timedelta(hours=2)
    print(f"Letzte Aktualisierung: {last_update:%H:%M}.")

    if next_sitzung:
        print(f"Nächste Sitzung: {next_sitzung:%H:%M}.")
    print("-----")

    time.sleep(30)
