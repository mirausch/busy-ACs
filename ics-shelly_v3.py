from ics import Calendar
import requests
from datetime import datetime, timedelta
import time
import holidays
import re

# Weather cache to avoid excessive scraping
weather_cache = {
    'date': None,
    'max_temp': None,
    'last_fetch': None
}

calendar_url = "https://outlook.office365.com/owa/calendar/af46911eeef94613a79a8f1f7332574b@uzh.ch/abffc4c067fe4b328df80c853ad641305800255049284926518/calendar.ics"
# zu beachten: if calendar links are refreshed on outlook browser then they need to be updated


# SHELLY PLUGS (andere noch adden)
# SHELLY APP: EXISTIERENDES  ZKG SHELLY KONTO??
shelly_OG3_VT = "192.168.50.209"
# shelly_Beletage_VT = "192.168.50.47"

# test plug only: requests.get(f"http://{shelly_OG3_VT}/relay/0?turn=on")

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

def get_weather_forecast(): # daily cache
    """max temp. forecast: scrape search.ch"""
    today = datetime.now().date()
    now = datetime.now()
    
    # Check if we have valid cached data for today
    if (weather_cache['date'] == today and 
        weather_cache['max_temp'] is not None):
        return weather_cache['max_temp']
    
    # Check if we should fetch weather (only once per day, preferably early morning)
    should_fetch = (
        weather_cache['date'] != today or  # New day
        (weather_cache['last_fetch'] is None and now.hour >= 6) or  # First run after 6 AM
        (weather_cache['max_temp'] is None and now.hour >= 6)  # Failed previous attempt
    )
    
    if not should_fetch:
        # Too early or already fetched today
        if weather_cache['max_temp'] is not None:
            return weather_cache['max_temp']
        else:
            print("Weather: Waiting for 6 AM to fetch forecast")
            return None
    
    try:
        print("Scraping weather forecast from search.ch...")
        url = "https://search.ch/meteo/zuerich"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            content = response.text
            
            # Look for temperature patterns like "zwischen 18 und 23 Grad"
            temp_patterns = [
                r'zwischen \d+ und (\d+) Grad',  # "zwischen 18 und 23 Grad"
                r'bis (\d+) Grad',  # "bis 23 Grad" 
                r'(\d+)\s*°C',  # "23°C"
                r'max[.\s]*(\d+)',  # "max 23" or "max. 23"
            ]
            
            max_temp = None
            for pattern in temp_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                if matches:
                    # Take the highest temperature found
                    temps = [int(match) for match in matches if match.isdigit()]
                    if temps:
                        max_temp = max(temps)
                        break
            
            if max_temp:
                # Cache the result
                weather_cache['date'] = today
                weather_cache['max_temp'] = max_temp
                weather_cache['last_fetch'] = now
                
                print(f"Wettervorhersage: Max {max_temp}°C heute (von search.ch, cached)")
                return max_temp
            else:
                print("Could not parse temperature from search.ch")
                
        else:
            print(f"Weather scraping error: {response.status_code}")
            
    except Exception as e:
        print(f"Error scraping weather: {e}")
    
    # Update cache to avoid repeated failed attempts
    weather_cache['date'] = today
    weather_cache['last_fetch'] = now
    return None

def check_calendar():
    # current time (timezone-aware)
    now = datetime.now()
    
    # Swiss holidays
    swiss_holidays = holidays.Switzerland()
    
    # Check if it's a workday (Mon-Fri, not holiday)
    is_workday = now.weekday() < 5 and now.date() not in swiss_holidays
    
    # Check if it's office hours (07:30-18:00)
    office_start = now.replace(hour=7, minute=30, second=0, microsecond=0)
    office_end = now.replace(hour=18, minute=0, second=0, microsecond=0)
    is_office_hours = office_start <= now <= office_end
    
    if not is_workday or not is_office_hours:
        if not is_workday:
            if now.date() in swiss_holidays:
                print("Status: FEIERTAG - AC OFF")
            elif now.weekday() >= 5:
                print("Status: WOCHENENDE - AC OFF")
        else:
            print("Status: AUSSERHALB BÜROZEITEN - AC OFF")
        control_shelly_plug(shelly_OG3_VT, turn_on=False)
        return None
    
    # Get weather forecast
    max_temp_today = get_weather_forecast()
    is_hot_day = max_temp_today and max_temp_today > 20
    
    # Only during workday office hours: check for meetings
    has_meetings_today = False
    is_meeting_active = False
    next_sitzung = None
    upcoming_meeting_soon = False
    
    try:
        # fetch and parse calendar
        c = Calendar(requests.get(calendar_url, headers={'Cache-Control': 'no-cache'}).text)

        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        today_end = now.replace(hour=23, minute=59, second=59, microsecond=999999)

        for event in c.events:
            event_start_utc = event.begin.datetime.astimezone(tz=None).replace(tzinfo=None) - timedelta(hours=2)
            event_end_utc = event.end.datetime.astimezone(tz=None).replace(tzinfo=None) - timedelta(hours=2)

            # Check if there are any meetings today
            if today_start.replace(tzinfo=None) <= event_start_utc <= today_end.replace(tzinfo=None):
                has_meetings_today = True

            # Check if meeting is currently active
            if event_start_utc <= now.replace(tzinfo=None) <= event_end_utc:
                is_meeting_active = True
                break

            # Check for next meeting
            if event_start_utc > now.replace(tzinfo=None) and (next_sitzung is None or event_start_utc < next_sitzung):
                next_sitzung = event_start_utc
                
                # Check if next meeting is within 10 minutes (only relevant if not hot day)
                time_until_meeting = event_start_utc - now.replace(tzinfo=None)
                if not is_hot_day and time_until_meeting <= timedelta(minutes=10):
                    upcoming_meeting_soon = True

    except Exception as e:
        print(f"Error fetching calendar: {e}")
        # If calendar fails, use conservative approach
        if is_hot_day:
            print("Status: HEISS + KALENDER FEHLER - AC ON (sicherheitshalber)")
            control_shelly_plug(shelly_OG3_VT, turn_on=True)
        else:
            control_shelly_plug(shelly_OG3_VT, turn_on=False)
        return None

    # Decision logic based on weather and meetings:
    if is_hot_day:
        # Hot day (>24°C): AC on all day, off only during active meetings
        if is_meeting_active:
            print("Status: HEISS TAG + MEETING LÄUFT - AC OFF")
            control_shelly_plug(shelly_OG3_VT, turn_on=False)
        else:
            print(f"Status: HEISS TAG ({max_temp_today:.1f}°C) - AC ON")
            control_shelly_plug(shelly_OG3_VT, turn_on=True)
    else:
        # Normal day (<24°C): original logic (AC only when meetings planned)
        if not has_meetings_today:
            print("Status: TEMP < 24 + KEINE SITZUNGEN - AC OFF")
            control_shelly_plug(shelly_OG3_VT, turn_on=False)
        elif is_meeting_active:
            print("Status: NORMALE TEMP + SITZUNG LÄUFT - AC OFF")
            control_shelly_plug(shelly_OG3_VT, turn_on=False)
        elif upcoming_meeting_soon:
            if next_sitzung:
                print(f"Status: NORMALE TEMP + MEETING IN 10 MIN - AC OFF (Meeting um {next_sitzung:%H:%M})")
            else:
                print("Status: NORMALE TEMP + MEETING IN 10 MIN - AC OFF")
            control_shelly_plug(shelly_OG3_VT, turn_on=False)
        else:
            print("Status: NORMALE TEMP + MEETINGS GEPLANT - AC ON")
            control_shelly_plug(shelly_OG3_VT, turn_on=True)

    return next_sitzung

# run calendar check every 5 min -- CHANGE BACK TO 5 MIN
while True:
    next_sitzung = check_calendar()
    last_update = datetime.now()

    if next_sitzung:
        print(f"Nächste Sitzung: {next_sitzung:%H:%M}.")

    print(f"Letzte Aktualisierung: {last_update:%H:%M:%S}.")
    print("-----")

    time.sleep(120)
# FIX TIMEZONE!!!!!
# changes: 
    # should be off when no meeting. mon-fri 8-18 ON, when meeting OFF, when no meetings that day OFF. 
    # what if no meetings and someone makes meeting kfr.
    # datetime.datetime.now is deprecated

# BUFFER TIMES!! 10 min?
# heisser tag from what temp?
