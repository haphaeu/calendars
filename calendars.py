"""
Create calendar reminders in Outlook to pick-up kids from school.

Update the global variables `WEEKS` and `TIME`, with the week numbers
you have to do the pick-up from school, and the time the event should
start. Events will be created for those weeks, from Monday to Friday.

By default the script will do a dry-run.

If you need to delete appointments created by accident, see below the 
method `delete_events`.

# TODO

 - [x] Create a method to delete events
 - [x] Add morning drop off events
 - [x] Create a method to add events to Google Calendar
 - [x] Delete events from Good Calendar
 - [x] Split lib (purpose agnostic) and app (pick-up, drop-off)
 - [x] Create an abstract parent class (interface) for Calendar 

"""
import pytz
import os.path
from abc import ABC, abstractmethod 
from datetime import datetime as dt
from datetime import timedelta
from win32com.client import Dispatch


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


TZ = pytz.timezone('Europe/oslo')


def week_to_days(week_num, time='00:00', year=2023):
    """Given a week number, returns a list with `datetime` objects
    representing the days from Monday to Friday at that week.
    
    These dates are given at the input `time`.
    """
    strtime = f'{year:d}-W{week_num:d}-1 {time}'
    return [
        dt.strptime(strtime, '%Y-W%W-%w %H:%M')
        + timedelta(days=offset)
        for offset in range(5)
    ]


class Calendar(ABC):
    """
    Abstract class for a calendar interface.
    """
    def __init__(self):
        pass
        
    @abstractmethod
    def create_event(self, 
                     start: dt, 
                     subject: str, 
                     duration: int,
                     ):
        """Create a calendar event."""
        raise NotImplementedError
    
    @abstractmethod
    def delete_events(self, 
                      from_date: dt, 
                      to_date: dt, 
                      subject: str, 
                      dry: bool,
                      ):
        """Delete calendar events within a date range."""
        raise NotImplementedError

class OutlookCalendar(Calendar):
    """
    API ref for AppointmentItems:
    (Dispatch("Outlook.Application").CreateItem(1)):
    https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem
    """

    def __init__(self):
        super().__init__()
        self.outlook = Dispatch("Outlook.Application")

    def create_event(self, start, subject, duration=60):
        """Create an appointment in Outlook calendar.
        
        start: datetime object
        subject: title of the appointment
        """
        appt = self.outlook.CreateItem(1) # AppointmentItem
        appt.Start = f'{start}'
        appt.Subject = subject
        appt.Duration = duration
        appt.ReminderMinutesBeforeStart = 30
        appt.BusyStatus = 3 ## Out of the Office
        appt.Categories = 'Private'
        appt.Save()
        #appt.Send()  #??

    def delete_events(self, from_date, to_date, subject, dry=True):

        cal = self.outlook.GetNamespace("MAPI").GetDefaultFolder(9) 
        appointments = cal.Items
        
        appts_2b_deleted = []
        for appt in appointments:
            if (
                appt.Subject == subject and 
                appt.Start >= TZ.localize(from_date) and
                appt.Start <= TZ.localize(to_date)
            ):
                appts_2b_deleted.append(appt)
                
        if not appts_2b_deleted:
            print('No events found in Outlook.')
            return
            
        print(f'Found {len(appts_2b_deleted)} Outlook appointments '
              f'to be deleted:')
        for appt in appts_2b_deleted:
            print(appt.Subject, appt.Start)
            
        if dry:
                print('Dry run - not proceeding.')
                return

        answer = input('Proceed? ')
        if not answer == 'yes':
            print('Cancelled.')
            return

        for appt in appts_2b_deleted:
            appt.Delete()
        print('Done.')


class GoogleCalendar(Calendar):
    """
    API reference:
    https://developers.google.com/calendar/api/v3/reference/events
    """

    def __init__(self, calendarId='primary'):
        """To switch to other calendars, change `calendarId` to:
            Primary: 'primary'
            Treino: 'sjor2ciqns5rju6bd7u6tlis8k@group.calendar.google.com'
            Escalada: "ea85e0d13649f36af78eea1c526911bc1f375a432b876db9e50eada3fed2f765@group.calendar.google.com"
            BISS: "vsm0v2gr6ln7v0uqrubhje302k@group.calendar.google.com"
        """
        super().__init__()
        # If modifying these scopes, delete the file token.json.
        self.SCOPES = ['https://www.googleapis.com/auth/calendar']
        self.calendarId = calendarId
        self.creds = self._get_credentials()

    def _get_credentials(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, 
        # and is created automatically when the authorization flow completes
        # for the first time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', 
                                                          self.SCOPES)
        # If there are no (valid) creds available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        return creds

    def format_event_body(self, start, subject, duration):
        """Return a calendar event body.
        
        start: datetime object
        subject: string
        """

        end = start + timedelta(minutes=duration)
        
        body = {
          'summary': f'{subject}',
          'location': '',
          'description': '',
          'start': {
            'dateTime': f'{start.astimezone().isoformat()}',
            'timeZone': 'Europe/Oslo',
          },
          'end': {
            'dateTime': f'{end.astimezone().isoformat()}',
            'timeZone': 'Europe/Oslo',
          },
        }

        return body

    def create_event(self, start, subject, duration=60):
        """Add an event to the calendar.
        
        start: datetime object
        subject: string
        """
        try:
            service = build('calendar', 'v3', credentials=self.creds)
            body = self.format_event_body(start, subject, duration)
            event = service.events().insert(calendarId=self.calendarId, 
                                            body=body).execute()
            print('Event created: %s' % (event.get('htmlLink')))

        except HttpError as error:
            print('An error occurred: %s' % error)

    def delete_events(self, from_date, to_date, subject, dry=True):
        """Delete all events from `from_date` containing `subject`"""

        from_date = from_date.astimezone().isoformat()
        to_date = to_date.astimezone().isoformat()

        try:
            service = build('calendar', 'v3', credentials=self.creds)

            # https://developers.google.com/calendar/api/v3/reference/events/list
            events_result = service.events().list(
                calendarId=self.calendarId, 
                timeMin=from_date, timeMax=to_date,
                maxResults=999, singleEvents=True,
                orderBy='startTime').execute()

            events = events_result.get('items', [])

            if not events:
                print('No upcoming events found in Google Calendar.')
                return

            events_2b_deleted = []
            for event in events:
                if event['summary'] == subject:
                    events_2b_deleted.append(event)
            
            if not events_2b_deleted:
                print('No events to delete found in Google calendar.')
                return 

            print(f'Found {len(events_2b_deleted)} Google events '
                  f'to be deleted:')
            for event in events_2b_deleted:
                start = event['start'].get('dateTime', event['start'].get('date'))
                print(start, event['id'], event['summary'])

            if dry:
                print('Dry run - not proceeding.')
                return

            answer = input('Proceed and delete those events? ')
            if answer == 'yes':
                for event in events_2b_deleted:
                    eventId = event['id']
                    service.events().delete(calendarId=self.calendarId, 
                                            eventId=eventId).execute()

        except HttpError as error:
            print('An error occurred: %s' % error)
