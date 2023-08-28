# Outlook and Google Calendars

Create and remove events to Outlook and Google calendars.

## Library

The module `calendars.py` has the `OutlookCalendar` and 
`GoogleCalendar` classes that manage the creation and removal of 
entries, given dates and the events names.

The events are hard-coded to be simple, with only start date and time
and title as parameters. Other properties are hard-coded.

Outlook calendar is handled locally using a locally installed Outlook 
via the `win32com` windows library.

Google Calendar is handled via the Google API, which must be properly
setup and the credentials saved to a local file named 
`credentials.json`. During first use, you'll be requested to log in,
which will create a token file `token.json`. To get started with Google
API, refer to the [Google Python
API](https://developers.google.com/calendar/api/quickstart/python).

## App

The app `app_cals.py` is a simple example tool that creates calendar
events in both Outlook and Google, according to a school pick-up and
drop-off weekly schedule.

The main input are week numbers, so events will be created from Monday 
to Fridays of those weeks, at the input `time` parameter titled after the
`subject` parameter if the input dictionary, which is simply a global list
in the app.



