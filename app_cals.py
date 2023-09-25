import sys
from calendars import OutlookCalendar, GoogleCalendar, week_to_days


WEEKS = [{
    # Weeks when wife is late-shift, I do the pick-up at school
    'weeks': [36, 38, 39, 40, 44, 45, 47, 49, 50],
    'time': '15:30',
    'subject': 'Buscar da escola',
    },{
    # Weeks when she's early, I do the drop-off
    'weeks': [37, 42, 43, 46, 48],
    'time': '07:30',
    'subject': 'Levar pra escola',
    }]


##WEEKS = [{
##    'weeks': [39],
##    'time': '15:30',
##    'subject': 'Buscar da escola',
##    }]

OUTLOOK_CALENDAR = OutlookCalendar()
GOOGLE_CALENDAR = GoogleCalendar()


def clear_all(dry=True):

    for week in WEEKS:
        start_date = week_to_days(min(week['weeks']), time='00:00')[0]
        end_date = week_to_days(min(week['weeks']), time='23:59')[-1]
        subject = week['subject']
        OUTLOOK_CALENDAR.delete_events(start_date, end_date, subject, dry)
        GOOGLE_CALENDAR.delete_events(start_date, end_date, subject, dry)


def add_all(dry=True):

    if not dry:
        print('WARNING')
        print('Several appointments will be created in Outlook?')
        answer = input('Proceed? [yes/*] ')
        if not answer == 'yes':
            print('Cancelled.')
            return

    for week in WEEKS:
        week_nums = week['weeks']
        time = week['time']
        subject = week['subject']
        
        for week_num in week_nums:
            print(f'Week {week_num}')
            days = week_to_days(week_num, time)
            print([f'{d}' for d in days], end='')
            if not dry:
                for day in days:
                    OUTLOOK_CALENDAR.create_event(day, subject)
                    GOOGLE_CALENDAR.create_event(day, subject)
            print()


if __name__ == '__main__':
    try:
        action = sys.argv[1]
    except IndexError:
        print('Missing argument.')
        print(f'Usage: python {sys.argv[0]} [add|delete] doit')
        raise SystemExit
    if action not in ('add', 'delete'):
        print('Error - missing action: "add" or "delete"')
        raise SystemExit
    try:
        dry = not sys.argv[2] == "doit"
    except IndexError:
        dry = True

    print(f'Dry run: {dry}')
    if action == 'add':
        print('Adding events to calendars')
        add_all(dry)
    elif action == 'delete':
        print('Deleting events from calendars')
        clear_all(dry)
    print(f'Done with {"dry run." if dry else "real run."}')
