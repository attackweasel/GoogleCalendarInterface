import httplib2
import datetime

from apiclient.discovery import build # Google API

from oauth2client.file import Storage
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run_flow

from rfc3339 import rfc3339

def Today():
    """Return the datetime.date object for today's date."""
    return datetime.date.today()

def Tomorrow():
    """Return the datetime.date object for tomorrow's date."""
    return datetime.date.fromordinal(Today().toordinal()+1)

class ObjectFromDict(object):
    """Form an object from a dict.
    The dict's key/values are converted to attributes of the object.
    Nested dicts are also converted to ObjectFromDict objects.

    Why? The Python Google Calendar API uses dicts as parameters. The
    point of this GoogleCalendarInterface is to provide an object-oriented
    interface to Google Calendars.

    Calendar and Event objects inherit from ObjectFromDict.
    """
    
    def __init__(self, d={}):
        """Form an object from a dict.

        Keyword arguments:
        d -- the dict to convert to an object
        """
        
        d = self.ObjectifyDicts(d)
        self.__dict__.update(d)
        
    def ObjectifyDicts(self,d):
        """Transform all dicts in d into ObjectFromDicts. Leaves
        non-dict objects untouched.

        Keyword arguments:
        d -- the dict to check for nested dicts
        """
        
        for k,v in d.items():
            try:
                v.keys()
                d[k] = ObjectFromDict(d[k])
            except:
                pass
        return d

    def attrs(self):
        """Return the attributes of this ObjectFromDict.
        Better than dir(self), for our purposes, because
        it ignores built-in magic methods.
        """
        return self.__dict__.keys()
    
    def ToDict(self, _obj='__none__'):
        """Return self as a dict, nested if necessary.

        Keyword arguments:
        _obj -- ObjectFromDict to convert to dict. Used internally for processing nested objects.
        """
        
        if _obj == '__none__': _obj = self

        try:
            _obj.attrs()
        except:
            if isinstance(_obj, CalendarInterface):
                return None
            else:
                return _obj
        
        d = dict()

        for attr in _obj.attrs():
            if not callable(_obj.__dict__[attr]):
                item = self.ToDict(_obj.__dict__[attr])
                if item != None:
                    d[attr] = self.ToDict(item)

        return d

class CalendarItemsList(list):
    """Form a list of calendar items.
    Base class for CalendarList and EventList.
    Ignores list items with no 'summary' attribute.
    """

    def __init__(self, *args):
        list.__init__(self, *args)

    def __setitem__(self, index, item):
        """Set item only if it has a 'summary' attribute."""
        if hasattr(item, 'summary'):
            list.__setitem__(self, index, item)

    def append(self, item):
        """Append item only if it has a 'summary' attribute."""        
        if hasattr(item, 'summary'):
            list.append(self, item)

    def Summaries(self):
        """Return list of summaries for all items in list."""
        return [item.summary for item in self]

    def Ids(self):
        """Return list of IDs for all items in list."""
        return [item.id for item in self]

    def SummariesAndIds(self):
        """Return list of "summary, ID" tuples for all items in list."""
        return [(item.summary, item.id) for item in self]

class CalendarList(CalendarItemsList):
    """Forms a list of Calendar objects.

    Provides Names() and NamesAndIds() to as more intelligible
    alternatives to Summaries() and SummariesAndIds().
    """
    def __init__(self, *args):
        CalendarItemsList.__init__(self, *args)
        self.Names = self.Summaries
        self.NamesAndIds = self.SummariesAndIds
        
class EventList(CalendarItemsList):
    """Forms a list of Event objects."""
    def __init__(self, *args):
        CalendarItemsList.__init__(self, *args)
    
        
class Calendar(ObjectFromDict):
    """Forms a Calendar object from a dict."""
    
    def __init__(self, d={}, interface=None):
        """Initialize Calendar object.

        Keyword arguments:
        d -- dict to convert to Calendar
        interface -- CalendarInterface object containing this Calendar object
        """
        
        self.interface = interface
        ObjectFromDict.__init__(self, d)

    def __repr__(self):
        if hasattr(self, 'summary'):
            return 'Calendar<' + self.summary[:15] + ('','...')[len(self.summary)>15] + '>'
        return ObjectFromDict.__repr__(self)

    def Events(self, timeMin = 'datetime.datetime', timeMax = 'datetime.datetime', **options):
        """Return EventList populated with this Calendar's events.

        To include a range of dates, make sure to set the end time for the
        end of the day.
        
        Keyword arguments:
        timeMin -- datetime.datetime object for start date and time
        timeMax -- datetime.datetime object for end date and time

        Additional Google Calendar API keyword options may be specified.
        """
        eventList = EventList()

        options['calendarId'] = self.id
        
        if timeMin != 'datetime.datetime':
            options['timeMin'] = rfc3339(timeMin)
        if timeMax != 'datetime.datetime':
            options['timeMax'] = rfc3339(timeMax)
        
        events = self.interface.service.events().list(**options).execute()
        
        while True:
            for event in events.get('items',[]):
                eventList.append(Event(event, self))
            page_token = events.get('nextPageToken')
            if page_token:
                events = self.interface.service.events().list(pageToken=page_token,
                                                              **options).execute()
            else:
                break
            
        return eventList

    def EventsForDate(self, date = 'datetime.date', **options):
        """Return EventList populated with this Calendar's events for this date.

        Keyword arguments:
        date -- datetime.date object for date

        Additional Google Calendar API keyword options may be specified.
        """

        if date == 'datetime.date':
            date=Today()
        
        return self.Events(date, date + datetime.timedelta(days = 1), **options)    

    def GetEventById(self, eventId, **options):
        """Return Event object given its ID.

        Keyword arguments:
        date -- datetime.datetime object for date

        
        Additional Google Calendar API keyword options may be specified.
        """
        return Event(self.interface.service.events().get(calendarId=self.id,
                                                         eventId=eventId,
                                                         **options).execute(), self)

    def CreateEvent(self, summary, start=None, end=None, allDay=True, **options):
        """Create an Event in this calendar and return Event object.

        Keyword arguments:
        summary -- something like "Meeting with Bob"
        start -- datetime.datetime object for start date (and time if allDay is False).
        end -- datetime.datetime object for send date (and time if allDay is False).
        allDay -- boolean - all day event?
        
        Additional Google Calendar API keyword options may be specified.
        """
        
        if not start:
            start = Today()
            end = Tomorrow()
        if allDay:
            start = {'date' : start.strftime('%Y-%m-%d')}
            if end:
                end = {'date' : end.strftime('%Y-%m-%d')}
            else:
                end = start
        else:
            start = {'dateTime' : rfc3339(start)}
            end = {'dateTime' : rfc3339(end)}
        event = {
            'summary': summary,
            'start': start,
            'end': end,
            }
        event.update(options)

        return Event(self.interface.service.events().insert(calendarId='primary',
                                                            body=event).execute(),
                     self)
        
    def QuickAddEvent(self, event):
        """Using a descriptive phrase (such as "Meet with Bob tomorrow at 7pm"),
        create an Event in this calendar and return Event object.

        Keyword arguments:
        event -- something like "Meet with Bob tomorrow at 7pm"
        """

        return Event(self.interface.service.events().quickAdd(
            calendarId=self.id,
            text=event).execute(), self)

    def DeleteEvent(self, event):
        """Delete an event from the current calendar.

        Keyword arguments:
        event -- Event object to delete from calendar.
        """

        return self.interface.service.events().delete(calendarId=self.id,
                                               eventId=event.id).execute()

    def Update(self, **options):
        '''Apply changes already made to Calendar attributes.
        Also, supply additional changes as keyword arguments.


        '''
        self.__dict__.update(options)
        return self.interface.service.calendars().update(calendarId=self.id,
                                                         body=self.ToDict()).execute()


class Event(ObjectFromDict): 
    """Forms an Event object from a dict."""
    
    def __init__(self, d={}, calendar=None):
        """Initialize event object.

        Keyword arguments:
        d -- dict to convert to Calendar
        calendar -- Calendar object containing this Event object
        """
        self.calendar = calendar
        ObjectFromDict.__init__(self, d)

    def __repr__(self):
        if hasattr(self, 'summary'):
            return 'Event<' + self.summary[:15] + ('','...')[len(self.summary)>15] + '>'
        return ObjectFromDict.__repr__(self)

    def Update(self, **options):
        '''Apply changes already made to Event attributes.
        Also, supply additional changes as keyword arguments.


        '''
        self.__dict__.update(options)
        self.calendar.interface.service.events().update(calendarId=self.calendar.id,
                                                        eventId=self.id,
                                                        body=self.ToDict()).execute()

    def IsAllDay(self):
        return hasattr(self.start, 'date') and not hasattr(self.start, 'datetime')
    
    def IsRecurring(self):
        return hasattr(self, 'recurrence') and bool(self.recurrence)
    
class CalendarInterface:

    def __init__(self,
                 client_id,
                 client_secret,
                 user_agent,
                 user='Default'):

        # Set up a Flow object to be used if we need to authenticate. This
        # sample uses OAuth 2.0, and we set up the OAuth2WebServerFlow with
        # the information it needs to authenticate. Note that it is called
        # the Web Server Flow, but it can also handle the flow for native
        # applications
        # The client_id and client_secret are copied from the API Access tab on
        # the Google APIs Console
        FLOW = OAuth2WebServerFlow(
            client_id=client_id,
            client_secret=client_secret,
            scope='https://www.googleapis.com/auth/calendar',
            user_agent=user_agent,
            )

        # If the Credentials don't exist or are invalid, run through the native client
        # flow. The Storage object will ensure that if successful the good
        # Credentials will get written back to a file.
        storage = Storage('CalendarUser-'+user+'.dat')
        self.credentials = storage.get()
        if self.credentials is None or self.credentials.invalid == True:
            self.credentials = run_flow(FLOW, storage)

        # Create an httplib2.Http object to handle our HTTP requests and authorize it
        # with our good Credentials.
        self.http = httplib2.Http()
        self.http = self.credentials.authorize(self.http)

        # Build a service object for interacting with the API. Visit
        # the Google APIs Console to get a developerKey for your own application.
        self.service = build(serviceName='calendar', version='v3', http=self.http,)
        # developerKey='YOUR_DEVELOPER_KEY') # <-- Not necessary?

    def Calendars(self):
        '''Return a list of calendars'''
        calendarList = CalendarList()
        calendar_list = self.service.calendarList().list().execute()
        while True:
            for calendar_list_entry in calendar_list['items']:
                calendarList.append(Calendar(calendar_list_entry, self))
            page_token = calendar_list.get('nextPageToken')
            if page_token:
                calendar_list = service.calendarList().list(pageToken=page_token).execute()
            else:
                break

        return calendarList

    def GetCalendarByName(self, name):
        for calendar in self.Calendars():
            if calendar.summary == name:
                return calendar
        return None

    def GetCalendarById(self, id):
        for calendar in self.Calendars():
            if calendar.id == id:
                return calendar
        return None

    def CreateCalendar(self, summary='New Calendar', timeZone='America/Chicago', **options):
        calendar = {'summary': summary,
                    'timeZone': timeZone}
        calendar.update(options)
        return Calendar(self.service.calendars().insert(body = calendar).execute(), self)

    def DeleteCalendar(self, calendar):
        self.service.calendars().delete(calendarId = calendar.id).execute()
        
if __name__ == '__main__':

    execfile("config.py")

    interface = CalendarInterface(
        client_id,
        client_secret,
        user_agent,
        user)
    calendars = interface.Calendars()
    
    print """
For testing:

    interface = CalendarInterface(
        client_id=%s,
        client_secret=%s,
        user_agent=%s,
        user=%s)
    calendars = interface.Calendars()""" % (
        client_id,
        client_secret,
        user_agent,
        user)


