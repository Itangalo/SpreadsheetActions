/**
 * @file: Allows creating and manipulating calendar events en masse, including
 * adding/removing reminders.
 */

var plugin = new SAplugin('calendar');

// The name of the calendar to work with. (If there are more than one with this name, the first will be used.)
plugin.options.calendarName = 'Test calendar';
// This option allows creating a new calendar with the given name, if it does not already exist.
plugin.options.createIfNeeded = true;

// Columns with start time, end time and names of events.
plugin.options.startTime = 6;
plugin.options.endTime = 7;
plugin.options.eventName = 8;
plugin.options.eventDescription = 9;
plugin.options.eventLocation = 10;

// Column to write event IDs to. These are required to allow editing of events.
plugin.options.eventId = 11;

// Used for adding reminders. Allowed values are 'sms', 'email' and 'popup'.
plugin.options.reminderType = 'sms';
// Number of minutes before the event to trigger the reminder. Between 4 and 40320.
plugin.options.minutesBefore = 4;


plugin.dependencies = {
  SA : {
    version : 1
  },
  basics : {
    version : 1
  },
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  calendarCreateOrUpdateEvents : {
    title : 'Create or update events',
    description : 'Here is some description...'
  },
  calendarDeleteEvents : {
    title : 'Delete events',
    description : 'Here is some description...'
  },
  calendarAddReminder : {
    title : 'Add reminders',
    description : 'Here is some description...'
  },
  calendarRemoveReminders : {
    title : 'Remove reminders',
    description : 'Here is some description...'
  },
}

// Menu callbacks.
function calendarCreateOrUpdateEvents() {
  SA.executeBulkAction('calendar', 'createOrUpdateEvents');
}
function calendarDeleteEvents() {
  SA.executeBulkAction('calendar', 'deleteEvents');
}
function calendarAddReminder() {
  SA.executeBulkAction('calendar', 'addReminder');
}
function calendarRemoveReminders() {
  SA.executeBulkAction('calendar', 'removeReminders');
}

plugin.createOrUpdateEvents = function(row) {
  var calendar = SA.fetch.calendar(this.options.calendarName);

  // Verify that the event doesn not already exist.
  var id = SA.fetch.cell(row, this.options.eventId).getValue();
  if (id) {
    var event = calendar.getEventSeriesById(id);
  }
  else {
    var event = calendar.createEvent(SA.fetch.cell(row, this.options.eventName).getValue(), SA.fetch.cell(row, this.options.startTime).getValue(), SA.fetch.cell(row, this.options.endTime).getValue());
    SA.fetch.cell(row, this.options.eventId).setValue(event.getId());
  }
  if (this.options.eventName) {
    event.setTitle(SA.fetch.cell(row, this.options.eventName).getValue());
  }
  if (this.options.eventDescription) {
    event.setDescription(SA.fetch.cell(row, this.options.eventDescription).getValue());
  }
  if (this.options.eventLocation) {
    event.setLocation(SA.fetch.cell(row, this.options.eventLocation).getValue());
  }
}

plugin.deleteEvents = function(row) {
  var calendar = SA.fetch.calendar(this.options.calendarName);
  var event = calendar.getEventSeriesById(SA.fetch.cell(row, this.options.eventId).getValue());
  event.deleteEventSeries();
  SA.fetch.cell(row, this.options.eventId).setValue('');
}

plugin.addReminder = function(row) {
  var calendar = SA.fetch.calendar(this.options.calendarName);
  var event = calendar.getEventSeriesById(SA.fetch.cell(row, this.options.eventId).getValue());
  var callbacks = {
    'sms' : 'addSmsReminder',
    'email' : 'addEmailReminder',
    'popup' : 'addPopupReminder'
  };
  if (!callbacks[this.options.reminderType]) {
    throw 'Type of reminder is not properly specified.';
  }
  event[callbacks[this.options.reminderType]](this.options.minutesBefore);
}

plugin.removeReminders = function(row) {
  var calendar = SA.fetch.calendar(this.options.calendarName);
  var event = calendar.getEventSeriesById(SA.fetch.cell(row, this.options.eventId).getValue());
  event.removeAllReminders();
}

// Fetches or creates an owned calendar with the given name. The calendar is cached locally,
// so it doesn't have to be loaded for each processed row.
plugin.fetchers.calendar = function(name, readOnly) {
  // Check for cached values first.
  if (this.calendar[name]) {
    return this.calendar[name];
  }

  if (readOnly) {
    var c = CalendarApp.getCalendarsByName(name);
  }
  else {
    var c = CalendarApp.getOwnedCalendarsByName(name);
  }

  if (c.length == 0) {
    if (SA.plugins.calendar.options.createIfNeeded && readOnly) {
      c = CalendarApp.createCalendar(name).setTimeZone(Session.getScriptTimeZone());
    }
    throw('Calendar ' + name + ' does not exist.');
  }
  else {
    c = c[0];
  }

  // Cache the result before returning.
  this.calendar[name] = c;
  return c;
}
