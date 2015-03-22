/**
 * @file: Allows importing calendar events from a specified calendar. Builds on the
 * calendar plugin.
 */

var plugin = new SAplugin('calendarImport');

/**
 * This plugin uses options from other plugins:
 *   calendar: calendar name, eventId, start time, end time, event name, decription and location.
 */

// Date interval for events to be imported.
plugin.options.startDate = new Date(); // Now.
plugin.options.endDate = new Date(plugin.options.startDate.getTime() + 1000 * 60 * 60 * 24 * 365); // One year from the start date.

plugin.dependencies = {
  calendar : {
    version : 1
  },
};

// Declaration of menu entries for this plugin.
plugin.globalActions = {
  calendarImportImport : {
    title : 'Import events from a calendar',
    description : 'Here is some description...'
  },
}

// Menu callbacks.
function calendarImportImport() {
  SA.plugin.calendar.options.createIfNeeded = false;
  SA.executeGlobalAction('calendarImport', 'import');
}

plugin.import = function() {
  var calendar = SA.fetch.calendar(SA.plugins.calendar.options.calendarName, true);
  var sheet = globalOptions.sheet;
  var row = globalOptions.startRow;

  var events = calendar.getEvents(this.options.startDate, this.options.endDate);

  for (var i in events) {
    while (SA.fetch.cell(row, SA.plugins.calendar.options.eventId).getValue() !== '') {
      row++;
    }
    SA.fetch.cell(row, SA.plugins.calendar.options.eventId).setValue(events[i].getId());
    SA.fetch.cell(row, SA.plugins.calendar.options.startTime).setValue(events[i].getStartTime());
    SA.fetch.cell(row, SA.plugins.calendar.options.endTime).setValue(events[i].getEndTime());
    if (SA.plugins.calendar.options.eventName) {
      SA.fetch.cell(row, SA.plugins.calendar.options.eventName).setValue(events[i].getTitle());
    }
    if (SA.plugins.calendar.options.eventLocation) {
      SA.fetch.cell(row, SA.plugins.calendar.options.eventLocation).setValue(events[i].getLocation());
    }
    if (SA.plugins.calendar.options.eventDescription) {
      SA.fetch.cell(row, SA.plugins.calendar.options.eventDescription).setValue(events[i].getDescription());
    }
  }
}
