/**
 * @file: Allows locking sheets and cells in student workbooks. Builds on the studentWorkbook plugin.
 * Current version of the plugin locks the sheet/range for all but the person running the script.
 */

var plugin = new SAplugin('lockStudentWorkbook');

/**
 * This plugin uses options from other plugins:
 *   studentWorkbooks: student workbook link/id column
 */

// The name of the sheet to lock, or the sheet where a range should be locked.
plugin.options.sheetName = 'Betygsunderlag';
// The range to lock, in A1 notation. If false, the selected range will be used.
plugin.options.range = false;

plugin.title = 'Lock student sheets';
plugin.dependencies = {
  studentWorkbooks : {
    version : 1
  },
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  lockStudentWorkbookLockSheet : {
    title : 'Lock sheet in student workbook',
    description : 'Here is some description...'
  },
  lockStudentWorkbookLockRange : {
    title : 'Lock range in student workbook',
    description : 'Here is some description...'
  },
}

// Menu callbacks.
function lockStudentWorkbookLockSheet() {
  SA.executeBulkAction('lockStudentWorkbook', 'lockSheet');
}
function lockStudentWorkbookLockRange() {
  SA.plugins.colorStudentWorkbook.options.color = SA.plugins.colorStudentWorkbook.options.colorOk;
  SA.executeBulkAction('lockStudentWorkbook', 'lockRange');
}

// Most of this code is taken from https://developers.google.com/apps-script/reference/spreadsheet/sheet#protect()
plugin.lockSheet = function(row) {
  // Protect the active sheet, then remove all other users from the list of editors.
  var sheet = SA.fetch.studentSheet(row, this.options.sheetName);
  var protection = sheet.protect().setDescription('Protected by spreadsheet actions');

  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script will throw an exception upon removing the group.
  var actor = Session.getEffectiveUser();
  protection.addEditor(actor);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

// Most of this code is taken from https://developers.google.com/apps-script/reference/spreadsheet/range#protect()
plugin.lockRange = function(row) {
  // Protect given range, then remove all other users from the list of editors.
  var range = SA.fetch.studentSheet(row, this.options.sheetName).getRange(this.options.range || SpreadsheetApp.getActiveRange().getA1Notation());
  var protection = range.protect().setDescription('Protected by spreadsheet actions');

  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script will throw an exception upon removing the group.
  var actor = Session.getEffectiveUser();
  protection.addEditor(actor);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
