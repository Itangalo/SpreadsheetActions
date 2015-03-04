/**
 * @file: Creates and connects each row with a copy of a spreadsheet workbook.
 */

var plugin = new SAplugin('studentWorkbooks');

/**
 * This plugin uses options from other plugins:
 *   basics: Name, Google ID
 */

// ID for the workbook used as template for each student.
plugin.options.templateWorkbookId = '1QtUr7wYfh91d8M1qqYKv6I3C4YiJsVtUikwgiaNW5OQ';
// Name of the sheet in the workbook to make any edits to.
plugin.options.sheetName = 'Betygsunderlag';
// The column for reading/writing ID of the student's workbook ID (and link).
plugin.options.studentWorkbookIdColumn = 4;
plugin.options.studentWorkbookLinkColumn = 5;
// Pattern for naming student copies of the template workbook.
plugin.options.studentWorkbookNamePattern = 'Workbook for %2%';

plugin.dependencies = {
  SA : {
    version : 1
  },
  basics : {
    version : 1
  },
  fileManagement : {
    version : 1
  },
  fileAccess : {
    version : 1
  }
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  studentWorkbooksCreateCopy : {
    title : 'Create student workbook copies',
    description : 'Here is some description...'
  },
};
plugin.globalActions = {
  studentWorkbooksInsertSheetCopy : {
    title : 'Insert a copy of the workbook template',
    description : 'Here is some description...'
  }
};

// Menu callbacks.
function studentWorkbooksCreateCopy() {
  SA.executeBulkAction('studentWorkbooks', 'createCopy');
}
function studentWorkbooksInsertSheetCopy() {
  SA.executeGlobalAction('studentWorkbooks', 'insertSheetCopy');
}

plugin.createCopy = function(row) {
  var copy = SpreadsheetApp.openById(this.options.templateWorkbookId).copy(SA.fetch.replacedText(row, this.options.studentWorkbookNamePattern));
  SA.fetch.cell(row, this.options.studentWorkbookIdColumn).setValue(copy.getId());
  SA.fetch.cell(row, this.options.studentWorkbookLinkColumn).setValue(copy.getUrl());
}

plugin.insertSheetCopy = function() {
  var sheet = SpreadsheetApp.openById(this.options.templateWorkbookId).getSheetByName(this.options.sheetName).copyTo(globalOptions.workbook);
  sheet.activate();
  globalOptions.workbook.moveActiveSheet(2);
}

SA.fetch.studentWorkbook = function(row) {
  var id = SA.fetch.cell(row, SA.plugins.studentWorkbooks.options.studentWorkbookIdColumn).getValue();
  return SpreadsheetApp.openById(id);
}

SA.fetch.studentSheet = function(row, sheetName) {
  sheetName = sheetName || SA.plugins.studentWorkbooks.options.sheetName;
  return SA.fetch.studentWorkbook(row).getSheetByName(sheetName);
}

SA.fetch.studentSheetRange = function(row, A1notation) {
  return SA.fetch.studentSheet(row).getRange(A1notation);
}
