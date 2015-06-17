/**
 * @file: Allows updating cell content in student workbooks. Builds on the studentWorkbook plugin.
 */

var plugin = new SAplugin('updateContentStudentWorkbook');

/**
 * This plugin uses options from other plugins:
 *   studentWorkbooks: Sheet name
 */

// The range to edit, in A1 notation. If false, the selected range will be used.
plugin.options.range = false;
// The color to use, in hex code or accepted CSS codes. If false, the colors of the selected range will be used.
plugin.options.content = false;

plugin.title = 'Update content in student sheets';
plugin.dependencies = {
  studentWorkbooks : {
    version : 1
  },
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  updateContentStudentWorkbook : {
    title : 'Set cell content in student workbook',
    description : 'Here is some description...'
  },
}

// Menu callbacks.
function updateContentStudentWorkbook() {
  SA.executeBulkAction('updateContentStudentWorkbook', 'updateContent');
}

plugin.updateContent = function(row) {
  var range = SA.fetch.studentSheetRange(row, this.options.range || SpreadsheetApp.getActiveRange().getA1Notation());
  if (this.options.content) {
    range.setValue(this.options.content);
  }
  else {
    range.setValues(SpreadsheetApp.getActiveRange().getValues());
  }
}
