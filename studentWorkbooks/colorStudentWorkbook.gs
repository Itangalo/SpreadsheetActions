/**
 * @file: Allows coloring cells in student workbooks. Builds on the studentWorkbook plugin.
 */

var plugin = new SAplugin('colorStudentWorkbook');

/**
 * This plugin uses options from other plugins:
 *   studentWorkbooks: Sheet name
 */

// The range to edit, in A1 notation. If false, the selected range will be used.
plugin.options.range = false;
// The color to use, in hex code or accepted CSS codes. If false, the colors of the selected range will be used.
plugin.options.color = false;
// Preset colors for special actions.
plugin.options.colorOk = 'lawngreen';
plugin.options.colorNotOk = 'red';
plugin.options.colorHalfway = 'yellow';

plugin.title = 'Color student sheets';
plugin.dependencies = {
  studentWorkbooks : {
    version : 1
  },
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  colorStudentSheetColor : {
    title : 'Color cells in student workbook',
    description : 'Here is some description...'
  },
  colorStudentSheetColorOk : {
    title : 'Mark cells in student workbook ok',
    description : 'Here is some description...'
  },
  colorStudentSheetColorNotOk : {
    title : 'Mark cells in student workbook NOT ok',
    description : 'Here is some description...'
  },
  colorStudentSheetColorHalfway : {
    title : 'Mark cells in student workbook halfway done',
    description : 'Here is some description...'
  },
}

// Menu callbacks.
function colorStudentSheetColor() {
  SA.executeBulkAction('colorStudentWorkbook', 'color');
}
function colorStudentSheetColorOk() {
  SA.plugins.colorStudentWorkbook.options.color = SA.plugins.colorStudentWorkbook.options.colorOk;
  SA.executeBulkAction('colorStudentWorkbook', 'color');
}
function colorStudentSheetColorNotOk() {
  SA.plugins.colorStudentWorkbook.options.color = SA.plugins.colorStudentWorkbook.options.colorNotOk;
  SA.executeBulkAction('colorStudentWorkbook', 'color');
}
function colorStudentSheetColorHalfway() {
  SA.plugins.colorStudentWorkbook.options.color = SA.plugins.colorStudentWorkbook.options.colorHalfway;
  SA.executeBulkAction('colorStudentWorkbook', 'color');
}

plugin.color = function(row) {
  var range = SA.fetch.studentSheetRange(row, this.options.range || SpreadsheetApp.getActiveRange().getA1Notation());
  if (this.options.color) {
    range.setBackground(this.options.color);
  }
  else {
    range.setBackgrounds(SpreadsheetApp.getActiveRange().getBackgrounds());
  }
}
