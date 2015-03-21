/**
 * @file: This plugin allows reading color status in student workbooks.
 * It builds on settings in studentWorkbook and colorStudentWorkbook.
 */

var plugin = new SAplugin('readColorsStudentWorkbook');

/**
 * This plugin uses options from other plugins:
 *   studentWorkbooks: Sheet name
 *   colorStudentWorkbook: ok color, not ok color and half-way color
 */

// The range to read, in A1 format. If false, the current selection will be used.
plugin.options.range = false;

// Columns where number of ok cells, not ok cells, and half-way cells should be recorded.
// Set to false to omit writing that count.
plugin.options.okColumn = 9;
plugin.options.notOkColumn = 10;
plugin.options.halfwayColumn = 11;

// Version and dependencies.
plugin.version = 1;
plugin.subVersion = 1;
plugin.dependencies = {
  SA : {
    version : 1
  },
  studentWorkbooks : {
    version : 1
  },
  colorStudentWorkbook : {
    version : 1,
    subVersion : 2
  },
};
// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  readColorsStudentWorkbookExecute : {
    title : 'Read number of colored cells',
    description : 'Here is some description...'
  },
}

// Menu callbacks.
function readColorsStudentWorkbookExecute() {
  SA.executeBulkAction('readColorsStudentWorkbook', 'execute');
}

plugin.execute = function(row) {
  // Get the background colors in the student sheet.
  var range = SA.fetch.studentSheetRange(row, this.options.range || SpreadsheetApp.getActiveRange().getA1Notation());
  var backgrounds = range.getBackgrounds();

  // Count how many times the three signal color appears in the backgrounds.
  var counts = {};
  counts[SA.plugins.colorStudentWorkbook.options.colorOk] = 0;
  counts[SA.plugins.colorStudentWorkbook.options.colorNotOk] = 0;
  counts[SA.plugins.colorStudentWorkbook.options.colorHalfway] = 0;
  var matches = Object.keys(counts);
  for (var i in backgrounds) {
    for (var j in backgrounds[i]) {
      if (matches.indexOf(backgrounds[i][j]) > -1) {
        counts[backgrounds[i][j]]++;
      }
    }
  }

  // Print the counts to the main sheet.
  if (this.options.okColumn) {
    SA.fetch.cell(row, this.options.okColumn).setValue(counts[SA.plugins.colorStudentWorkbook.options.colorOk]);
  }
  if (this.options.notOkColumn) {
    SA.fetch.cell(row, this.options.notOkColumn).setValue(counts[SA.plugins.colorStudentWorkbook.options.colorNotOk]);
  }
  if (this.options.halfwayColumn) {
    SA.fetch.cell(row, this.options.halfwayColumn).setValue(counts[SA.plugins.colorStudentWorkbook.options.colorHalfway]);
  }
}
