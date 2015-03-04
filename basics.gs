/**
 * @file: This plugin provides some common functionality used by many other
 * plugins, for example specifying name and Google ID column and a method for
 * doing replacements in texts.
 */

var plugin = new SAplugin('basics');

// The column used for name (if any).
plugin.options.nameColumn = 2;
// The column used for Google ID.
plugin.options.googleIdColumn = 3;

// Version and dependencies.
plugin.version = 1;
plugin.subVersion = 1;
plugin.dependencies = {
  SA : {
    version : 1
  }
};

// A fetcher, fetching a cell from given row and column in the main sheet.
plugin.fetchers.cell = function(row, column) {
  row = parseInt(row);
  column = parseInt(column);
  return globalOptions.sheet.getRange(row, column);
}

// A fetcher, replacing all '%3%' in a text with content from column 3. For example.
plugin.fetchers.replacedText = function(row, text) {
  var find;
  var re;
  for (var i = 1; i <= globalOptions.sheet.getLastColumn(); i++) {
    find = '%' + i + '%';
    if (text.search(find) > -1) {
      re = new RegExp(find, 'g');
      text = text.replace(re, SA.fetch.cell(row, i).getValue());
    }
  }
  return text;
}
