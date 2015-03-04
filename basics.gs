var plugin = new SAplugin('basics');

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
