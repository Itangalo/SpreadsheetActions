/**
 * @file: Start each plugin with a comment block, explaining briefly what this
 * plugin does.
 * This plugin serves as a code example. It allows writing content to specified
 * columns in the main sheet.
 */

var plugin = new SAplugin('pluginExample');

/**
 * The options should go right after the plugin creation, making it easier for
 * users to find and modify them. Each option should be explained in a comment.
 * (The options below are just examples. Use whatever your plugin needs.)
 */

// The number of the column to write to.
plugin.options.column = 6;
// Any static text or formula you wish to add to cells.
plugin.options.text = 'I am an example text';
// Set to true to allow overwriting existing content.
plugin.options.allowOverwrite = false;

/**
 * Declaration of dependencies.
 *
 * Declaration of sub version dependency is optional.
 */
plugin.dependencies = {
  // Require Spreadsheet Actions version 1, sub version ≥ 1.
  SA : {
    version : 1,
    subVersion : 1
  },
  // Require plugin 'basics' version 1, any sub version.
  basics : {
    version : 1,
  }
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  pluginExampleExecute : {
    title : 'Write to cells',
    description : 'Here is some description...'
  },
  pluginExampleExecute2 : {
    title : 'Write to column 7'
  }
}

/**
 * Callbacks for the menu plugins. Names must match the keys in the menu entry
 * declarations. It is wise to prefix with the plugin name, to avoid name
 * conflicts.
 */
function pluginExampleExecute() {
  SA.executeBulkAction('pluginExample', 'execute');
}
// A callback can for example change options, or combine several actions in sequence.
function pluginExampleExecute2() {
  SA.plugins.writeToCell.options.column = 7;
  SA.executeBulkAction('pluginExample', 'execute');
  SA.plugins.writeToCell.options.column = 8;
  SA.plugins.writeToCell.options.text = 'New text';
  SA.executeBulkAction('pluginExample', 'execute');
}

// The function actually writing cellt content.
plugin.execute = function(row, column, text) {
  if (this.options.allowOverwrite) {
    SA.fetch.cell(row, column || this.options.column).setValue(text || this.options.text);
  }
  else if (SA.fetch.cell(row, column || this.options.column).isBlank()) {
    SA.fetch.cell(row, column || this.options.column).setValue(text || this.options.text);
  }
}
