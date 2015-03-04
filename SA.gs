var globalOptions = {
  workbook : SpreadsheetApp.getActiveSpreadsheet(),
  mainSheetName : 'Sheet1',
  startRow : 3,
  selectColumn : 1
};

function onOpen() {
  SA.buildMenu();
}

var SA = {
  // Version declarations.
  version : 1,
  subVersion : 1,

  // Method for executing bulk actions on selected rows.
  executeBulkAction : function(plugin, callback, mode) {
    this.plugins[plugin].initialize();
    globalOptions.sheet = globalOptions.workbook.getSheetByName(globalOptions.mainSheetName);
    globalOptions.numRows = globalOptions.sheet.getLastRow() - globalOptions.startRow + 1;

    var items = globalOptions.sheet.getRange(globalOptions.startRow, globalOptions.selectColumn, globalOptions.numRows).getValues();
    for (var r in items) {
      for (var c in items[r]) {
        if (items[r][c] == 1) {
          SpreadsheetApp.getActiveSpreadsheet().toast('Processing...', 'Row ' + (parseInt(r) + 1))
          this.plugins[plugin][callback](parseInt(r) + globalOptions.startRow);
          SpreadsheetApp.getActiveSpreadsheet().toast('Done.', 'Row ' + (parseInt(r) + 1))
        }
      }
    }
  },

  // Method for executing actions on the sheet as a whole.
  executeGlobalAction : function(plugin, callback) {
    this.plugins[plugin].initialize();
    globalOptions.sheet = globalOptions.workbook.getSheetByName(globalOptions.mainSheetName);
    globalOptions.numRows = globalOptions.sheet.getLastRow() - globalOptions.startRow + 1;
    this.plugins[plugin][callback]();
  },

  // Adds all enabled actions to the menu.
  buildMenu : function() {
    var ui = SpreadsheetApp.getUi();

    // Create the dynamic bulk actions and global actions sub menus.
    var bulkActions = ui.createMenu('Bulk actions');
    var empty = true;
    for (var i in SA.plugins) {
      for (var j in SA.plugins[i].bulkActions) {
        if (this.callbackIsEnabled(j)) {
          bulkActions.addItem(SA.plugins[i].bulkActions[j].title, j);
          empty = false;
        }
      }
    }
    if (empty) {
      bulkActions.addItem('Empty (run setup to add bulk actions)', 'SAsetup');
    }

    var globalActions = ui.createMenu('Actions');
    var empty = true;
    for (var i in SA.plugins) {
      for (var j in SA.plugins[i].globalActions) {
        if (this.callbackIsEnabled(j)) {
          globalActions.addItem(SA.plugins[i].globalActions[j].title, j);
          empty = false;
        }
      }
    }
    if (empty) {
      globalActions.addItem('Empty (run setup to add actions)', 'SAsetup');
    }

    // Create menu and add sub menus plus static menu entries.
    ui.createMenu('Spreadsheet actions')
    .addSubMenu(bulkActions)
    .addSubMenu(globalActions)
    .addSeparator()
    .addItem('Setup', 'SAsetup')
    .addItem('Help', 'SAhelp')
    .addToUi();
  },

  // Methods for enabling/disabling menu items.
  callbackIsEnabled : function(id) {
    if (PropertiesService.getScriptProperties().getProperty('SAdisabled-' + id) == 1) {
      return false;
    }
    return true;
  },
  enableCallback : function(id) {
    PropertiesService.getScriptProperties().deleteProperty('SAdisabled-' + id);
  },
  disableCallback : function(id) {
    PropertiesService.getScriptProperties().setProperty('SAdisabled-' + id, 1);
  },

  initializedPlugins : {},

  plugins : {},
  tests : {},
  fetch : {},
};

function SAsetup() {
  var checked;
  var output = '<h3>Enable/disable bulk actions</h3>';
  for (var i in SA.plugins) {
    for (var j in SA.plugins[i].bulkActions) {
      if (PropertiesService.getScriptProperties().getProperty('SAdisabled-' + j) == 1) {
        checked = '';
      }
      else {
        checked = 'checked ';
      }
      output += '<input type="checkbox" ' + checked + 'onclick="google.script.run.toggleEnabled(\'' + j + '\')" >' + SA.plugins[i].bulkActions[j].title + '<br />';
    }
  }
  output += '<h3>Enable/disable actions</h3>';
  for (var i in SA.plugins) {
    for (var j in SA.plugins[i].globalActions) {
      if (PropertiesService.getScriptProperties().getProperty('SAdisabled-' + j) == 1) {
        checked = '';
      }
      else {
        checked = 'checked ';
      }
      output += '<input type="checkbox" ' + checked + 'onclick="google.script.run.toggleEnabled(\'' + j + '\')" >' + SA.plugins[i].globalActions[j].title + '<br />';
    }
  }

  var htmlOutput = HtmlService
     .createHtmlOutput(output)
     .setSandboxMode(HtmlService.SandboxMode.NATIVE)
     .setTitle('Spreadsheet Actions setup');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);

  globalOptions.workbook.setFrozenRows(globalOptions.startRow - 1);
  globalOptions.sheet = globalOptions.workbook.getSheetByName(globalOptions.mainSheetName);
  SA.plugins.basics.initialize();
  if (globalOptions.startRow > 2) {
    for (var i = 1; i <= globalOptions.sheet.getLastColumn(); i++) {
      SA.fetch.cell(1, i).setValue(i);
    }
  }
}

function toggleEnabled(e) {
  if (SA.callbackIsEnabled(e)) {
    SA.disableCallback(e);
  }
  else {
    SA.enableCallback(e);
  }
  SA.buildMenu();
}

function SAhelp() {
  var htmlOutput = HtmlService.createHtmlOutput()
     .setSandboxMode(HtmlService.SandboxMode.NATIVE)
     .setTitle('Spreadsheet Actions help');
  htmlOutput.append('<p>Spreadsheet Actions allows running actions on a whole list of spreadsheet entries, or selected rows in a list.</p>');
  htmlOutput.append('<p>This can be used to send customized emails, copying and sharing Google Drive files, adding people to Google Contacts, and whatnot.</p>');
  htmlOutput.append('<p>Due to changes in how user interface works in Google scripts, the UI in Spreadsheet Actions is kept to a bare minimum. Instead, you change the settings inside the scripts themselves.</p>');
  htmlOutput.append('<p>To find and edit settings, follow these steps:</p>');
  htmlOutput.append('<ul><li>Go to "tools" and "script editor".</li>');
  htmlOutput.append('<li>In the window that opens, locate the relevant script/plugin file in the left-hand list. (You might have to open a few before you find the right one.)</li>');
  htmlOutput.append('<li>Click on a script file to open it. At the top of the file there should be a list of options, and an explanation of what each option does.</li>');
  htmlOutput.append('<li>Make any changes to the options, and then save (by clicking the floppy disc icon).</li>');
  htmlOutput.append('<li>Go back to the spreadsheet and run the action from the "Student Actions" menu.</li></ul>');
  htmlOutput.append('<p>Once you have changed the options, you don\'t have to do it again until you want to use other settings.</p>');
  htmlOutput.append('<p>You can turn on/off menu items by visiting the settings.</p>');
  htmlOutput.append('<p>You can download and install new Student Actions plugins, to allow more actions.</p>');

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function SAplugin(id) {
  this.id = id;
  SA.plugins[id] = this;
  this.title = 'This plugin has no title.';
  this.description = 'This plugin has no description.';
  this.version = 1;
  this.subVersion = 1;
  this.dependencies = {
    SA : {
      version : 1,
    }
  };
  this.options = {};
  this.bulkActions = {};
  this.globalActions = {};
  this.fetchers = {};
  this.initialize = function() {
    if (SA.initializedPlugins[this.id]) {
      return;
    }
    for (var i in this.fetchers) {
      if (SA.fetch[i] == undefined) {
        SA.fetch[i] = this.fetchers[i];
      }
      else {
        throw('Cannot add fetcher "' + i + '" in plugin ' + this.id + '. A fetcher with that name already exists');
      }
    }
    this.verifyDependencies();
    SA.initializedPlugins[this.id] = true;
  }
  return this;
}

SAplugin.prototype.verifyDependencies = function() {
  var gotVersion, gotSubVersion;
  for (var p in this.dependencies) {
    if (p == 'SA') {
      gotVersion = SA.version;
      gotSubVersion = SA.subVersion;
    }
    else {
      SA.plugins[p].initialize();
      if (SA.plugins[p] == undefined) {
        throw this.title + ': Plugin ' + p + ' is required, but is missing.';
      }
      gotVersion = SA.plugins[p].version;
      gotSubVersion = SA.plugins[p].subVersion;
    }
    if (this.dependencies[p].version != gotVersion) {
      throw this.title + ': Plugin ' + p + ' must be version ' + this.dependencies[p].version + ' but is ' + gotVersion + '.';
    }
    if (this.dependencies[p].subVersion > gotSubVersion) {
      throw this.title + ': Plugin ' + p + ' must be at least sub version ' + this.dependencies[p].subVersion + ' but is only ' + gotSubVersion + '.';
    }
  }
}
