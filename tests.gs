var plugin = new SAplugin('tests');

// The ID of the plugin to test. If false, all will be tested.
plugin.options.pluginId = false;

// Declaration of menu entries for this plugin.
plugin.globalActions = {
  testsExecute : {
    title : 'Run plugin tests',
    description : 'Here is some description...'
  },
}

// Callbacks for the menu plugins.
function testsExecute() {
  SA.executeGlobalAction('tests', 'execute');
}

plugin.execute = function() {
  var plugins = Object.keys(SA.plugins);
  if (this.options.pluginId) {
    plugins = [this.options.pluginId];
  }
  for (var p in plugins) {
    SA.plugins[plugins[p]].initialize();
    if (typeof SA.tests[plugins[p]] == 'function') {
      SA.tests[plugins[p]]();
    }
    else {
      Logger.log('Warning: Plugin ' + [plugins[p]] + ' has no tests.');
    }
  }
}
