/**
 * @file: Imports or exports to Google Contact groups.
 */

var plugin = new SAplugin('contactGroups');

/**
 * This plugin uses options from other plugins:
 *   basics: Name, Google ID
 *   fileManagement: Source file ID, source folder ID
 */

/**
 * The column containing the name of the Google contact group to export to.
 * If set to false, the fallback group name will be used. (The fallback is
 * always used for importing.)
 */
plugin.options.groupNameColumn = 6;
plugin.options.groupNameFallback = 'Mattelärare';
// Set to true to overwrite existing contacts.
plugin.options.allowOverwrite = false;

plugin.dependencies = {
  SA : {
    version : 1
  },
  basics : {
    version : 1
  }
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  contactGroupsExport : {
    title : 'Export to contact group',
    description : 'Here is some description...'
  }
}
plugin.globalActions = {
  contactGroupsImport : {
    title : 'Import from contact group',
    description : 'Here is some description...'
  },
}

function contactGroupsExport() {
  SA.executeBulkAction('contactGroups', 'export');
}
function contactGroupsImport() {
  SA.executeGlobalAction('contactGroups', 'import');
}

plugin.export = function(row) {
  var group;
  if (this.options.groupNameColumn) {
    group = ContactsApp.getContactGroup(SA.fetch.cell(row, this.options.groupNameColumn).getValue());
  }
  else {
    group = ContactsApp.getContactGroup(this.options.groupNameFallback);
  }
  var contact = SA.fetch.contact(SA.fetch.cell(row, SA.plugins.basics.options.googleIdColumn).getValue(), SA.fetch.cell(row, SA.plugins.basics.options.nameColumn).getValue());
  group.addContact(contact);
}

plugin.import = function() {
  var group = ContactsApp.getContactGroup(this.options.groupNameFallback);
  var contacts = group.getContacts();
  var row = parseInt(globalOptions.startRow) - 1;
  for (var i in contacts) {
    row++;
    while (this.options.allowOverwrite || SA.fetch.cell(row, this.options.nameColumn).getValue() != '') {
      row++;
    }
    SA.fetch.cell(row, this.options.nameColumn).setValue(contacts[i].getFullName());
    SA.fetch.cell(row, this.options.googleIdColumn).setValue(contacts[i].getPrimaryEmail());
  }
}

// Fetches a contact by email or, if a matching one doesn't exist, creates a new contact
// using email and fullName.
plugin.fetchers.contact = function(email, fullName) {
  var contact = ContactsApp.getContact(email);
  if (contact == null) {
    contact = ContactsApp.createContact(fullName.split(' ')[0], fullName.split(' ')[1], email);
  }
  return contact;
}
