/**
 * @file: Sets and revokes access to files and folders in Google Drive.
 * Also contains a bulk action for creating 'class folder structure', meaning
 * a main (class) folder containing one folder for each student. This folder contains a
 * folder viewable by the student, which in turn contains a folder that is editable
 * by the student.
 */

var plugin = new SAplugin('fileAccess');

/**
 * This plugin uses options from other plugins:
 *   basics: Google ID
 *   fileManagement: Source file ID, source folder ID
 */

/**
 * Options determining what kind of access to grant/revoke. Note that all access types
 * do not apply to all file types.
 */
plugin.options.grantViewAccess = true;
plugin.options.grantEditAccess = false;
plugin.options.grantCommentAccess = false;
/**
 * A string with account email addresses that should be granted edit access to all processed
 * files. Separated by commas.
 */
plugin.options.editorEmails = '';
// Folder ID for any main folder. Used by 'create class folder structure'.
plugin.options.mainFolderIdFallback = '0BzgECFpHWbvRfmgySUExQTAzbTNSamFpVHdlak84QnRRd0U5VHZIWS1BbjBiS3dyZWdBR00';
// Name patterns for folders in class folder structure.
plugin.options.teacherOnlyFolderNamePattern = '%2%';
plugin.options.teacherOnlyFolderIdColumn = '6';
plugin.options.studentViewFolderNamePattern = 'Matte 1b: %2% (endast visa)';
plugin.options.studentViewFolderIdColumn = '7';
plugin.options.studentEditFolderNamePattern = 'Matte 1b: %2% (redigerbar)';
plugin.options.studentEditFolderIdColumn = '8';


plugin.dependencies = {
  SA : {
    version : 1
  },
  fileManagement : {
    version : 1
  }
};

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  fileAccessFileGrant : {
    title : 'Grant access to files',
    description : 'Here is some description...'
  },
  fileAccessFolderGrant : {
    title : 'Grant access to folders',
    description : 'Here is some description...'
  },
  fileAccessFileReset : {
    title : 'Reset access to files',
    description : 'Here is some description...'
  },
  fileAccessFolderReset : {
    title : 'Reset access to folders',
    description : 'Here is some description...'
  },
  fileAccessClassFolders : {
    title : 'Create class folder structure',
    description : 'Here is some description...'
  },
}

function fileAccessFileGrant() {
  SA.executeBulkAction('fileAccess', 'fileGrant');
}
function fileAccessFolderGrant() {
  SA.executeBulkAction('fileAccess', 'folderGrant');
}
function fileAccessFileReset() {
  SA.executeBulkAction('fileAccess', 'fileReset');
}
function fileAccessFolderReset() {
  SA.executeBulkAction('fileAccess', 'folderReset');
}
function fileAccessClassFolders() {
  // Create folders that only teacher can view.
  SA.plugins.fileManagement.options.fileNamePattern = SA.plugins.fileAccess.options.teacherOnlyFolderNamePattern;
  SA.plugins.fileManagement.options.newFileIdColumn = SA.plugins.fileAccess.options.teacherOnlyFolderIdColumn;
  SA.executeBulkAction('fileManagement', 'createFolder');
  // Move the folder to the main folder.
  SA.plugins.fileManagement.options.sourceFolderIdColumn = SA.plugins.fileAccess.options.teacherOnlyFolderIdColumn;
  SA.plugins.fileManagement.options.targetFolderIdColumn = false;
  SA.plugins.fileManagement.options.targetFolderIdFallback = SA.plugins.fileAccess.options.mainFolderIdFallback;
  SA.executeBulkAction('fileManagement', 'moveFolder');

  // Create folders that students may view.
  SA.plugins.fileManagement.options.fileNamePattern = SA.plugins.fileAccess.options.studentViewFolderNamePattern;
  SA.plugins.fileManagement.options.newFileIdColumn = SA.plugins.fileAccess.options.studentViewFolderIdColumn;
  SA.plugins.fileManagement.options.targetFolderIdColumn = SA.plugins.fileAccess.options.teacherOnlyFolderIdColumn;
  SA.executeBulkAction('fileManagement', 'createFolder');
  // Move the folder to the teacher-only folder.
  SA.plugins.fileManagement.options.sourceFolderIdColumn = SA.plugins.fileAccess.options.studentViewFolderIdColumn;
  SA.plugins.fileManagement.options.targetFolderIdColumn = SA.plugins.fileAccess.options.teacherOnlyFolderIdColumn;
  SA.executeBulkAction('fileManagement', 'moveFolder');
  // Grant view access to the student.
  SA.plugins.fileAccess.options.grantEditAccess = false;
  SA.plugins.fileAccess.options.grantViewAccess = true;
  SA.executeBulkAction('fileAccess', 'folderGrant');

  // Create folders that students may edit.
  SA.plugins.fileManagement.options.fileNamePattern = SA.plugins.fileAccess.options.studentEditFolderNamePattern;
  SA.plugins.fileManagement.options.newFileIdColumn = SA.plugins.fileAccess.options.studentEditFolderIdColumn;
  SA.executeBulkAction('fileManagement', 'createFolder');
  // Move the folder to the viewable folder.
  SA.plugins.fileManagement.options.sourceFolderIdColumn = SA.plugins.fileAccess.options.studentEditFolderIdColumn;
  SA.plugins.fileManagement.options.targetFolderIdColumn = SA.plugins.fileAccess.options.studentViewFolderIdColumn;
  SA.executeBulkAction('fileManagement', 'moveFolder');
  // Grant view and edit access to the student.
  SA.plugins.fileAccess.options.grantEditAccess = true;
  SA.executeBulkAction('fileAccess', 'folderGrant');
}

plugin.fileGrant = function(row) {
  var file = SA.fetch.file(row, 'sourceFileId');
  var googleId = SA.fetch.cell(row, SA.plugins.basics.options.googleIdColumn).getValue();
  if (this.options.grantViewAccess) {
    file.addViewer(googleId);
  }
  if (this.options.grantEditAccess) {
    file.addEditor(googleId);
  }
  if (this.options.grantCommentAccess) {
    file.addCommenter(googleId)
  }
}
plugin.folderGrant = function(row) {
  var folder = SA.fetch.folder(row, 'sourceFolderId');
  var googleId = SA.fetch.cell(row, SA.plugins.basics.options.googleIdColumn).getValue();
  if (this.options.grantViewAccess) {
    folder.addViewer(googleId);
  }
  if (this.options.grantEditAccess) {
    folder.addEditor(googleId);
  }
}
plugin.fileReset = function(row) {
  var file = SA.fetch.file(row, 'sourceFileId');
  var users = file.getViewers();
  for (var i in users) {
    file.revokePermissions(users[i]);
  }
  var users = file.getEditors();
  for (var i in users) {
    file.revokePermissions(users[i]);
  }
}
plugin.folderReset = function(row) {
  var folder = SA.fetch.folder(row, 'sourceFolderId');
  var users = folder.getViewers();
  for (var i in users) {
    folder.revokePermissions(users[i]);
  }
  var users = folder.getEditors();
  for (var i in users) {
    folder.revokePermissions(users[i]);
  }
}
