/**
 * @file: Allows copying, moving and sharing files and folders in Google Drive.
 */

var plugin = new SAplugin('fileManagement');

/**
 * Column containing any source file (or folder), used for copying, moving, etc.
 * If set to false, the fallback file ID will be used.
 */
plugin.options.sourceFileIdColumn = 9;
plugin.options.sourceFileIdFallback = '1ZOrHHmdfkM3jQX0mjTsNh7Ln9o8icSLk4pFYvecrPFg';
plugin.options.sourceFolderIdColumn = 6;
plugin.options.sourceFolderIdFallback = '0BzgECFpHWbvRfmgySUExQTAzbTNSamFpVHdlak84QnRRd0U5VHZIWS1BbjBiS3dyZWdBR00';

/**
 * If new files are created, or files are renamed, a pattern for file names are specified here.
 * You can use replacement patterns like '%2%' to insert the value in column 2.
 */
plugin.options.fileNamePattern = 'Copy for %2%';

/**
 * If new files are created, or renamed, these column can be used to write ID, name and link for the file.
 * Set options to false to avoid printing out the data.
 */
plugin.options.newFileIdColumn = 9;
plugin.options.newFileNameColumn = 10;
plugin.options.newFileLinkColumn = 11;

/**
 * If any files/folders should be moved, use this setting to identify to which folder.
 * If set to false, the fallback file ID will be used for all rows.
 */
plugin.options.targetFolderIdColumn = false;
plugin.options.targetFolderIdFallback = '0BzgECFpHWbvRfmgySUExQTAzbTNSamFpVHdlak84QnRRd0U5VHZIWS1BbjBiS3dyZWdBR00';
// When adding files to a folder: Set to true to remove the file from the root folder.
plugin.options.removeFileFromRoot = true;
// When adding files to a folder: Set to true to remove the file from all folders it was placed in before the move.
plugin.options.removeFileFromOldFolders = true;

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
  fileManagementMoveFile : {
    title : 'Move file(s) to folder',
    description : 'Here is some description...'
  },
  fileManagementMoveFolder : {
    title : 'Move folder(s) to folder',
    description : 'Here is some description...'
  },
  fileManagementCopyFile : {
    title : 'Copy file(s)',
    description : 'Here is some description...'
  },
  fileManagementRenameFile : {
    title : 'Rename files',
    description : 'Here is some description...'
  },
  fileManagementCreateFolder : {
    title : 'Create folders',
    description : 'Here is some description...'
  },
}

// Callbacks for the menu plugins.
function fileManagementMoveFile() {
  SA.executeBulkAction('fileManagement', 'moveFile');
}
function fileManagementMoveFolder() {
  SA.executeBulkAction('fileManagement', 'moveFolder');
}
function fileManagementCopyFile() {
  SA.executeBulkAction('fileManagement', 'copyFile');
}
function fileManagementRenameFile() {
  SA.plugins.fileManagement.options.newFileIdColumn = false;
  SA.executeBulkAction('fileManagement', 'renameFile');
  SA.executeBulkAction('fileManagement', 'printFileData');
}
function fileManagementCreateFolder() {
  SA.executeBulkAction('fileManagement', 'createFolder');
}

// Moves a file to a folder.
plugin.moveFile = function(row) {
  // Load source file and target folder, possibly from fallback.
  var file = SA.fetch.file(row, 'sourceFileId');
  var folder = SA.fetch.folder(row, 'targetFolderId');

  if (this.options.removeFileFromRoot) {
    DriveApp.getRootFolder().removeFile(file);
  }
  if (this.options.removeFileFromOldFolders) {
    var folders = file.getParents();
    while (folders.hasNext()) {
      folders.next().removeFile(file);
    }
  }
  folder.addFile(file);
}

// Moves a folder to a folder.
plugin.moveFolder = function(row) {
  // Load source file and target folder, possibly from fallback.
  var sourceFolder = SA.fetch.folder(row, 'sourceFolderId');
  var targetFolder = SA.fetch.folder(row, 'targetFolderId');

  if (this.options.removeFileFromRoot) {
    DriveApp.getRootFolder().removeFolder(sourceFolder);
  }
  if (this.options.removeFileFromOldFolders) {
    var folders = sourceFolder.getParents();
    while (folders.hasNext()) {
      folders.next().removeFolder(sourceFolder);
    }
  }
  targetFolder.addFolder(sourceFolder);
}

// Copies a file.
plugin.copyFile = function(row) {
  var file = SA.fetch.file(row, 'sourceFileId');
  var copy = file.makeCopy(SA.fetch.replacedText(row, this.options.fileNamePattern));
  this.printFileData(row, copy);
}

// Renames a file.
plugin.renameFile = function(row) {
  var file = SA.fetch.file(row, 'sourceFileId');
  file.setName(SA.fetch.replacedText(row, this.options.fileNamePattern));
}

// Prints file data to the main sheet. If no file is passed as argument, sourceFileId
// column or fallback will be used.
plugin.printFileData = function(row, file) {
  if (!file) {
    file = SA.fetch.file(row, 'sourceFileId');
  }
  if (this.options.newFileIdColumn) {
    SA.fetch.cell(row, this.options.newFileIdColumn).setValue(file.getId());
  }
  if (this.options.newFileNameColumn) {
    SA.fetch.cell(row, this.options.newFileNameColumn).setValue(file.getName());
  }
  if (this.options.newFileLinkColumn) {
    SA.fetch.cell(row, this.options.newFileLinkColumn).setValue(file.getUrl());
  }
}

// Creates a folder.
plugin.createFolder = function(row) {
  var folder = DriveApp.createFolder(SA.fetch.replacedText(row, this.options.fileNamePattern));
  this.printFileData(row, folder);
}

// Loads a Google Drive file with column from the specified option, including fallbacks.
plugin.fetchers.file = function(row, option) {
  var file;
  if (SA.plugins.fileManagement.options[option + 'Column']) {
    file = DriveApp.getFileById(SA.fetch.cell(row, SA.plugins.fileManagement.options[option + 'Column']).getValue());
  }
  else {
    file = DriveApp.getFileById(SA.plugins.fileManagement.options[option + 'Fallback']);
  }
//  if (file.getMimeType() == 'application/vnd.google-apps.folder') {
//    file = DriveApp.getFolderById(file.getId());
//  }
  return file;
}

// Loads a Google Drive folder with column from the specified option, including fallbacks.
plugin.fetchers.folder = function(row, option) {
  var folder;
  if (SA.plugins.fileManagement.options[option + 'Column']) {
    folder = DriveApp.getFolderById(SA.fetch.cell(row, SA.plugins.fileManagement.options[option + 'Column']).getValue());
  }
  else {
    folder = DriveApp.getFolderById(SA.plugins.fileManagement.options[option + 'Fallback']);
  }
  return folder;
}
