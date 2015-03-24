/**
 * @file: Allows copying, moving and sharing files and folders in Google Drive.
 */

var plugin = new SAplugin('fileManagement');
plugin.version = 2;

/**
 * Options for moving files/folders to folders:
 * If column options are used, URL or ID for file/folder will be read from that column.
 * If set to false, the fallback ID will be used for all processed rows.
 */
plugin.options.sourceFileUrlColumn = 5;
plugin.options.sourceFileUrlFallback = 'https://docs.google.com/document/d/1ZOrHHmdfkM3jQX0mjTsNh7Ln9o8icSLk4pFYvecrPFg/edit';
plugin.options.sourceFolderUrlColumn = 6;
plugin.options.sourceFolderUrlFallback = 'https://docs.google.com/folderview?id=0BzgECFpHWbvRfmgySUExQTAzbTNSamFpVHdlak84QnRRd0U5VHZIWS1BbjBiS3dyZWdBR00';
plugin.options.targetFolderUrlColumn = false;
plugin.options.targetFolderUrlFallback = 'https://docs.google.com/folderview?id=0BzgECFpHWbvRfmgySUExQTAzbTNSamFpVHdlak84QnRRd0U5VHZIWS1BbjBiS3dyZWdBR00';
// When adding files/folders to a folder: Set to true to also remove it from all folders it was placed in before the move.
plugin.options.removeFromOldFolders = true;
// When adding files/folders to a folder: Set to true to also remove it from the root folder.
plugin.options.removeFromRoot = true;

/**
 * If new files/folders are created or renamed, a pattern for file names are specified here.
 * You can use replacement patterns like '%2%' to insert the value in column 2.
 */
plugin.options.fileNamePattern = 'Workbook for %2%';

/**
 * If new files/folders are created, or renamed, these column can be used to write ID, name and URL.
 * Set options to false to avoid printing out the data.
 */
plugin.options.newFileIdColumn = false;
plugin.options.newFileNameColumn = 7;
plugin.options.newFileLinkColumn = false;
plugin.options.newFolderIdColumn = false;
plugin.options.newFolderNameColumn = 12;
plugin.options.newFolderLinkColumn = 13;

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
  SA.plugins.fileManagement.options.newFileLinkColumn = false;
  SA.executeBulkAction('fileManagement', 'renameFile');
  SA.executeBulkAction('fileManagement', 'printFileData');
}
function fileManagementCreateFolder() {
  SA.executeBulkAction('fileManagement', 'createFolder');
}

// Moves a file to a folder.
plugin.moveFile = function(row) {
  // Load source file and target folder, possibly from fallback.
  var file = SA.fetch.file(row, 'sourceFileUrl');
  var folder = SA.fetch.folder(row, 'targetFolderUrl');

  if (this.options.removeFromRoot) {
    DriveApp.getRootFolder().removeFile(file);
  }
  if (this.options.removeFromOldFolders) {
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
  var sourceFolder = SA.fetch.folder(row, 'sourceFolderUrl');
  var targetFolder = SA.fetch.folder(row, 'targetFolderUrl');

  if (this.options.removeFromRoot) {
    DriveApp.getRootFolder().removeFolder(sourceFolder);
  }
  if (this.options.removeFromOldFolders) {
    var folders = sourceFolder.getParents();
    while (folders.hasNext()) {
      folders.next().removeFolder(sourceFolder);
    }
  }
  targetFolder.addFolder(sourceFolder);
}

// Copies a file.
plugin.copyFile = function(row) {
  var file = SA.fetch.file(row, 'sourceFileUrl');
  var copy = file.makeCopy(SA.fetch.replacedText(row, this.options.fileNamePattern));
  this.printFileData(row, copy);
}

// Renames a file.
plugin.renameFile = function(row) {
  var file = SA.fetch.file(row, 'sourceFileUrl');
  file.setName(SA.fetch.replacedText(row, this.options.fileNamePattern));
}

// Prints file data to the main sheet. If no file is passed as argument, sourceFileId
// column or fallback will be used.
plugin.printFileData = function(row, file) {
  if (!file) {
    file = SA.fetch.file(row, 'sourceFileUrl');
  }
  Logger.log(file.getName());
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

// Prints folder data to the main sheet.
plugin.printFolderData = function(row, folder) {
  if (this.options.newFolderIdColumn) {
    SA.fetch.cell(row, this.options.newFolderIdColumn).setValue(folder.getId());
  }
  if (this.options.newFolderNameColumn) {
    SA.fetch.cell(row, this.options.newFolderNameColumn).setValue(folder.getName());
  }
  if (this.options.newFolderLinkColumn) {
    SA.fetch.cell(row, this.options.newFolderLinkColumn).setValue(folder.getUrl());
  }
}

// Creates a folder.
plugin.createFolder = function(row) {
  var folder = DriveApp.createFolder(SA.fetch.replacedText(row, this.options.fileNamePattern));
  this.printFolderData(row, folder);
}

// Loads a Google Drive file with column from the specified option, including fallbacks.
plugin.fetchers.file = function(row, option) {
  var file;
  if (SA.plugins.fileManagement.options[option + 'Column']) {
    file = SA.fetch.fileByUrl(SA.fetch.cell(row, SA.plugins.fileManagement.options[option + 'Column']).getValue());
  }
  else {
    file = SA.fetch.fileByUrl(SA.plugins.fileManagement.options[option + 'Fallback']);
  }
  return file;
}

// Loads a Google Drive folder with column from the specified option, including fallbacks.
plugin.fetchers.folder = function(row, option) {
  var folder;
  if (SA.plugins.fileManagement.options[option + 'Column']) {
    folder = SA.fetch.folderByUrl(SA.fetch.cell(row, SA.plugins.fileManagement.options[option + 'Column']).getValue());
  }
  else {
    folder = SA.fetch.folderByUrl(SA.plugins.fileManagement.options[option + 'Fallback']);
  }
  return folder;
}

// Fetches a Google Drive file based on a url or a file ID.
plugin.fetchers.fileByUrl = function(url) {
  var file;
  var parts = url.split('/').reverse();
  for (var i in parts) {
    try {
      file = DriveApp.getFileById(parts[i]);
      return file;
    }
    catch(e) {}
  }
  throw('The url ' + url + ' does not lead to a valid and accessible Google Drive file.');
}

// Fetches a Google Drive folder based on a url or a folder ID.
plugin.fetchers.folderByUrl = function(url) {
  var folder;
  var parts = url.split('/').reverse();
  for (var i in parts) {
    try {
      folder = DriveApp.getFolderById(parts[i]);
      return folder;
    }
    catch(e) {}
  }
  throw('The url ' + url + ' does not lead to a valid and accessible Google Drive folder.');
}
