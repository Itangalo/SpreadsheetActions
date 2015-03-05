/**
 * @file: Sends email based on a template in a Google document.
 */

var plugin = new SAplugin('mailFromTemplate');

/**
 * This plugin uses options from other plugins:
 *   basics: Google ID
 */

/**
 * ID for the Google document used as template for the email. Any occurances of '%2%'
 * will be replaced with the content in the corresponding column.
 */
plugin.options.templateId = '1ZOrHHmdfkM3jQX0mjTsNh7Ln9o8icSLk4pFYvecrPFg';
// Subject for the email. May contain replacement patterns on the form '%2%'.
plugin.options.subject = 'Feedback from test';
// Row to use for displaying sample of email.
plugin.options.sampleRow = 3;

plugin.dependencies = {
  SA : {
    version : 1
  },
  basics : {
    version : 1
  }
}

// Declaration of menu entries for this plugin.
plugin.bulkActions = {
  mailFromTemplateSendEmail : {
    title : 'Send email based on document template',
    description : 'Here is some description...'
  },
};
plugin.globalActions = {
  mailFromTemplatePreviewEmail : {
    title : 'Preview email based on document template',
    description : 'Here is some description...'
  }
};

// Menu callbacks.
function mailFromTemplateSendEmail() {
  SA.executeBulkAction('mailFromTemplate', 'sendEmail');
}
function mailFromTemplatePreviewEmail() {
  SA.executeGlobalAction('mailFromTemplate', 'previewEmail');
}

plugin.sendEmail = function(row) {
  var message = DocumentApp.openById(this.options.templateId).getBody().getText();
  MailApp.sendEmail(SA.fetch.cell(row, SA.plugins.basics.options.googleIdColumn).getValue(),
                    SA.fetch.replacedText(row, this.options.subject),
                    SA.fetch.replacedText(row, message));
}
plugin.previewEmail = function() {
  var message = DocumentApp.openById(this.options.templateId).getBody().getText();
  message = SA.fetch.replacedText(this.options.sampleRow, message);
  message = message.replace(/\n/g, '<br />');
  var subject = SA.fetch.replacedText(this.options.sampleRow, this.options.subject);
  var htmlOutput = HtmlService.createHtmlOutput(message)
     .setSandboxMode(HtmlService.SandboxMode.NATIVE)
     .setTitle(subject);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
