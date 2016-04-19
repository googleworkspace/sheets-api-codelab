
var google = require('googleapis');
var googleAuth = require('google-auth-library');
var util = require('util');

/**
 * Create a new Sheets helper.
 * @param {string} accessToken An authorized OAuth2 access token.
 * @constructor
 */
var SheetsHelper = function(accessToken) {
  var authClient = new googleAuth();
  var auth = new authClient.OAuth2();
  auth.credentials = {
    access_token: accessToken
  };
  this.service = google.sheets({version: 'v4', auth: auth});
};

module.exports = SheetsHelper;

/**
 * Create a spreadsheet with the given name.
 * @param  {string}   title    The name of the spreadsheet.
 * @param  {Function} callback The callback function.
 */
SheetsHelper.prototype.createSpreadsheet = function(title, callback) {
  var self = this;
  var request = {
    resource: {
      properties: {
        title: title
      },
      sheets: [
        {
          properties: {
            title: 'Data',
            gridProperties: {
              columnCount: 6,
              frozenRowCount: 1
            }
          }
        },
        // TODO: Add more sheets.
      ]
    }
  };
  self.service.spreadsheets.create(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    // TODO: Add header rows.
    return callback(null, spreadsheet);
  });
};
