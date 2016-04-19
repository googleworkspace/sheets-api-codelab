
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
    // Add header rows.
    var dataSheetId = spreadsheet.sheets[0].properties.sheetId;
    var requests = [
      buildHeaderRowRequest(dataSheetId),
    ];
    // TODO: Add pivot table and chart.
    var request = {
      spreadsheetId: spreadsheet.spreadsheetId,
      resource: {
        requests: requests
      }
    };
    self.service.spreadsheets.batchUpdate(request, function(err, response) {
      if (err) {
        return callback(err);
      }
      return callback(null, spreadsheet);
    });

  });
};


var COLUMNS = [
  { field: 'id', header: 'ID' },
  { field: 'customerName', header: 'Customer Name'},
  { field: 'productCode', header: 'Product Code' },
  { field: 'unitsOrdered', header: 'Units Ordered' },
  { field: 'unitPrice', header: 'Unit Price' },
  { field: 'status', header: 'Status'}
];

/**
 * Builds a request that sets the header row.
 * @param  {string} sheetId The ID of the sheet.
 * @return {Object}         The reqeuest.
 */
function buildHeaderRowRequest(sheetId) {
  var cells = COLUMNS.map(function(column) {
    return {
      userEnteredValue: {
        stringValue: column.header
      },
      userEnteredFormat: {
        textFormat: {
          bold: true
        }
      }
    }
  });
  return {
    updateCells: {
      start: {
        sheetId: sheetId,
        rowIndex: 0,
        columnIndex: 0
      },
      rows: [
        {
          values: cells
        }
      ],
      fields: 'userEnteredValue,userEnteredFormat.textFormat.bold'
    }
  };
}
