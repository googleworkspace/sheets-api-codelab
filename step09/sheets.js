/*
  Copyright 2016 Google, Inc.

  Licensed to the Apache Software Foundation (ASF) under one or more contributor
  license agreements. See the NOTICE file distributed with this work for
  additional information regarding copyright ownership. The ASF licenses this
  file to you under the Apache License, Version 2.0 (the "License"); you may not
  use this file except in compliance with the License. You may obtain a copy of
  the License at

  http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
  WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
  License for the specific language governing permissions and limitations under
  the License.
*/

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

/**
 * Sync the orders to a spreadsheet.
 * @param  {string}   spreadsheetId The ID of the spreadsheet.
 * @param  {string}   sheetId       The ID of the sheet.
 * @param  {Array}    orders        The list of orders.
 * @param  {Function} callback      The callback function.
 */
SheetsHelper.prototype.sync = function(spreadsheetId, sheetId, orders, callback) {
  var requests = [];
  // Resize the sheet.
  requests.push({
    updateSheetProperties: {
      properties: {
        sheetId: sheetId,
        gridProperties: {
          rowCount: orders.length + 1,
          columnCount: COLUMNS.length
        }
      },
      fields: 'gridProperties(rowCount,columnCount)'
    }
  });
  // Set the cell values.
  requests.push({
    updateCells: {
      start: {
        sheetId: sheetId,
        rowIndex: 1,
        columnIndex: 0
      },
      rows: buildRowsForOrders(orders),
      fields: '*'
    }
  });
  // Send the batchUpdate request.
  var request = {
    spreadsheetId: spreadsheetId,
    resource: {
      requests: requests
    }
  };
  this.service.spreadsheets.batchUpdate(request, function(err) {
    if (err) {
      return callback(err);
    }
    return callback();
  });
};

/**
 * Builds an array of RowData from the orders provided.
 * @param  {Array} orders The orders.
 * @return {Array}        The RowData.
 */
function buildRowsForOrders(orders) {
  return orders.map(function(order) {
    var cells = COLUMNS.map(function(column) {
      switch (column.field) {
        case 'unitsOrdered':
          return {
            userEnteredValue: {
              numberValue: order.unitsOrdered
            },
            userEnteredFormat: {
              numberFormat: {
                type: 'NUMBER',
                pattern: '#,##0'
              }
            }
          };
          break;
        case 'unitPrice':
          return {
            userEnteredValue: {
              numberValue: order.unitPrice
            },
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '"$"#,##0.00'
              }
            }
          };
          break;
        case 'status':
          return {
            userEnteredValue: {
              stringValue: order.status
            },
            dataValidation: {
              condition: {
                type: 'ONE_OF_LIST',
                values: [
                  { userEnteredValue: 'PENDING' },
                  { userEnteredValue: 'SHIPPED' },
                  { userEnteredValue: 'DELIVERED' }
                ]
              },
              strict: true,
              showCustomUi: true
            }
          };
          break;
        default:
          return {
            userEnteredValue: {
              stringValue: order[column.field].toString()
            }
          };
      }
    });
    return {
      values: cells
    };
  });
}
