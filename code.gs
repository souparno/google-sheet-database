/*   
   Copyright 2011 Martin Hawksey

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

// Usage
//  1. Enter sheet name where data is to be written below
        var SHEET_NAME = "Form responses 1";
        
//  2. Run > setup
//
//  3. Publish > Deploy as web app 
//    - enter Project Version name and click 'Save New Version' 
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously) 
//
//  4. Copy the 'Current web app URL' and post this in your form/script action 
//
//  5. Insert column names on your destination sheet matching the parameter names of the data you are passing in (exactly matching case)

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e){
  return handleResponse(e);
}

function doPost(e){
  return handleResponse(e);
}

function handleResponse(e) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
    var action = e.parameter.action;
    
    if (action == 'create') {
      return create(e);
    } else if (action == 'retrieve') {
      return retrieve(e);
    } else if (action == 'update') {
      return update(e);  
    } else if (action == 'delete') {
      return del(e);
    }
  } catch(e){
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function getDataArr(headers, e){
    var row = [];
    // loop through the header columns
    for (i in headers){
      if (headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
        row.push(new Date());
      } else { // else use header name to get data
        row.push(e.parameter[headers[i]]);
      }
    }
    
    return row;
}


function create(e) {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(SHEET_NAME);
    
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = e.parameter.header_row || 1;
    var numColumns = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, numColumns).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var row = getDataArr(headers, e);
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
}

function retrieve(e) {
  var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  var sheet = doc.getSheetByName(SHEET_NAME);
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  var range =  sheet.getRange(1, 1, numRows, numColumns);
  var values = range.getValues();
    
  return ContentService
    .createTextOutput(JSON.stringify({"result":"success", "values": values}))
    .setMimeType(ContentService.MimeType.JSON);
}

function update(e) {
  var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  var sheet = doc.getSheetByName(SHEET_NAME);
  var numColumns = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, numColumns).getValues()[0];
  var row = getDataArr(headers, e);
  var rowId = e.parameter.rowId;
    
  // more efficient to set values as [][] array than individually
  sheet.getRange(rowId, 1, 1, numColumns).setValues([row]);
  // return json success results
  return ContentService
      .createTextOutput(JSON.stringify({"result":"success", "row": rowId}))
      .setMimeType(ContentService.MimeType.JSON);
}

function del(e) {

}

function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty("key", doc.getId());
}
