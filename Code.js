/** 
 * ConvertKit > Google Sheets Tool
 * 
 * https://developers.convertkit.com/#overview
 * 
 */

/********************************************************************************************
 * SETUP
********************************************************************************************/

/**
 * Global variables
 */
const API_KEY = getApiKey();
const API_SECRET = getApiSecret();

/**
 * function to get my ConvertKit API Key from properties service
 */
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty("CK_API_KEY");
}

/**
 * function to get my ConvertKit API Secret from properties service
 */
function getApiSecret() {
  return PropertiesService.getScriptProperties().getProperty("CK_API_SECRET");
}

/**
 * test script properties
 */
function test(){
  console.log(API_KEY);
  console.log(API_SECRET);
}

/**
 * setup menu to run print ConverKit function from Sheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('ConvertKit Menu')
    .addItem('Get ConvertKit data', 'postConvertKitDataToSheet')
    .addItem('Email ConvertKit report', 'exportAndSend')
    .addToUi();

}

/********************************************************************************************
 * EMAIL FUNCTIONS
********************************************************************************************/

/**
 * send pdf of sheet to stakeholders
 */
function exportAndSend() {

  // get today's date
  const d = formatDate(new Date());

  // get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportUrl = ss.getUrl();

  // make copy of Sheet
  const copiedSheet = ss.copy(`Copy of  ${ss.getName()} ${d}`);

  // copy - paste report as values to avoid broken links when sheets are deleted
  const copiedSheetReport = copiedSheet.getSheetByName('Report');
  const vals = copiedSheetReport.getRange(1,1,copiedSheetReport.getMaxRows(),copiedSheetReport.getMaxColumns()).getValues();
  copiedSheetReport.getRange(1,1,copiedSheetReport.getMaxRows(),copiedSheetReport.getMaxColumns()).setValues(vals);

  // delete redundant sheets
  const sheets = copiedSheet.getSheets();
  sheets.forEach(function(sheet){
    if (sheet.getSheetName() != copiedSheetReport.getSheetName()) {
      copiedSheet.deleteSheet(sheet);
    }
  });

  // create email
  const body = `A pdf copy of your ConvertKit report is attached.<br><br>
    To access the live version in your Google Sheets,; 
      <a href="${reportUrl}">click here</a>`;

  // send email
  GmailApp.sendEmail('example@example.com',`ConvertKit Report ${d}`,'',
    {
      htmlBody: body,
      attachments: [copiedSheet.getAs(MimeType.PDF)]
    });

  // delete temporary sheet
  DriveApp.getFileById(copiedSheet.getId()).setTrashed(true);

}

/********************************************************************************************
 * SHEET FUNCTIONS
********************************************************************************************/

/**
 * add the data to our sheet
 */
function postConvertKitDataToSheet() {
  
  // Get Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName('listData');
  const broadcastSheet = ss.getSheetByName('broadcastData');
  const lastRow = listSheet.getLastRow();

  // get yesterday date
  const yesterday = getYesterday();

  // get data
  const totalSubs = getConvertKitTotalSubs();
  const broadcastData = getBroadcastData();

  // paste list growth results into Sheet
  listSheet.getRange(lastRow+1,1).setValue(yesterday);
  listSheet.getRange(lastRow+1,2).setValue(totalSubs);
  listSheet.getRange(lastRow+1,3).setFormulaR1C1("=R[0]C[-1]-R[-1]C[-1]");

  // paste latest campaign data into Sheet
  broadcastSheet.getRange(2,1,broadcastData.length,8).setValues(broadcastData);
  
}


/********************************************************************************************
 * API CALLS
********************************************************************************************/

/**
 * function to return broadcast data
 */
function getBroadcastData() {

  // URL for the ConvertKit API
  const root = 'https://api.convertkit.com/v3/';  
  const endpoint = 'broadcasts';
  const query = `?api_secret=${API_SECRET}`;

  // setup params object
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true
  };
  
  // check api
  console.log(root + endpoint + query);
  
  // call the ConvertKit API
  const response = UrlFetchApp.fetch(root + endpoint + query, params);
  
  // parse data
  const data = response.getContentText();
  const jsonData = JSON.parse(data);
  const broadcastData = jsonData.broadcasts;
  
  // test broadcast data
  //console.log(broadcastData);

  // empty array to hold broadcast data
  const broadcastDataArray = [];

  // add individual broadcast details
  broadcastData.forEach(function(row){

    const id = row.id;
    const created_date = row.created_at;
    const subject = row.subject;

    // get individual data
    const individualData = getIndividualBroadcastData(id);

    const recipients = individualData.broadcast.stats.recipients;
    const open_rate = individualData.broadcast.stats.open_rate;
    const click_rate = individualData.broadcast.stats.click_rate;
    const unsubscribes = individualData.broadcast.stats.unsubscribes;
    const total_clicks = individualData.broadcast.stats.total_clicks;
    
    // push into broadcast data array
    broadcastDataArray.push([
      id,
      created_date,
      subject,
      recipients,
      open_rate,
      click_rate,
      unsubscribes,
      total_clicks
    ]);

  })

  console.log(broadcastDataArray);

  // return broadcast data
  return broadcastDataArray;

}

/**
 * function to return individual broadcast data
 */
function getIndividualBroadcastData(broadcastID) {

  // URL for the ConvertKit API
  const root = 'https://api.convertkit.com/v3/';  
  const endpoint = `broadcasts/${broadcastID}/stats`;
  const query = `?api_secret=${API_SECRET}`;

  // setup params object
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true
  };
  
  // check api
  //console.log(root + endpoint + query);
  
  // call the ConvertKit API
  const response = UrlFetchApp.fetch(root + endpoint + query, params);
  
  // parse data
  const data = response.getContentText();
  const jsonData = JSON.parse(data);

  // return data
  return jsonData;

}

/**
 * function to retrieve ConvertKit List Size
 */
function getConvertKitTotalSubs() {
  
  // get yesterday in correct format
  const yesterday = getYesterday();

  // URL for the ConvertKit API
  const root = 'https://api.convertkit.com/v3/';  
  const endpoint = 'subscribers';
  const query = `?api_secret=${API_SECRET}`;

  // setup params object
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true
  };
  
  // check api
  console.log(root + endpoint + query);
  
  // call the ConvertKit API
  const response = UrlFetchApp.fetch(root + endpoint + query, params);
  
  // parse data
  const data = response.getContentText();
  const jsonData = JSON.parse(data);
  const totalSubs = jsonData.total_subscribers;
  console.log(totalSubs)
  
  // return total new subscribers yesterday
  return totalSubs;
}


/********************************************************************************************
 * HELPER FUNCTIONS
********************************************************************************************/

/**
 * get yesterday's date in correct format
 */
function getYesterday() {

  // get yesterday's date
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const formatYesterday = formatDate(yesterday);
  
  // return formatted yesterday date YYYY-MM-DD
  return formatYesterday;

}

/**
 * format date to YYYY-MM-DD
 */
function formatDate(date) {

  // create new date object
  const d = new Date(date);

  // get component parts
  let month = '' + (d.getMonth() + 1);
  let day = '' + d.getDate();
  const year = d.getFullYear();

  // add 0 to single digit days or months
  if (month.length < 2) 
      month = '0' + month;
  if (day.length < 2) 
      day = '0' + day;

  // return new date string
  return [year, month, day].join('-');
}
