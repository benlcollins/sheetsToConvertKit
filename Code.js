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
 * setup menu to run print ConvertKit function from Sheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('ConvertKit Menu')
    .addItem('Get list growth', 'postConvertKitDataToSheet')
    .addToUi();

}

/********************************************************************************************
 * ENDPOINT CALLS
********************************************************************************************/

/**
 * function to retrieve ConvertKit Subs
 */
function getConvertKitUnsubs() {

  // get yesterday in correct format
  const yesterday = getYesterday();

  // URL for the ConvertKit API
  const root = 'https://api.convertkit.com/v3/';  
  const endpoint = 'subscribers';
  const query = `?api_secret=${API_SECRET}&from=${yesterday}&to=${yesterday}&sort_field=cancelled_at`;

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
  const newUnsubs = jsonData.total_subscribers;
  
  // return unsubscribes yesterday
  return newUnsubs;
}

/**
 * function to retrieve ConvertKit Subs
 */
function getConvertKitSubs() {
  
  // get yesterday in correct format
  const yesterday = getYesterday();

  // URL for the ConvertKit API
  const root = 'https://api.convertkit.com/v3/';  
  const endpoint = 'subscribers';
  const query = `?api_secret=${API_SECRET}&from=${yesterday}&to=${yesterday}`;

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
  const newSubs = jsonData.total_subscribers;
  
  // return total new subscribers yesterday
  return newSubs;
}

/**
 * add the data to our sheet
 */
function postConvertKitDataToSheet() {
  
  // Get Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('listData');
  const lastRow = sheet.getLastRow();

  // get yesterday date
  const yesterday = getYesterday();

  // get data
  const newSubs = getConvertKitSubs();
  const newUnsubs = getConvertKitUnsubs();

  // paste results into Sheet
  sheet.getRange(lastRow+1,1).setValue(yesterday);
  sheet.getRange(lastRow+1,2).setValue(newSubs);
  sheet.getRange(lastRow+1,3).setValue(newUnsubs);
  sheet.getRange(lastRow+1,4).setFormulaR1C1("=R[0]C[-2]-R[0]C[-1]");
  sheet.getRange(lastRow+1,5).setFormulaR1C1("=R[-1]C[0]+R[0]C[-1]");
  
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
