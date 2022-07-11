/***************************************************
 * Bulk URL PageSpeed Tool (PageSpeed Insights v5)
 * by james@upbuild.io
 ***************************************************/

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    { name: 'Set Report & Log Schedule', functionName: 'scheduleboth' },
    { name: 'Manual Push Report', functionName: 'runTool' },
    { name: 'Manual Push Log', functionName: 'runLog' },
    { name: 'Reset Schedule', functionName: 'resetSchedule' },
  ];
  sheet.addMenu('PageSpeed Menu', entries);
}

var timeMapping = {
  '1AM': 1,
  '2AM': 2,
  '3AM': 3,
  '4AM': 4,
  '5AM': 5,
  '6AM': 6,
  '7AM': 7,
  '8AM': 8,
  '9AM': 9,
  '10AM': 10,
  '11AM': 11,
  '12PM': 12,
  '1PM': 13,
  '2PM': 14,
  '3PM': 15,
  '4PM': 16,
  '5PM': 17,
  '6PM': 18,
  '7PM': 19,
  '8PM': 20,
  '9PM': 21,
  '10PM': 22,
  '11PM': 23,
  '12AM': 24,
};

// Menu 1
function scheduleboth() {
  startScheduledReportOne();
  startScheduledReportTwo();
  startScheduledReportThree();
  startScheduledReportFour();
  startScheduledLog();
  Browser.msgBox('Success! - Report and Log Times Scheduled');
}

function createTrigger(reportDay, newhour, triggerName) {
  // Browser.msgBox('test');
  // Browser.msgBox(`${reportDay} - ${newhour} - ${triggerName}`);

  // throw new Error();
  if (reportDay == 'MONDAY') {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .create();
  }

  if (reportDay == 'TUESDAY') {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .create();
  }

  if (reportDay == 'WEDNESDAY') {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .create();
  }

  if (reportDay == 'THURSDAY') {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .create();
  }

  if (reportDay == 'FRIDAY') {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .create();
  }

  if (reportDay == 'SATURDAY') {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .create();
  }

  if (reportDay == 'SUNDAY') {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .atHour(newhour)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .create();
  }
}

//Run the Report - Phase One
function startScheduledReportOne() {
  var settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var reportDay = settingsSheet.getRange('C10:C10').getValue();
  var reportTime = settingsSheet.getRange('E10:E10').getValue();

  var newhour = timeMapping[reportTime];
  createTrigger(reportDay, newhour, 'runTool');
}

//Run the Report - Phase Two
function startScheduledReportTwo() {
  var settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var reportDay = settingsSheet.getRange('C11:C11').getValue();
  var reportTime = settingsSheet.getRange('E11:E11').getValue();

  var newhour = timeMapping[reportTime];
  createTrigger(reportDay, newhour, 'runTool');
}

//Run the Report - Phase Three
function startScheduledReportThree() {
  var settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var reportDay = settingsSheet.getRange('C12:C12').getValue();
  var reportTime = settingsSheet.getRange('E12:E12').getValue();

  var newhour = timeMapping[reportTime];
  createTrigger(reportDay, newhour, 'runTool');
}

//Run the Report - Phase Four
function startScheduledReportFour() {
  var settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var reportDay = settingsSheet.getRange('C13:C13').getValue();
  var reportTime = settingsSheet.getRange('E13:E13').getValue();

  var newhour = timeMapping[reportTime];
  createTrigger(reportDay, newhour, 'runTool');
}

function startScheduledLog() {
  var settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var reportDay = settingsSheet.getRange('C17:C17').getValue();
  var reportTime = settingsSheet.getRange('E17:E17').getValue();

  var newhour = timeMapping[reportTime];
  createTrigger(reportDay, newhour, 'runLog');
}

// Menu 2
//Run the formula to get the PageSpeed V5 data from each URL
function runTool() {
  var activeSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  var rows = activeSheet.getLastRow();

  for (var i = 6; i <= rows; i++) {
    var workingCell = activeSheet.getRange(i, 2).getValue();
    var stuff = '=runCheck';

    if (workingCell != '') {
      activeSheet.getRange(i, 3).setFormulaR1C1(stuff + '(R[0]C[-1])');
    }
  }
}

// used by runTool
function runCheck(Url) {
  var settingsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var key = settingsSheet.getRange('C7:C7').getValue();
  var strategy = settingsSheet.getRange('C20').getValue();

  var serviceUrl =
    'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=' +
    Url +
    '&key=' +
    key +
    '&strategy=' +
    strategy +
    '';

  var array = [];

  if (key == 'YOUR_API_KEY') return 'Please enter your API key to the script';

  var response = UrlFetchApp.fetch(serviceUrl);

  if (response.getResponseCode() == 200) {
    var content = JSON.parse(response.getContentText());

    if (content != null && content['lighthouseResult'] != null) {
      if (content['captchaResult']) {
        var score =
          content['lighthouseResult']['categories']['performance']['score'];
        var timetointeractive = content['lighthouseResult']['audits'][
          'interactive'
        ]['displayValue'].slice(0, -2);
        var largestcontentfulpaint = content['lighthouseResult']['audits'][
          'largest-contentful-paint'
        ]['displayValue'].slice(0, -2);
        var firstcontentfulpaint = content['lighthouseResult']['audits'][
          'first-contentful-paint'
        ]['displayValue'].slice(0, -2);
        var firstmeaningfulpaint = content['lighthouseResult']['audits'][
          'first-meaningful-paint'
        ]['displayValue'].slice(0, -2);
        var cumulativelayoutshift =
          content['lighthouseResult']['audits']['cumulative-layout-shift'][
            'displayValue'
          ];
        var maxpotentialfid =
          content['lighthouseResult']['audits']['max-potential-fid'][
            'displayValue'
          ];
        var serverresponsetime = content['lighthouseResult']['audits'][
          'server-response-time'
        ]['displayValue'].slice(19, -3);
        var speedindex = content['lighthouseResult']['audits']['speed-index'][
          'displayValue'
        ].slice(0, -2);
      } else {
        var score = 'An error occured';
        var timetointeractive = 'An error occured';
        var largestcontentfulpaint = 'An error occured';
        var firstcontentfulpaint = 'An error occured';
        var firstmeaningfulpaint = 'An error occured';
        var cumulativelayoutshift = 'An error occured';
        var maxpotentialfid = 'An error occured';
        var serverresponsetime = 'An error occured';
        var speedindex = 'An error occured';
      }
    }

    var currentDate = new Date().toJSON().slice(0, 10).replace(/-/g, '/');

    array.push([
      largestcontentfulpaint,
      firstcontentfulpaint,
      cumulativelayoutshift,
      maxpotentialfid,
      serverresponsetime,
      score,
      timetointeractive,
      firstmeaningfulpaint,
      speedindex,
      currentDate,
      'complete',
    ]);
    Utilities.sleep(1000);
    return array;
  }
}

// Menu 3
//Log the values and clear PageSpeed Results data:
function runLog() {
  var columnNumberToWatch = 13; // column A = 1, B = 2, etc.
  var valueToWatch = 'complete';
  var sheetNameToMoveTheRowTo = 'Log';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Results').activate();
  var cell = sheet.getRange('M6:M');
  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
  var lastRow = sheet.getLastRow();

  var Avals = ss.getRange('M6:M').getValues();
  var Alast = Avals.filter(String).length;

  if (
    sheet.getName() != sheetNameToMoveTheRowTo &&
    cell.getColumn() == columnNumberToWatch &&
    cell.getValue().toLowerCase() == valueToWatch
  ) {
    var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    sheet
      .getRange(cell.getRow(), 2, Alast, sheet.getLastColumn())
      .copyTo(targetRange, type, false);
    sheet.getRange('C6:M').clearContent();
  }
}

// Menu 4
function resetSchedule() {
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  resetsuccess();
}

// used by resetSchedule
function resetsuccess() {
  Browser.msgBox('Success! - Report and Log Times Reset');
}

// Logging
// throw Error(stuff);
