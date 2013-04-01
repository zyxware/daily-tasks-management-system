/** 
 * Name: Daily Tasks Management System
 * Author: Anoop John, Zyxware Technologies(http://www.zyxware.com)
 * License: GNU GPL v3
 */

/** 
 * Create menu for the application
 */
function onOpen() {
  var menuEntries = [];
  menuEntries.push({name: "Send Reminders", functionName: "sendEmails"});
  menuEntries.push({name: "Archive Tasks", functionName: "archiveTasks"});
  menuEntries.push({name: "Add Queued Tasks", functionName: "addQueuedTasks"});
  menuEntries.push({name: "Save Config", functionName: "saveConfig"});
  menuEntries.push({name: "About App", functionName: "aboutApp"});
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Task Utils", menuEntries);
}

/** 
 * Send emails to people with pending tasks for the day
 */
function sendEmails() {
  // Run only on weekdays
  var d = new Date();
  // If the application is set to exclude weekends then don't send emails on weekends.
  if (getAppConfig('exclude_weekends') == 1 && (d.getDay() == 0 || d.getDay() == 6)) {
    return;
  }
  var users = getDataRange('Team').getValues();
  var tasks = getDataRange('Tasks').getValues();
  var pending_tasks = new Array();
  var today = getToday();
  // Find all users for which tasks are to be tracked
  for (var cur_user = 0; cur_user < users.length; cur_user++) {
    var user_name = users[cur_user][0];
    if (user_name != '') {
      pending_tasks[user_name] = new Array();
      pending_tasks[user_name]['firstname'] = users[cur_user][1];
      pending_tasks[user_name]['designation'] = users[cur_user][2];
      pending_tasks[user_name]['email'] = users[cur_user][3];
      pending_tasks[user_name]['tasks'] = new Array(); 
    }
  }
  // Find all pending tasks and compile against users
  for (var cur_task = 0; cur_task < tasks.length; cur_task++) {
    var task_date = formatDate(tasks[cur_task][0]);
    var task_status = tasks[cur_task][4];
    // Check only for tasks assigned for the day for non-empty rows and where tasks are not Done
    if (task_date != '' && task_date == today && task_status != 'Done') {
      var task_user = tasks[cur_task][3];
      // Bother collecting tasks only for rows with valid names
      if (pending_tasks[task_user] instanceof Array) {
        pending_tasks[task_user]['tasks'].push(cur_task); 
      }
    }
  }
  // Send notification about pending tasks to users.
  for (var cur_user = 0; cur_user < users.length; cur_user++) {
    var user_name = users[cur_user][0];
    // Send emails only to those users who have pending tasks
    if (user_name != '' && pending_tasks[user_name]['tasks'].length > 0) {
      var num_tasks = pending_tasks[user_name]['tasks'].length;
      var emailAddress = pending_tasks[user_name]['email'];
      var message = "Dear " + pending_tasks[user_name]['firstname'] + ",\n\n" +
        "You have " + pending_tasks[user_name]['tasks'].length + " " + formatPlural(num_tasks, "task that is", "tasks that are") + " left to be done today.\n\n" +
        "Please check \n\n" + 
        "https://docs.google.com/spreadsheet/ccc?key=" + ScriptProperties.getProperty('sheet_id') + "#gid=0\n\n" + 
        "to see the list of all tasks and mark finished tasks as 'Done'.\n\n" +
        "The following " + formatPlural(num_tasks, "is the pending task", "are the pending tasks") + " for today\n\n";
      var task_list = "";
      var num_tasks = 0;
      for (i = 0; i < pending_tasks[user_name]['tasks'].length; i++) {
        cur_task = pending_tasks[user_name]['tasks'][i];
        num_tasks++;
        task_list += num_tasks + ". " + tasks[cur_task][1];
        if (tasks[cur_task][2] != '') {
          task_list += ' (Project: ' + tasks[cur_task][2] + ')';
        }
        task_list += "\n";
      }
      message += task_list + "\n";
      message += "Remember to complete " + formatPlural(num_tasks, "this task", "these tasks") + " before EOD today.\n\n" +
        "Best Regards\n\n" +
        getAppConfig('sender_name') + "\n" +
        getAppConfig('sender_designation') + "\n";
      var subject = "High priority tasks for " + getToday() + " (" + pending_tasks[user_name]['tasks'].length + " pending " + formatPlural(num_tasks, "task", "tasks") + ")";
      MailApp.sendEmail(emailAddress, subject, message);
      // Browser.msgBox(emailAddress + "\\n" + subject + "\\n\\n" + message);
    }
  }
  Logger.log('%s: Sent emails.', getNow());
}

/** 
 * Archive tasks that have been completed to reduce clutter in the tasks sheet
 */
function archiveTasks() {
  var sa = getSpreadsheetApp();
  var s_a = sa.getSheetByName('Archive');
  var tasks = getDataRange('Tasks').getValues();
  var today = getToday();
  var cur_task, task_date, task_status, row;
  // Find all completed tasks older than today and move to archive
  for (cur_task = 0; cur_task < tasks.length; cur_task++) {
    task_date = formatDate(tasks[cur_task][0]);
    task_status = tasks[cur_task][4].replace(/^\s+|\s+$/g, '');
    // Check only for tasks assigned for the past days for non-empty rows and where tasks are Done
    if (task_date != '' && task_date < today && task_status == 'Done') {
      row = getDataRange('Tasks').offset(cur_task, 0, 1, tasks[cur_task].length);
      var val = row.getValues();
      // Add a row if there is data in all available rows.
      if (s_a.getMaxRows() == s_a.getLastRow()) {
        s_a.insertRowsAfter(s_a.getLastRow(), 1);
      }
      row.moveTo(s_a.getRange(s_a.getLastRow() + 1, 1));
    }
  }
  getDataRange('Tasks').sort(1);
  getDataRange('Archive').sort(1);
  Logger.log('%s: Archived tasks.', getNow());
}

/** 
 * Take today's tasks from the task queue and add to DailyTasks
 */
function addQueuedTasks() {
  var sa = getSpreadsheetApp();
  var s_t = sa.getSheetByName('Tasks');
  var tasks = getDataRange('Queue').getValues();
  var today = getToday();
  var cur_task, task_date, row;
  // Find all tasks in the queue for today and move to DailyTasks sheet
  for (cur_task = 0; cur_task < tasks.length; cur_task++) {
    task_date = formatDate(tasks[cur_task][0]);
    // Check only for tasks assigned for today for non-empty rows
    if (task_date != '' && task_date == today) {
      row = getDataRange('Queue').offset(cur_task, 0, 1, tasks[cur_task].length);
      var val = row.getValues();
      // Add a row if there is data in all available rows.
      if (s_t.getMaxRows() == s_t.getLastRow()) {
        s_t.insertRowsAfter(s_t.getLastRow(), 1);
      }
      row.moveTo(s_t.getRange(s_t.getLastRow() + 1, 1));
    }
  }
  getDataRange('Tasks').sort(1);
  getDataRange('Queue').sort(1);
  Logger.log('%s: Added queued tasks.', getNow());
}

/** 
 * Save the configuration and set the triggers
 */
function saveConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  // Save sheet id for use when running triggers
  ScriptProperties.setProperty('sheet_id', ssId);
  // Delete existing triggers
  var allTriggers = ScriptApp.getScriptTriggers();
  // Loop over all current triggers and delete them
  for(var i=0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  // Create new triggers
  var addQueuedTasks_time = getAppConfig('addQueuedTasks_time');
  ScriptApp.newTrigger("addQueuedTasks").timeBased().everyDays(1).atHour(addQueuedTasks_time).create();
  var archiveTasks_time = getAppConfig('archiveTasks_time');
  ScriptApp.newTrigger("archiveTasks").timeBased().everyDays(1).atHour(archiveTasks_time).create();
  var sendEmails_time = getAppConfig('sendEmails_time');
  if (!(sendEmails_time instanceof Array)) {
    sendEmails_time = [sendEmails_time];
  }
  for (var i=0; i < sendEmails_time.length; i++) {
    ScriptApp.newTrigger("sendEmails").timeBased().everyDays(1).atHour(sendEmails_time[i]).create();
  }
  Logger.log('%s: Saved configuration.', getNow());
  Browser.msgBox('Daily Tasks Management System ', 'Saved configuration and created scheduled tasks.', Browser.Buttons.OK);
}

/** 
 * Get the spreadsheet application corresponding to the current spreadsheet 
 * Abstraction required for running the script via triggers.
 */
function getSpreadsheetApp() {
  if (this.app != null) {
    return this.app;
  }
  var ssId = ScriptProperties.getProperty('sheet_id');
  if (ssId == null) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss != null) {
      ssId = ss.getId();
    }
  }
  this.app = SpreadsheetApp.openById(ssId);
  return this.app;
}

/** 
 * Function to read configuration from inside the Config sheet
 */
function getAppConfig(key) {
  if (this.config != null) {
    return this.config[key];
  }
  var sa = getSpreadsheetApp();
  var configVals = getDataRange('Config').getValues();
  this.config = new Array();
  // Browser.msgBox(config);
  var cur_item, k, v;
  for (cur_item = 0; cur_item < configVals.length; cur_item++) {
    k = configVals[cur_item][0];
    v = configVals[cur_item][1];
    if (this.config[k] == undefined) {
      this.config[k] = v;
    }
    else {
      if (!(v instanceof Array)) {
        this.config[k] = [this.config[k]];
      }
      this.config[k].push(v);
    }
  }
  return this.config[key];
}

/** 
 * Utility function to get data range from a sheet excluding the number of
 * frozen rows which is assumed to be the header. The function will return
 * only the rows that has some data in them. The function will return at 
 * least 1 row even if all rows in the sheet are empty.
 */
function getDataRange(sheetName) {
  var sa = getSpreadsheetApp();
  var s = sa.getSheetByName(sheetName);
  var range = s.getDataRange();
  // Browser.msgBox(range.getValues());
  var frozen_rows = s.getFrozenRows();
  var num_rows = range.getNumRows() - frozen_rows;
  if (num_rows == 0) {
    num_rows = 1;
  }
  // Browser.msgBox(sheetName + "\\n" + frozen_rows + " rows frozen \\nRows = " + range.getNumRows() + "\\n" + "Columsn = " + range.getNumColumns());
  var new_range = range.offset(frozen_rows, 0, num_rows, range.getNumColumns());
  // Trim empty rows and the end of the range
  var cur_row, vals;
  num_rows = new_range.getNumRows();
  vals = new_range.getValues();
  for (cur_row = num_rows - 1; cur_row >= 0; cur_row--) {
    if (vals[cur_row].join('').replace(/^\s+|\s+$/g, '') != "") {
      break;
    }
  }
  num_rows = cur_row + 1;
  // Ensure at least one row in the range if the whole range is empty
  if (num_rows == 0) {
    num_rows = 1;
  }
  var new_range2 = new_range.offset(0, 0, num_rows, range.getNumColumns());
  vals = new_range2.getValues();
  // Browser.msgBox(new_range.getValues());
  return new_range;
}

/** 
 * Show a popup with information about the application
 */
function aboutApp() {
  renderAboutDialog();
  //Browser.msgBox('Daily Tasks Management System ', 'A simple task management application to manage tasks allocated to a small team of people. Application developed and maintained by Zyxware Technologies. You can get support and the latest version from http://www.zyxware.com.', Browser.Buttons.OK);
}

/**
 * Render thea about box.
 * Ref: https://developers.google.com/apps-script/articles/twitter_tutorial
 */
function renderAboutDialog() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication();
  
  var helpLabel = app.createLabel(
      "A simple task management application to manage tasks allocated to a small team of " 
      + "people. Application developed and maintained by Zyxware Technologies. You can get " 
      + "support and the latest version from http://www.zyxware.com. ");
  helpLabel.setStyleAttribute("text-align", "left");
  
  var popupPanel = app.createPopupPanel();
  popupPanel.setPixelSize(340, 100);
  popupPanel.add(helpLabel);
  popupPanel.setStyleAttribute("border-width", "0px");
  app.add(popupPanel);
  app.setTitle("About Daily Task Management System");
  app.setHeight(100).setWidth(350);
  doc.show(app);
}

/** 
 * Utility function to return today as a string of the form yyyy-mm-dd
 */
function getToday() {
  if (this.today != null) {
    return this.today
  }
  this.today = formatDate(new Date());
  return this.today;
}

/** 
 * Utility function to return now as a string of the form yyyy-mm-dd
 */
function getNow() {
  var sa = getSpreadsheetApp();
  return Utilities.formatDate(new Date(), sa.getSpreadsheetTimeZone(), "yyyy-MM-dd hh:mm:ss");
}

/** 
 * Utility function to return a date as a string of the form yyyy-mm-dd
 */
function formatDate(d) {
  var sa = getSpreadsheetApp();
  return Utilities.formatDate(new Date(d), sa.getSpreadsheetTimeZone(), "yyyy-MM-dd");
}

/** 
 * Utility function to test if the passed date equal to today
 */
function isToday(val) {
  today = Date();
  d = Date.parse(val);
  if (d.getDate() == today.getDate() && d.getMonth() == today.getMonth() && d.getFullYear() == today.getFullYear()) {
    return true;
  }
  return false;
}

/** 
 * Utility function to return singular / plurar versions of text based on a count
 */
function formatPlural(count, singular, plural) {
  if (count == 1) {
    return singular;
  }
  return plural;
}

/** 
 * Clear the logs
 */
function clearLogs() {
  Logger.clear();
}

