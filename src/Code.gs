/**
 * @OnlyCurrentDoc
 *
 * Main entry points for the scheduling system.
 * This file now serves as a thin interface layer that delegates to the service architecture.
 */

// Global app instance
let schedulerApp;

/**
 * Initialize the application
 */
function initializeApp() {
  if (!schedulerApp) {
    schedulerApp = new SchedulerApp();
  }
  return schedulerApp;
}

/**
 * Creates a custom menu in the spreadsheet UI to allow easy setup.
 * Runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scheduler')
    .addItem('Setup Sheet', 'setupSheet')
    .addSeparator()
    .addItem('🔄 Run Daily Check', 'dailySchedulingCheck')
    .addItem('🎯 Run Bi-Weekly Check', 'checkAndScheduleEvents')
    .addSeparator()
    .addItem('🎨 Format Response Sheet', 'formatResponseSheet')
    .addItem('🎨 Format Archive Sheet', 'formatArchiveSheet')
    .addToUi();
}

/**
 * Sets up the response sheet by creating the proper structure with roster-based columns.
 */
function setupSheet() {
  const app = initializeApp();
  app.setupSheet();
}

/**
 * Provides immediate UI feedback and updates the row's status column.
 * This should be triggered by an 'On edit' event.
 */
function onEditFeedback(e) {
  const app = initializeApp();
  app.onCellEdit(e);
}

/**
 * Daily scheduling check - lightweight processing for immediate scheduling needs
 */
function dailySchedulingCheck() {
  const app = initializeApp();
  app.runDailyCheck();
}

/**
 * Main scheduling function - processes events, sends notifications, and performs maintenance
 * This should be run on a time-based trigger (e.g., bi-weekly).
 */
function checkAndScheduleEvents() {
  const app = initializeApp();
  app.runBiWeeklyCheck();
}

/**
 * Test function for bi-weekly mode
 */
function testBiWeeklyMode() {
  Logger.log('Testing bi-weekly mode...');
  checkAndScheduleEvents();
}

/**
 * Applies comprehensive formatting to the response sheet
 */
function formatResponseSheet() {
  const app = initializeApp();
  app.formatResponseSheet();
}

/**
 * Applies formatting to the archive sheet
 */
function formatArchiveSheet() {
  const app = initializeApp();
  app.formatArchiveSheet();
}

// Legacy function aliases for backward compatibility
function onEdit(e) {
  onEditFeedback(e);
}

// Utility function to set up triggers programmatically if needed
function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditFeedback' || 
        trigger.getHandlerFunction() === 'checkAndScheduleEvents' ||
        trigger.getHandlerFunction() === 'dailySchedulingCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new triggers
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // On edit trigger
  ScriptApp.newTrigger('onEditFeedback')
    .spreadsheet(ss)
    .onEdit()
    .create();

  // Daily trigger
  ScriptApp.newTrigger('dailySchedulingCheck')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // Bi-weekly trigger (1st and 16th of each month)
  ScriptApp.newTrigger('checkAndScheduleEvents')
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .create();

  ScriptApp.newTrigger('checkAndScheduleEvents')
    .timeBased()
    .onMonthDay(16)
    .atHour(10)
    .create();

  Logger.log('Triggers set up successfully');
}
