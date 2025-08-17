/**
 * Main application controller
 * Coordinates between services and handles high-level application flow
 */

class SchedulerApp {
  constructor() {
    this.sheetService = new SheetService();
    this.discordService = new DiscordService();
    this.eventService = new EventSchedulingService(this.sheetService, this.discordService);
  }

  /**
   * Setup the response sheet with proper structure
   */
  async setupSheet() {
    try {
      const playerInfo = this.sheetService.getPlayerRoster();
      const playerNames = Object.keys(playerInfo);

      if (playerNames.length === 0) {
        const message = CONFIG.messages.ui.setupSheet.noPlayersFound
          .replace('{rosterSheetName}', CONFIG.rosterSheetName);
        SpreadsheetApp.getUi().alert(message);
        return;
      }

      // Check if sheet exists and confirm recreation
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const existingSheet = ss.getSheetByName(CONFIG.responseSheetName);

      if (existingSheet) {
        const confirmMessage = CONFIG.messages.ui.setupSheet.confirmRecreate
          .replace('{responseSheetName}', CONFIG.responseSheetName);
        const response = SpreadsheetApp.getUi().alert(
          'Setup Response Sheet',
          confirmMessage,
          SpreadsheetApp.getUi().ButtonSet.YES_NO
        );

        if (response !== SpreadsheetApp.getUi().Button.YES) {
          SpreadsheetApp.getUi().alert(CONFIG.messages.ui.setupSheet.setupCancelled);
          return;
        }
      }

      // Create the sheet
      this.sheetService.createResponseSheet(playerNames);

      // Apply formatting
      const formatter = new SheetFormatter();
      formatter.formatResponseSheet();

      // Create future date rows
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      UtilityService.createFutureDateRows(ss, today);

      // Send Discord notification
      await this.discordService.sendSheetSetupNotification();

      Logger.log('Sheet setup completed successfully');

    } catch (error) {
      Logger.log(`Error in setupSheet: ${error.toString()}`);
      SpreadsheetApp.getUi().alert(`Setup failed: ${error.toString()}`);
    }
  }

  /**
   * Handle cell edit events for immediate feedback
   */
  onCellEdit(e) {
    try {
      if (!e || !e.range) return;

      const range = e.range;
      const sheet = range.getSheet();
      const responseInfo = this.sheetService.getResponseSheetInfo();

      // Only process edits to the response sheet
      if (sheet.getName() !== CONFIG.responseSheetName) return;

      const row = range.getRow();
      const col = range.getColumn();

      // Only process player response columns
      if (row < CONFIG.firstDataRow ||
          col < CONFIG.firstPlayerColumn ||
          col >= responseInfo.statusColumnIndex) {
        return;
      }

      this.processPlayerResponse(row, col, e.value, responseInfo);

    } catch (error) {
      Logger.log(`Error in onCellEdit: ${error.toString()}`);
    }
  }

  /**
   * Process individual player response and update status
   */
  processPlayerResponse(row, col, value, responseInfo) {
    const sheet = responseInfo.sheet;

    // Validate the input if it's not empty
    if (value && value.toString().trim() !== '') {
      const validation = UtilityService.validateTimeFormat(value);
      if (!validation.isValid &&
          !['y', 'n', '?'].includes(value.toString().trim().toLowerCase())) {

        // Show error message
        const message = CONFIG.messages.ui.invalidTimeFormat.message
          .replace('{userInput}', value);
        SpreadsheetApp.getUi().alert(
          CONFIG.messages.ui.invalidTimeFormat.title,
          message,
          SpreadsheetApp.getUi().ButtonSet.OK
        );

        // Clear invalid input
        sheet.getRange(row, col).setValue('');
        return;
      }
    }

    // Update row status
    this.updateRowStatus(row, responseInfo);
  }

  /**
   * Update the status of a row based on player responses
   */
  updateRowStatus(row, responseInfo) {
    const sheet = responseInfo.sheet;
    const playerInfo = this.sheetService.getPlayerRoster();
    const numPlayers = Object.keys(playerInfo).length;

    // Get all responses for this row
    const responses = sheet.getRange(row, CONFIG.firstPlayerColumn, 1, responseInfo.numPlayerColumns)
      .getValues()[0];

    let yesCount = 0;
    let totalResponses = 0;

    responses.forEach((response, index) => {
      const playerName = responseInfo.allPlayerNames[index];
      if (playerName && playerInfo[playerName]) {
        if (response && response.toString().trim() !== '') {
          totalResponses++;
          if (UtilityService.isYesResponse(response)) {
            yesCount++;
          }
        }
      }
    });

    // Determine status
    let status;
    if (totalResponses === numPlayers && yesCount === numPlayers) {
      status = CONFIG.messages.status.readyForScheduling;
    } else if (totalResponses < numPlayers) {
      status = CONFIG.messages.status.awaitingResponses;
    } else {
      status = CONFIG.messages.status.awaitingResponses; // Some said no or maybe
    }

    // Update status column
    sheet.getRange(row, responseInfo.statusColumnIndex).setValue(status);
  }

  /**
   * Run daily scheduling check
   */
  async runDailyCheck() {
    try {
      Logger.log('Starting daily scheduling check');
      await this.eventService.processEvents('daily');
      Logger.log('Daily check completed');
    } catch (error) {
      Logger.log(`Error in daily check: ${error.toString()}`);
    }
  }

  /**
   * Run bi-weekly scheduling process
   */
  async runBiWeeklyCheck() {
    try {
      Logger.log('Starting bi-weekly scheduling check');
      await this.eventService.processEvents('biWeekly');

      // Also run maintenance tasks
      await this.runMaintenanceTasks();

      Logger.log('Bi-weekly check completed');
    } catch (error) {
      Logger.log(`Error in bi-weekly check: ${error.toString()}`);
    }
  }

  /**
   * Run maintenance tasks (archiving, future date creation)
   */
  async runMaintenanceTasks() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Archive old data
      UtilityService.archiveOldData(ss);

      // Create future dates
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      UtilityService.createFutureDateRows(ss, today);

      Logger.log('Maintenance tasks completed');
    } catch (error) {
      Logger.log(`Error in maintenance tasks: ${error.toString()}`);
    }
  }

  /**
   * Format response sheet
   */
  formatResponseSheet() {
    try {
      const formatter = new SheetFormatter();
      formatter.formatResponseSheet();
      Logger.log('Response sheet formatting completed');
    } catch (error) {
      Logger.log(`Error formatting response sheet: ${error.toString()}`);
    }
  }

  /**
   * Format archive sheet
   */
  formatArchiveSheet() {
    try {
      const formatter = new SheetFormatter();
      formatter.formatArchiveSheet();
      Logger.log('Archive sheet formatting completed');
    } catch (error) {
      Logger.log(`Error formatting archive sheet: ${error.toString()}`);
    }
  }
}

/**
 * Simplified sheet formatting class
 */
class SheetFormatter {
  formatResponseSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.responseSheetName);
    if (!sheet) {
      Logger.log(`Response sheet '${CONFIG.responseSheetName}' not found`);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < CONFIG.firstDataRow || lastCol < CONFIG.firstPlayerColumn) {
      Logger.log('Insufficient data to format response sheet');
      return;
    }

    this.applyHeaderFormatting(sheet, lastCol);
    this.applyDataFormatting(sheet, lastRow, lastCol);
    this.setupDataValidation(sheet, lastRow);
  }

  formatArchiveSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.archiveSheetName);
    if (!sheet) {
      Logger.log(`Archive sheet '${CONFIG.archiveSheetName}' not found`);
      return;
    }

    // Apply basic formatting
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow >= 1 && lastCol >= 1) {
      // Header formatting
      sheet.getRange(1, 1, 1, lastCol)
           .setBackground('#4285f4')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
    }
  }

  applyHeaderFormatting(sheet, lastCol) {
    const headerRange = sheet.getRange(CONFIG.headerRow, 1, 1, lastCol);
    headerRange.setBackground('#4285f4')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setFontSize(12)
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');

    // Special formatting for key columns
    sheet.getRange(CONFIG.headerRow, CONFIG.dateColumn).setBackground('#1a73e8');
  }

  applyDataFormatting(sheet, lastRow, lastCol) {
    // Date column formatting
    const dateRange = sheet.getRange(CONFIG.firstDataRow, CONFIG.dateColumn, lastRow - CONFIG.firstDataRow + 1, 1);
    dateRange.setBackground('#f8f9fa')
             .setFontWeight('bold')
             .setHorizontalAlignment('center')
             .setNumberFormat('yyyy.mm.dd');

    // Player columns formatting
    const headers = sheet.getRange(CONFIG.headerRow, 1, 1, lastCol).getValues()[0];
    const statusColIndex = headers.indexOf(CONFIG.statusColumnName) + 1;

    if (statusColIndex > 0) {
      const playerEndCol = statusColIndex - 1;

      for (let col = CONFIG.firstPlayerColumn; col <= playerEndCol; col++) {
        const playerRange = sheet.getRange(CONFIG.firstDataRow, col, lastRow - CONFIG.firstDataRow + 1, 1);
        playerRange.setHorizontalAlignment('center')
                   .setVerticalAlignment('middle')
                   .setFontSize(11);
      }

      // Status column formatting
      const statusRange = sheet.getRange(CONFIG.firstDataRow, statusColIndex, lastRow - CONFIG.firstDataRow + 1, 1);
      statusRange.setBackground('#f8f9fa')
                 .setFontSize(10)
                 .setHorizontalAlignment('center')
                 .setVerticalAlignment('middle');
    }
  }

  setupDataValidation(sheet, lastRow) {
    // Add data validation for player response columns
    const headers = sheet.getRange(CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.indexOf(CONFIG.statusColumnName) + 1;

    if (statusColIndex > 0 && lastRow >= CONFIG.firstDataRow) {
      const playerEndCol = statusColIndex - 1;
      const validationRange = sheet.getRange(
        CONFIG.firstDataRow, CONFIG.firstPlayerColumn,
        lastRow - CONFIG.firstDataRow + 1,
        playerEndCol - CONFIG.firstPlayerColumn + 1
      );

      const validation = SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .setHelpText(CONFIG.messages.validation.playerResponseHelp)
        .build();

      validationRange.setDataValidation(validation);
    }
  }
}
