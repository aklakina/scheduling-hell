/**
 * Service layer for Google Sheets operations
 * Centralizes all sheet access and provides clean interfaces
 */

class SheetService {
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  /**
   * Get sheet by name with error handling
   */
  getSheet(sheetName) {
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }
    return sheet;
  }

  /**
   * Get player roster data
   */
  getPlayerRoster() {
    const rosterSheet = this.getSheet(CONFIG.rosterSheetName);
    if (rosterSheet.getLastRow() < 2) {
      return {};
    }

    const rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 3).getValues();
    const playerInfo = {};

    rosterData.forEach(row => {
      if (row[0]) {
        playerInfo[row[0]] = {
          discordHandle: row[1] || "",
          allowMention: Boolean(row[2])
        };
      }
    });

    return playerInfo;
  }

  /**
   * Get campaign details
   */
  getCampaignDetails() {
    const campaignSheet = this.getSheet(CONFIG.campaignDetailsSheetName);
    const campaignData = campaignSheet.getRange(2, 1, 1, 2).getValues()[0];

    return {
      eventTitle: campaignData[0],
      eventLink: campaignData[1]
    };
  }

  /**
   * Get response sheet structure info
   */
  getResponseSheetInfo() {
    const responseSheet = this.getSheet(CONFIG.responseSheetName);
    const headers = responseSheet.getRange(CONFIG.headerRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    const statusColumnIndex = headers.indexOf(CONFIG.statusColumnName) + 1;

    if (statusColumnIndex === 0) {
      throw new Error('Status column not found');
    }

    const numPlayerColumns = statusColumnIndex - CONFIG.firstPlayerColumn;
    if (numPlayerColumns <= 0) {
      throw new Error('No player columns found');
    }

    const allPlayerNames = responseSheet.getRange(CONFIG.headerRow, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

    return {
      sheet: responseSheet,
      headers,
      statusColumnIndex,
      numPlayerColumns,
      allPlayerNames
    };
  }

  /**
   * Get event data within date range
   */
  getEventsInRange(startDate, endDate) {
    const responseInfo = this.getResponseSheetInfo();
    const { sheet, statusColumnIndex } = responseInfo;

    const allData = sheet.getRange(CONFIG.firstDataRow, 1, sheet.getLastRow() - CONFIG.firstDataRow + 1, statusColumnIndex).getValues();

    return allData.map((row, index) => ({
      rowData: row,
      rowIndex: CONFIG.firstDataRow + index
    })).filter(item => {
      const dateValue = item.rowData[CONFIG.dateColumn - 1];
      if (!dateValue) return false;

      const eventDate = new Date(dateValue);
      const status = item.rowData[statusColumnIndex - 1] || '';

      return eventDate >= startDate &&
             eventDate < endDate &&
             !status.startsWith('Event created') &&
             !status.startsWith('Superseded');
    });
  }

  /**
   * Get responses for a specific event
   */
  getEventResponses(rowIndex, numPlayerColumns) {
    const responseSheet = this.getSheet(CONFIG.responseSheetName);
    return responseSheet.getRange(rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();
  }

  /**
   * Update event status
   */
  updateEventStatus(rowIndex, status) {
    const responseInfo = this.getResponseSheetInfo();
    responseInfo.sheet.getRange(rowIndex, responseInfo.statusColumnIndex).setValue(status);
  }

  /**
   * Create response sheet with proper structure
   */
  createResponseSheet(playerNames) {
    let sheet = this.spreadsheet.getSheetByName(CONFIG.responseSheetName);

    if (sheet) {
      this.spreadsheet.deleteSheet(sheet);
    }

    sheet = this.spreadsheet.insertSheet(CONFIG.responseSheetName);

    const headers = [
      'Date',
      'Day',
      ...playerNames,
      'Today',
      'Status'
    ];

    sheet.getRange(CONFIG.headerRow, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }

  /**
   * Get or create archive sheet
   */
  getOrCreateArchiveSheet() {
    let archiveSheet = this.spreadsheet.getSheetByName(CONFIG.archiveSheetName);

    if (!archiveSheet) {
      const responseSheet = this.getSheet(CONFIG.responseSheetName);
      archiveSheet = this.spreadsheet.insertSheet(CONFIG.archiveSheetName);

      const headers = responseSheet.getRange(CONFIG.headerRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
      archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    return archiveSheet;
  }
}
