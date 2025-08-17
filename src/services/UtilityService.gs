/**
 * Centralized utility service for common operations
 * Consolidates helper functions used across the application
 */

class UtilityService {
  /**
   * Time format validation and parsing
   */
  static validateTimeFormat(str) {
    const rangeRegex = /^(\d{1,2}(?::\d{2})?)-(\d{1,2}(?::\d{2})?)$/;
    const singleTimeRegex = /^(\d{1,2}(?::\d{2})?)$/;

    str = String(str).replace(/\s/g, '');

    const rangeMatch = str.match(rangeRegex);
    const singleMatch = str.match(singleTimeRegex);

    return {
      isValid: Boolean(rangeMatch || singleMatch),
      rangeMatch,
      singleMatch,
      type: rangeMatch ? 'range' : singleMatch ? 'single' : 'invalid'
    };
  }

  /**
   * Parse time range into start and end dates
   */
  static parseTimeRange(timeStr, baseDate) {
    const validation = this.validateTimeFormat(timeStr);

    if (!validation.isValid) {
      return { start: null, end: null };
    }

    const createDate = (timePart) => {
      const newDate = new Date(baseDate);
      const parts = timePart.split(':');
      const hours = parseInt(parts[0], 10);
      const minutes = parts.length > 1 ? parseInt(parts[1], 10) : 0;

      if (isNaN(hours) || isNaN(minutes) ||
          hours > 23 || minutes > 59 ||
          hours < 0 || minutes < 0) {
        return null;
      }

      newDate.setHours(hours, minutes, 0, 0);
      return newDate;
    };

    if (validation.rangeMatch) {
      const [, startTime, endTime] = validation.rangeMatch;
      return {
        start: createDate(startTime),
        end: createDate(endTime)
      };
    } else if (validation.singleMatch) {
      const [, time] = validation.singleMatch;
      return {
        start: createDate(time),
        end: null // All day from this time
      };
    }

    return { start: null, end: null };
  }

  /**
   * Calculate time intersection for multiple time ranges
   */
  static calculateTimeIntersection(timeRanges, baseDate) {
    if (!timeRanges || timeRanges.length === 0) {
      return { start: undefined, end: undefined };
    }

    let intersectionStart = null;
    let intersectionEnd = null;

    for (const timeRange of timeRanges) {
      const { start, end } = this.parseTimeRange(timeRange, baseDate);

      if (!start) continue; // Skip invalid time ranges

      const rangeEnd = end || new Date(baseDate.getTime() + 24 * 60 * 60 * 1000); // End of day if no end time

      if (intersectionStart === null) {
        intersectionStart = start;
        intersectionEnd = rangeEnd;
      } else {
        intersectionStart = new Date(Math.max(intersectionStart.getTime(), start.getTime()));
        intersectionEnd = new Date(Math.min(intersectionEnd.getTime(), rangeEnd.getTime()));
      }

      // If intersection becomes invalid, return early
      if (intersectionStart >= intersectionEnd) {
        return { start: undefined, end: undefined };
      }
    }

    return {
      start: intersectionStart,
      end: intersectionEnd
    };
  }

  /**
   * Generate combinations of specified size from array
   */
  static generateCombinations(array, size) {
    if (size > array.length || size <= 0) {
      return [];
    }

    if (size === 1) {
      return array.map(item => [item]);
    }

    const combinations = [];

    for (let i = 0; i <= array.length - size; i++) {
      const smallerCombinations = this.generateCombinations(array.slice(i + 1), size - 1);
      smallerCombinations.forEach(combination => {
        combinations.push([array[i], ...combination]);
      });
    }

    return combinations;
  }

  /**
   * Get week number for date grouping
   */
  static getWeekNumber(date) {
    const startOfYear = new Date(date.getFullYear(), 0, 1);
    const pastDaysOfYear = (date - startOfYear) / 86400000;
    return Math.ceil((pastDaysOfYear + startOfYear.getDay() + 1) / 7);
  }

  /**
   * Calculate duration in hours between two dates
   */
  static calculateDurationHours(startDate, endDate) {
    if (!startDate || !endDate) return 0;
    return (endDate.getTime() - startDate.getTime()) / 3600000;
  }

  /**
   * Format date for logging and display
   */
  static formatDate(date, includeTime = false) {
    if (!date) return 'Invalid Date';

    if (includeTime) {
      return `${date.toLocaleDateString()} ${date.toLocaleTimeString()}`;
    }
    return date.toLocaleDateString();
  }

  /**
   * Create future date rows for scheduling
   */
  static createFutureDateRows(spreadsheet, startDate) {
    const sheet = spreadsheet.getSheetByName(CONFIG.responseSheetName);
    if (!sheet) {
      throw new Error('Response sheet not found');
    }

    const endDate = new Date(startDate);
    endDate.setMonth(startDate.getMonth() + CONFIG.monthsToCreateAhead);

    const datesToCreate = [];
    const currentDate = new Date(startDate);

    while (currentDate <= endDate) {
      datesToCreate.push({
        date: new Date(currentDate),
        dayName: currentDate.toLocaleDateString('en-US', { weekday: 'long' })
      });
      currentDate.setDate(currentDate.getDate() + 1);
    }

    if (datesToCreate.length === 0) return;

    // Get current structure info
    const headers = sheet.getRange(CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const numColumns = headers.length;
    const startRow = sheet.getLastRow() + 1;

    // Prepare data for bulk insert
    const rowData = datesToCreate.map(dateInfo => {
      const row = new Array(numColumns).fill('');
      row[CONFIG.dateColumn - 1] = dateInfo.date;

      // Find Day column and set it
      const dayColumnIndex = headers.findIndex(h => h.toString().includes('Day'));
      if (dayColumnIndex !== -1) {
        row[dayColumnIndex] = dateInfo.dayName;
      }

      return row;
    });

    // Bulk insert all rows
    sheet.getRange(startRow, 1, rowData.length, numColumns).setValues(rowData);

    Logger.log(`Created ${datesToCreate.length} future date rows starting from ${this.formatDate(startDate)}`);
  }

  /**
   * Archive old data based on configuration
   */
  static archiveOldData(spreadsheet) {
    const responseSheet = spreadsheet.getSheetByName(CONFIG.responseSheetName);
    const archiveSheet = spreadsheet.getSheetByName(CONFIG.archiveSheetName) ||
                        this.createArchiveSheet(spreadsheet);

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const archiveThreshold = new Date(today);
    archiveThreshold.setDate(today.getDate() - (CONFIG.weeksToKeepBeforeArchive * 7));

    const lastRow = responseSheet.getLastRow();
    if (lastRow < CONFIG.firstDataRow) return;

    const allData = responseSheet.getRange(
      CONFIG.firstDataRow, 1,
      lastRow - CONFIG.firstDataRow + 1,
      responseSheet.getLastColumn()
    ).getValues();

    const rowsToArchive = [];

    allData.forEach((row, index) => {
      const dateValue = row[CONFIG.dateColumn - 1];
      if (dateValue && new Date(dateValue) < archiveThreshold) {
        rowsToArchive.push({
          rowIndex: CONFIG.firstDataRow + index,
          data: row
        });
      }
    });

    if (rowsToArchive.length === 0) {
      Logger.log('No rows to archive');
      return;
    }

    // Move to archive
    const archiveData = rowsToArchive.map(row => row.data);
    const archiveStartRow = archiveSheet.getLastRow() + 1;

    archiveSheet.getRange(
      archiveStartRow, 1,
      archiveData.length,
      archiveData[0].length
    ).setValues(archiveData);

    // Delete from response sheet (in reverse order to maintain indices)
    rowsToArchive.reverse().forEach(row => {
      responseSheet.deleteRow(row.rowIndex);
    });

    Logger.log(`Archived ${rowsToArchive.length} old rows`);
  }

  /**
   * Create archive sheet with proper headers
   */
  static createArchiveSheet(spreadsheet) {
    const responseSheet = spreadsheet.getSheetByName(CONFIG.responseSheetName);
    const archiveSheet = spreadsheet.insertSheet(CONFIG.archiveSheetName);

    const headers = responseSheet.getRange(CONFIG.headerRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    Logger.log(`Created archive sheet: ${CONFIG.archiveSheetName}`);
    return archiveSheet;
  }

  /**
   * Validate and sanitize user input
   */
  static sanitizeInput(input) {
    if (!input) return '';
    return String(input).trim();
  }

  /**
   * Check if value represents a "yes" response
   */
  static isYesResponse(response) {
    if (!response) return false;
    const responseStr = String(response).trim().toLowerCase();
    return responseStr === CONFIG.responses.yes || this.validateTimeFormat(responseStr).isValid;
  }

  /**
   * Check if value represents a "maybe" response
   */
  static isMaybeResponse(response) {
    if (!response) return false;
    return String(response).trim().toLowerCase() === CONFIG.responses.maybe;
  }

  /**
   * Check if value represents an empty/no response
   */
  static isEmptyResponse(response) {
    return !response || String(response).trim() === '';
  }
}
