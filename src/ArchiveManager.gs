/**
 * Archive and data management functionality
 */

/**
 * Archives old responses to the Archive sheet.
 * Keeps last week's data and archives older ones to maintain a clean active sheet.
 */
function archiveOldResponses(ss, processingStartDate) {
  const responseSheet = ss.getSheetByName(CONFIG.responseSheetName);
  let archiveSheet = ss.getSheetByName(CONFIG.archiveSheetName);

  if (!responseSheet) {
    Logger.log(`Error: Response sheet '${CONFIG.responseSheetName}' not found.`);
    return;
  }

  // Create archive sheet if it doesn't exist
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(CONFIG.archiveSheetName);
    // Copy headers from response sheet
    const headers = responseSheet.getRange(CONFIG.headerRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    Logger.log(`Created new archive sheet: ${CONFIG.archiveSheetName}`);
  }

  // Calculate archive threshold - keep last week's data (7 days before today)
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const archiveThreshold = new Date(today);
  archiveThreshold.setDate(today.getDate() - (CONFIG.weeksToKeepBeforeArchive * 7));

  // Get all data from response sheet
  const lastRow = responseSheet.getLastRow();
  if (lastRow < CONFIG.firstDataRow) {
    Logger.log('No data rows to process for archiving.');
    return;
  }

  const allData = responseSheet.getRange(CONFIG.firstDataRow, 1, lastRow - CONFIG.firstDataRow + 1, responseSheet.getLastColumn()).getValues();
  const rowsToArchive = [];

  // Identify rows to archive (dates older than archive threshold)
  allData.forEach((row, index) => {
    const dateValue = row[CONFIG.dateColumn - 1];
    if (dateValue) {
      const eventDate = new Date(dateValue);
      if (eventDate < archiveThreshold) {
        rowsToArchive.push({
          rowIndex: CONFIG.firstDataRow + index,
          data: row
        });
      }
    }
  });

  if (rowsToArchive.length === 0) {
    Logger.log('No old rows to archive.');
    return;
  }

  // Sort rows to archive by date (descending for archive sheet)
  rowsToArchive.sort((a, b) => {
    const dateA = new Date(a.data[CONFIG.dateColumn - 1]);
    const dateB = new Date(b.data[CONFIG.dateColumn - 1]);
    return dateB - dateA; // Descending order
  });

  // Add archived rows to the archive sheet (insert all at once to maintain descending order)
  if (rowsToArchive.length > 0) {
    // Insert the required number of rows after the header
    archiveSheet.insertRowsAfter(1, rowsToArchive.length);

    // Prepare the data array in the correct order
    const dataToInsert = rowsToArchive.map(item => item.data);

    // Insert all rows at once starting from row 2
    archiveSheet.getRange(2, 1, dataToInsert.length, dataToInsert[0].length).setValues(dataToInsert);
  }

  // Delete archived rows from response sheet (delete from bottom to top to maintain indices)
  rowsToArchive.sort((a, b) => b.rowIndex - a.rowIndex);
  rowsToArchive.forEach(item => {
    responseSheet.deleteRow(item.rowIndex);
  });

  Logger.log(`Archived ${rowsToArchive.length} old rows (older than ${archiveThreshold.toLocaleDateString()}) to '${CONFIG.archiveSheetName}' sheet.`);

  // Apply formatting to the archive sheet after archiving data
  try {
    formatArchiveSheet();
  } catch (error) {
    Logger.log(`Error formatting archive sheet: ${error.toString()}`);
  }
}

/**
 * Creates future date rows in the response sheet automatically.
 * Ensures there are always 2 months of future dates including today for scheduling.
 */
function createFutureDateRows(ss, today) {
  const responseSheet = ss.getSheetByName(CONFIG.responseSheetName);
  if (!responseSheet) {
    Logger.log(`Error: Response sheet '${CONFIG.responseSheetName}' not found.`);
    return;
  }

  // Calculate target end date for 2 months including today
  const targetEndDate = new Date(today);
  targetEndDate.setMonth(today.getMonth() + CONFIG.monthsToCreateAhead);

  // Find the last date in the sheet
  const lastRow = responseSheet.getLastRow();
  let lastDate = new Date(today.getTime() - (24 * 60 * 60 * 1000)); // Start from yesterday to ensure today is included

  if (lastRow >= CONFIG.firstDataRow) {
    // Look for the highest date in the sheet
    const dateRange = responseSheet.getRange(CONFIG.firstDataRow, CONFIG.dateColumn, lastRow - CONFIG.firstDataRow + 1, 1).getValues();
    dateRange.forEach(row => {
      const cellDate = new Date(row[0]);
      if (!isNaN(cellDate.getTime()) && cellDate > lastDate) {
        lastDate = cellDate;
      }
    });
  }

  // Create new daily dates starting from the next day after lastDate, up to target end date
  const newDates = [];
  let currentDate = new Date(lastDate);
  currentDate.setDate(lastDate.getDate() + 1);

  while (currentDate <= targetEndDate) {
    newDates.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + 1);
  }

  if (newDates.length === 0) {
    Logger.log('No new dates needed - sufficient future dates already exist.');
    return;
  }

  // Get the current sheet structure to build proper rows
  const headers = responseSheet.getRange(CONFIG.headerRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  const totalColumns = headers.length;

  // Find Today and Status column indices
  const todayColIndex = headers.findIndex(h => h.toString().includes('Today')) + 1;
  const statusColIndex = headers.findIndex(h => h.toString().includes('Status')) + 1;

  // Add new date rows to the response sheet
  newDates.forEach((date, index) => {
    const newRowData = new Array(totalColumns).fill('');
    const newRowIndex = lastRow + index + 1;

    newRowData[0] = date; // Date column (A)
    newRowData[1] = ''; // Day column - will set formula after appending

    // Add Today formula if Today column exists
    if (todayColIndex > 0) {
      newRowData[todayColIndex - 1] = ''; // Will set formula after appending
    }

    // Status column starts empty
    if (statusColIndex > 0) {
      newRowData[statusColIndex - 1] = '';
    }

    responseSheet.appendRow(newRowData);

    // Set formulas after the row is added
    responseSheet.getRange(newRowIndex, 2).setFormula("=TEXT(A" + newRowIndex + ";\"dddd\")"); // Day column formula

    if (todayColIndex > 0) {
      responseSheet.getRange(newRowIndex, todayColIndex).setFormula("=IF(TODAY()=A" + newRowIndex + ";\"<-----\";\"\")"); // Today formula
    }
  });

  Logger.log(`Created ${newDates.length} new date rows up to ${targetEndDate.toLocaleDateString()}.`);

  // Apply formatting to the newly created rows
  formatResponseSheet();
}
