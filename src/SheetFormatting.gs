/**
 * Sheet formatting and styling functionality
 */

/**
 * Applies comprehensive formatting to the response sheet for better user experience
 */
function formatResponseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.responseSheetName);
  if (!sheet) {
    Logger.log(`Error: Response sheet '${CONFIG.responseSheetName}' not found.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < CONFIG.firstDataRow || lastCol < CONFIG.firstPlayerColumn) {
    Logger.log('Insufficient data to format response sheet.');
    return;
  }

  // Get headers to identify columns
  const headers = sheet.getRange(CONFIG.headerRow, 1, 1, lastCol).getValues()[0];

  // Clear existing formatting
  sheet.getRange(1, 1, lastRow, lastCol).clearFormat();

  // --- Header Formatting ---
  const headerRange = sheet.getRange(CONFIG.headerRow, 1, 1, lastCol);
  headerRange.setBackground('#4285f4')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setFontSize(12)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');

  // Date column header special formatting
  sheet.getRange(CONFIG.headerRow, CONFIG.dateColumn)
       .setBackground('#1a73e8');

  // Status column special formatting
  const statusColIndex = headers.indexOf(CONFIG.statusColumnName) + 1;
  if (statusColIndex > 0) {
    sheet.getRange(CONFIG.headerRow, statusColIndex)
         .setBackground('#1a73e8');
  }

  // Find and format Today column
  const todayColIndex = headers.findIndex(h => h.toString().includes('Today')) + 1;
  if (todayColIndex > 0) {
    sheet.getRange(CONFIG.headerRow, todayColIndex)
         .setBackground('#1a73e8');
  }

  // Find and format Day column
  const dayColIndex = headers.findIndex(h => h.toString().includes('Day')) + 1;
  if (dayColIndex > 0) {
    sheet.getRange(CONFIG.headerRow, dayColIndex)
         .setBackground('#1a73e8');
  }

  // Calculate player column range
  const playerStartCol = CONFIG.firstPlayerColumn;
  const playerEndCol = statusColIndex > 0 ? statusColIndex - 1 : lastCol;

  // --- Data Row Formatting ---
  for (let row = CONFIG.firstDataRow; row <= lastRow; row++) {
    // Date column formatting - ensure yyyy.MM.dd format
    const dateCell = sheet.getRange(row, CONFIG.dateColumn);
    dateCell.setBackground('#f8f9fa')
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setNumberFormat('yyyy.mm.dd'); // yyyy.MM.dd format

    // Player response columns
    for (let col = playerStartCol; col <= playerEndCol; col++) {
      const cell = sheet.getRange(row, col);
      cell.setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setFontSize(11);
    }

    // Status column formatting
    if (statusColIndex > 0) {
      const statusCell = sheet.getRange(row, statusColIndex);
      statusCell.setBackground('#f8f9fa')
             .setFontSize(10)
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');
    }
  }

  // --- Data Validation for Player Columns ---
  if (lastRow >= CONFIG.firstDataRow) {
    // Get roster data to identify actual player columns
    const rosterSheet = ss.getSheetByName(CONFIG.rosterSheetName);
    if (rosterSheet && rosterSheet.getLastRow() >= 2) {
      const rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 1).getValues();
      const playerNames = rosterData.map(row => row[0]).filter(name => name && name.toString().trim() !== '');

      // Apply validation only to columns that match player names in the roster
      for (let col = playerStartCol; col <= playerEndCol; col++) {
        const headerValue = sheet.getRange(CONFIG.headerRow, col).getValue();
        const cleanHeaderName = headerValue ? headerValue.toString().replace(/^👤\s*/, '').trim() : '';

        // Only apply validation if this column header matches a player in the roster
        if (playerNames.includes(cleanHeaderName)) {
          const playerColumnRange = sheet.getRange(CONFIG.firstDataRow, col,
                                                  lastRow - CONFIG.firstDataRow + 1, 1);

          // Create dropdown with emojis for quick selection but allow custom entries
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(['Y', 'N', '?', ''], true)
            .setAllowInvalid(true) // Allow time ranges and custom entries
            .setHelpText(CONFIG.messages.validation.playerResponseHelp)
            .build();

          playerColumnRange.setDataValidation(rule);
        }
      }
    }
  }

  // --- Conditional Formatting Rules ---
  addConditionalFormattingRules(sheet, playerStartCol, playerEndCol, statusColIndex);

  // --- Column Widths ---
  sheet.setColumnWidth(CONFIG.dateColumn, 150);
  for (let col = playerStartCol; col <= playerEndCol; col++) {
    sheet.setColumnWidth(col, 80);
  }
  if (statusColIndex > 0) {
    sheet.setColumnWidth(statusColIndex, 200);
  }

  // Freeze header row and date column
  sheet.setFrozenRows(CONFIG.headerRow);
  sheet.setFrozenColumns(CONFIG.dateColumn);

  Logger.log('Response sheet formatting applied successfully.');
}

/**
 * Applies conditional formatting rules to enhance visual feedback
 */
function addConditionalFormattingRules(sheet, playerStartCol, playerEndCol, statusColIndex) {
  // Clear existing conditional formatting
  sheet.clearConditionalFormatRules();

  const rules = [];
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.firstDataRow) return;

  // Player response conditional formatting
  const playerRange = sheet.getRange(CONFIG.firstDataRow, playerStartCol,
                                   lastRow - CONFIG.firstDataRow + 1,
                                   playerEndCol - playerStartCol);

  // Yes responses (green) - matches cells containing 'y'
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Y')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([playerRange])
    .build());

  // No responses (red) - matches cells containing 'n'
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('N')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([playerRange])
    .build());

  // Maybe responses (yellow) - matches cells containing '?'
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('?')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([playerRange])
    .build());

  // Time range responses (blue) - matches cells that start with a number (time ranges)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH(INDIRECT(ADDRESS(ROW();COLUMN())); "^[\\d:]+(-[\\d:]+)?$")')
    .setBackground('#cce5ff')
    .setFontColor('#0056b3')
    .setRanges([playerRange])
    .build());

  // Status column conditional formatting
  if (statusColIndex > 0) {
    const statusRange = sheet.getRange(CONFIG.firstDataRow, statusColIndex,
                                     lastRow - CONFIG.firstDataRow + 1, 1);

    // Ready for scheduling (bright green)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Ready for scheduling')
      .setBackground('#28a745')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build());

    // Event created (success green)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Event created')
      .setBackground('#20c997')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build());

    // Cancelled (red)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Cancelled')
      .setBackground('#dc3545')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build());

    // Failed (orange)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Failed')
      .setBackground('#fd7e14')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build());

    // Awaiting responses (yellow)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Awaiting responses')
      .setBackground('#ffc107')
      .setFontColor('#212529')
      .setRanges([statusRange])
      .build());

    // Reminder sent (light blue)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Reminder sent')
      .setBackground('#17a2b8')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build());

    // Superseded (gray)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Superseded')
      .setBackground('#6c757d')
      .setFontColor('#ffffff')
      .setRanges([statusRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
}

/**
 * Applies formatting to the archive sheet for historical data viewing
 */
function formatArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName(CONFIG.archiveSheetName);

  if (!archiveSheet) {
    Logger.log('Archive sheet not found, skipping formatting.');
    return;
  }

  const lastRow = archiveSheet.getLastRow();
  let lastCol = archiveSheet.getLastColumn();

  // --- Remove Today column if it exists ---
  const headersRange = archiveSheet.getRange(CONFIG.headerRow, 1, 1, lastCol);
  const headers = headersRange.getValues()[0];
  const todayColIndex = headers.findIndex(h => h.toString().includes('Today')) + 1;
  if (todayColIndex > 0) {
    archiveSheet.deleteColumn(todayColIndex);
    // After deleting the column, we need to refetch lastCol and headers
    lastCol = archiveSheet.getLastColumn();
  }

  if (lastRow < 2 || lastCol < CONFIG.firstPlayerColumn) {
    Logger.log('Insufficient data to format archive sheet.');
    return;
  }

  // Clear existing formatting
  archiveSheet.getRange(1, 1, lastRow, lastCol).clearFormat();

  // --- Header Formatting (darker theme for archive) ---
  const headerRange = archiveSheet.getRange(CONFIG.headerRow, 1, 1, lastCol);
  headerRange.setBackground('#343a40')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setFontSize(12)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');

  // Special header formatting
  archiveSheet.getRange(CONFIG.headerRow, CONFIG.dateColumn)
             .setBackground('#212529')
             .setValue(CONFIG.messages.archive.dateColumnHeader);

  const updatedHeaders = archiveSheet.getRange(CONFIG.headerRow, 1, 1, lastCol).getValues()[0];
  const statusColIndex = updatedHeaders.indexOf(CONFIG.statusColumnName) > -1
    ? updatedHeaders.indexOf(CONFIG.statusColumnName) + 1
    : updatedHeaders.indexOf(CONFIG.messages.archive.statusColumnHeader) + 1;

  if (statusColIndex > 0) {
    archiveSheet.getRange(CONFIG.headerRow, statusColIndex)
               .setBackground('#212529')
               .setValue(CONFIG.messages.archive.statusColumnHeader);
  }

  // Player columns in archive
  const playerStartCol = CONFIG.firstPlayerColumn;
  const playerEndCol = statusColIndex > 0 ? statusColIndex - 1 : lastCol;
  for (let col = playerStartCol; col <= playerEndCol; col++) {
    const currentHeader = archiveSheet.getRange(CONFIG.headerRow, col).getValue();
    if (currentHeader && currentHeader.toString().trim() !== '') {
      archiveSheet.getRange(CONFIG.headerRow, col)
                 .setValue(currentHeader.toString().replace(/^👤\s*/, ''));
    }
  }

  // --- Data Row Formatting (muted colors for archive) ---
  for (let row = CONFIG.firstDataRow; row <= lastRow; row++) {
    // Alternating row colors for readability
    const rowColor = row % 2 === 0 ? '#f8f9fa' : '#ffffff';
    archiveSheet.getRange(row, 1, 1, lastCol).setBackground(rowColor);

    // Date column formatting
    const dateCell = archiveSheet.getRange(row, CONFIG.dateColumn);
    dateCell.setBackground('#e9ecef')
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setNumberFormat('yyyy.mm.dd');

    // Player response columns
    for (let col = playerStartCol; col <= playerEndCol; col++) {
      const cell = archiveSheet.getRange(row, col);
      cell.setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setFontSize(10)
          .setFontColor('#6c757d'); // Muted text for archived data
    }

    // Status column formatting
    if (statusColIndex > 0) {
      const statusCell = archiveSheet.getRange(row, statusColIndex);
      statusCell.setBackground('#e9ecef')
               .setFontSize(9)
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle')
               .setFontColor('#495057');
    }
  }

  // --- Archive-specific conditional formatting (muted) ---
  addArchiveConditionalFormatting(archiveSheet, playerStartCol, playerEndCol, statusColIndex);

  // --- Column Widths ---
  archiveSheet.setColumnWidth(CONFIG.dateColumn, 150);
  for (let col = playerStartCol; col <= playerEndCol; col++) {
    archiveSheet.setColumnWidth(col, 70);
  }
  if (statusColIndex > 0) {
    archiveSheet.setColumnWidth(statusColIndex, 180);
  }

  // Freeze header row and date column
  archiveSheet.setFrozenRows(CONFIG.headerRow);
  archiveSheet.setFrozenColumns(CONFIG.dateColumn);

  Logger.log('Archive sheet formatting applied successfully.');
}

/**
 * Adds muted conditional formatting to the archive sheet
 */
function addArchiveConditionalFormatting(sheet, playerStartCol, playerEndCol, statusColIndex) {
  const rules = [];
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.firstDataRow) return;

  // Player response conditional formatting (muted colors)
  const playerRange = sheet.getRange(CONFIG.firstDataRow, playerStartCol,
                                   lastRow - CONFIG.firstDataRow + 1,
                                   playerEndCol - playerStartCol + 1);

  // Muted yes responses
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('y')
    .setBackground('#e8f5e8')
    .setFontColor('#4a6741')
    .setRanges([playerRange])
    .build());

  // Muted no responses
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('n')
    .setBackground('#f5e8e8')
    .setFontColor('#674141')
    .setRanges([playerRange])
    .build());

  // Muted maybe responses
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('?')
    .setBackground('#f5f1e8')
    .setFontColor('#675d41')
    .setRanges([playerRange])
    .build());

  sheet.setConditionalFormatRules(rules);
}
