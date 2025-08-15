/**
 * @OnlyCurrentDoc
 *
 * This script provides an intelligent, week-aware scheduling solution for Google Sheets.
 * It is designed to be split into two primary functions:
 * 1. onEditFeedback(e): Provides immediate UI feedback and updates the row's status column.
 * This should be triggered by an 'On edit' event.
 * 2. checkAndScheduleEvents(): Analyzes a future 14-day window, finds the best
 * opportunity per week, and either schedules an event, sends reminders, or marks events as failed.
 * This should be run on a time-based trigger (e.g., every two weeks).
 */

/**
 * Creates a custom menu in the spreadsheet UI to allow easy setup.
 * Runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scheduler')
    .addItem('Setup Sheet', 'setupSheet')
    .addSeparator()
    .addItem('Run Now', 'checkAndScheduleEvents')
    .addSeparator()
    .addItem('🎨 Format Response Sheet', 'formatResponseSheet')
    .addItem('🎨 Format Archive Sheet', 'formatArchiveSheet')
    .addToUi();
}

/**
 * Sets up the response sheet by creating the proper structure with roster-based columns.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.responseSheetName);

  // Get roster data first to build the proper structure
  const rosterSheet = ss.getSheetByName(CONFIG.rosterSheetName);
  if (!rosterSheet || rosterSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(`Error: The "${CONFIG.rosterSheetName}" sheet was not found or has no player data. Please create the roster sheet first with player names in column A.`);
    return;
  }

  const rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 1).getValues();
  const playerNames = rosterData.map(row => row[0]).filter(name => name && name.toString().trim() !== '');

  if (playerNames.length === 0) {
    SpreadsheetApp.getUi().alert(`Error: No player names found in the "${CONFIG.rosterSheetName}" sheet. Please add player names in column A starting from row 2.`);
    return;
  }

  // Create or recreate the response sheet with proper structure
  if (sheet) {
    const response = SpreadsheetApp.getUi().alert(
      'Setup Response Sheet',
      `The "${CONFIG.responseSheetName}" sheet already exists. Do you want to recreate it with the current roster structure? This will delete all existing data.`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response === SpreadsheetApp.getUi().Button.YES) {
      ss.deleteSheet(sheet);
      sheet = ss.insertSheet(CONFIG.responseSheetName);
    } else {
      SpreadsheetApp.getUi().alert('Setup cancelled. No changes were made.');
      return;
    }
  } else {
    sheet = ss.insertSheet(CONFIG.responseSheetName);
  }

  // Build the header structure
  const headers = [
    'Date',              // Column 1 (A): Date column
    'Day',               // Column 2 (B): Localized day name
    ...playerNames.map(name => name), // Columns 3+: Player columns (no emoji)
    'Today',             // Today indicator column
    'Status'             // Status column
  ];

  // Set headers
  sheet.getRange(CONFIG.headerRow, 1, 1, headers.length).setValues([headers]);

  Logger.log('Response sheet setup completed successfully.');

  // Apply formatting after setup
  formatResponseSheet();
}

/**
 * TRIGGER 1: ON EDIT
 * This function runs on every edit to provide immediate UI feedback and
 * update the status column for the edited row.
 *
 * @param {Object} e The event object passed by the OnEdit trigger.
 */
function onEditFeedback(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const editedRow = range.getRow();
  const editedCol = range.getColumn();

  // --- Initial checks to exit early ---
  if (sheet.getName() !== CONFIG.responseSheetName || editedRow < CONFIG.firstDataRow) {
    return;
  }

  const headers = sheet.getRange(CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusColIndex = headers.indexOf(CONFIG.statusColumnName) + 1;

  // If status column doesn't exist, exit silently (setupSheet should be run first)
  if (statusColIndex === 0) {
    return;
  }

  // Find Today column index to exclude it from processing
  const todayColIndex = headers.findIndex(h => h.toString().includes('Today')) + 1;

  // Only process edits in player columns (between firstPlayerColumn and status column)
  // Exclude the Today column from processing
  if (editedCol < CONFIG.firstPlayerColumn || editedCol >= statusColIndex || editedCol === todayColIndex) {
    return;
  }

  // --- Get Player Info ---
  const rosterSheet = ss.getSheetByName(CONFIG.rosterSheetName);
  if (!rosterSheet) return; // Silently exit if roster not found

  const rosterLastRow = rosterSheet.getLastRow();
  if (rosterLastRow < 2) return; // No player data

  const rosterData = rosterSheet.getRange(2, 1, rosterLastRow - 1, 2).getValues();
  const playerInfo = {};
  rosterData.forEach(row => {
    if (row[0]) playerInfo[row[0]] = { discordHandle: row[1] };
  });
  const numPlayers = Object.keys(playerInfo).length;
  if (numPlayers === 0) return;

  // --- Validate the edited cell and provide UI feedback ---
  const ui = SpreadsheetApp.getUi();
  const value = e.value ? String(e.value).trim().toLowerCase() : '';
  const dateCell = sheet.getRange(editedRow, CONFIG.dateColumn).getValue();
  const date = new Date(dateCell);

  if (!['y', 'n', '?', ''].includes(value)) {
    if (isNaN(date.getTime())) {
      ui.alert("Error: Could not find a valid date in column A for this row.");
      return;
    }
    date.setHours(12, 0, 0, 0);
    const parsedTime = parseTimeRange(value, date);
    if (!parsedTime) {
      ui.alert(`Invalid Time Format`, `Your entry "${e.value}" is not a valid time or time range. Please use formats like "18:00", "18-22", or "18:30-22:00".`, ui.ButtonSet.OK);
      return; // Exit early if invalid format
    }
  }

  // --- Analyze the entire row to update the status ---
  const numPlayerColumns = statusColIndex - CONFIG.firstPlayerColumn;
  const allPlayerNames = sheet.getRange(CONFIG.headerRow, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

  analyzeRowResponses(sheet, editedRow, playerInfo, numPlayers, statusColIndex, numPlayerColumns, allPlayerNames);
}

/**
 * TRIGGER 2: TIME-DRIVEN (monthly)
 * This is the main processing function. It analyzes a future window based on
 * when it was last run, finds the best event per week, and schedules or sends reminders.
 *
 * For monthly triggers, this function dynamically adjusts its processing window
 * to ensure complete coverage without gaps or excessive overlap.
 */
function checkAndScheduleEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Get all required sheets and config ---
  const responseSheet = ss.getSheetByName(CONFIG.responseSheetName);
  if (!responseSheet) { Logger.log(`Error: Sheet '${CONFIG.responseSheetName}' not found.`); return; }
  const rosterSheet = ss.getSheetByName(CONFIG.rosterSheetName);
  if (!rosterSheet) { Logger.log(`Error: Sheet '${CONFIG.rosterSheetName}' not found.`); return; }
  const campaignSheet = ss.getSheetByName(CONFIG.campaignDetailsSheetName);
  if (!campaignSheet) { Logger.log(`Error: Sheet '${CONFIG.campaignDetailsSheetName}' not found.`); return; }

  const rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 2).getValues();
  const playerInfo = {};
  rosterData.forEach(row => {
    if (row[0]) playerInfo[row[0]] = { discordHandle: row[1] };
  });
  const numPlayers = Object.keys(playerInfo).length;
  if (numPlayers === 0) { Logger.log("No players found in roster."); return; }

  const campaignData = campaignSheet.getRange(2, 1, 1, 2).getValues()[0];
  const eventTitleFromSheet = campaignData[0];
  const eventLink = campaignData[1];

  const headers = responseSheet.getRange(CONFIG.headerRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  const statusColumnIndex = headers.indexOf(CONFIG.statusColumnName) + 1;
  if (statusColumnIndex === 0) { Logger.log(`Error: Status column not found.`); return; }
  const numPlayerColumns = statusColumnIndex - CONFIG.firstPlayerColumn;
  if (numPlayerColumns <= 0) { Logger.log(`Error: No player columns found.`); return; }
  const allPlayerNames = responseSheet.getRange(CONFIG.headerRow, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

  // --- Dynamic processing window calculation for monthly triggers ---
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Get last run date from script properties, default to 14 days ago if never run
  const lastRunDate = getLastRunDate() || new Date(today.getTime() - (14 * 24 * 60 * 60 * 1000));

  // Calculate processing window: start 3 days from now, extend based on time since last run
  const processingStartDate = new Date(today);
  processingStartDate.setDate(today.getDate() + 3);

  // For monthly triggers, we need a longer window to ensure we don't miss anything
  // If it's been more than 14 days since last run, extend the window
  const daysSinceLastRun = Math.floor((today - lastRunDate) / (24 * 60 * 60 * 1000));
  const windowDays = Math.max(14, Math.min(35, daysSinceLastRun + 14)); // 14-35 day window

  const processingEndDate = new Date(processingStartDate);
  processingEndDate.setDate(processingStartDate.getDate() + windowDays);

  Logger.log(`Processing window: ${processingStartDate.toLocaleDateString()} to ${processingEndDate.toLocaleDateString()} (${windowDays} days)`);
  Logger.log(`Days since last run: ${daysSinceLastRun}`);

  // --- Filter data for the calculated window ---
  const allData = responseSheet.getRange(CONFIG.firstDataRow, 1, responseSheet.getLastRow() - CONFIG.firstDataRow + 1, statusColumnIndex).getValues();
  const upcomingEventsData = allData.map((row, index) => ({ rowData: row, rowIndex: CONFIG.firstDataRow + index }))
    .filter(item => {
      const dateValue = item.rowData[CONFIG.dateColumn - 1];
      if (!dateValue) return false;
      const eventDate = new Date(dateValue);
      const status = item.rowData[statusColumnIndex - 1] || '';
      // Only process events within our defined window that have not been successfully processed.
      return eventDate >= processingStartDate && eventDate < processingEndDate && !status.startsWith('Event created') && !status.startsWith('Superseded');
    });

  if (upcomingEventsData.length === 0) {
    Logger.log(`No upcoming events to process between ${processingStartDate.toLocaleDateString()} and ${processingEndDate.toLocaleDateString()}.`);
    // Update last run date even if no events processed
    setLastRunDate(today);
    return;
  }

  // --- Group events by week ---
  const eventsByWeek = {};
  upcomingEventsData.forEach(event => {
    const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
    const weekNumber = getWeekNumber(eventDate);
    if (!eventsByWeek[weekNumber]) {
      eventsByWeek[weekNumber] = [];
    }
    eventsByWeek[weekNumber].push(event);
  });

  // --- Process each week ---
  const globalReminderEmails = new Set(); // Collect all reminder emails globally

  for (const week in eventsByWeek) {
    let bestEvent = null;
    let maxDuration = 0;
    const failedReadyEvents = []; // Keep track of events that were ready but unschedulable

    // Find the best schedulable event for the week
    eventsByWeek[week].forEach(event => {
      const status = event.rowData[statusColumnIndex - 1];
      if (status === "Ready for scheduling") {
        const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
        const allResponses = responseSheet.getRange(event.rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();
        const { intersectionStart, intersectionEnd } = calculateIntersection(allResponses, eventDate);
        if (intersectionStart == null && intersectionEnd == null) {
          // This is an all-day event (all 'Y')
          bestEvent = {
            date: eventDate,
            start: null,
            end: null,
            rowIndex: event.rowIndex
          }
          maxDuration = 24 * 60 * 60 * 1000; // 24 hours in milliseconds
        } else if (intersectionStart !== undefined && intersectionEnd !== undefined) {
          // This is a valid, schedulable event (timed or all-day)
          const duration = (intersectionEnd ? intersectionEnd.getTime() : new Date(eventDate).setHours(24)) - (intersectionStart ? intersectionStart.getTime() : new Date(eventDate).setHours(0));
          if (duration > maxDuration) {
            maxDuration = duration;
            bestEvent = {
              date: eventDate,
              start: intersectionStart,
              end: intersectionEnd,
              rowIndex: event.rowIndex
            };
          }
        } else {
          // This event was "Ready" but failed validation (e.g., too short)
          failedReadyEvents.push(event.rowIndex);
        }
      }
    });

    // If a best event was found, schedule it and update status for the whole week
    if (bestEvent) {
      // Send Discord notification for the scheduled event
      sendDiscordEventNotification(bestEvent.date, bestEvent.start, bestEvent.end, eventTitleFromSheet, eventLink);

      responseSheet.getRange(bestEvent.rowIndex, statusColumnIndex).setValue(`Event scheduled on ${today.toLocaleDateString()}`);
      Logger.log(`Scheduled best event for week ${week} on ${bestEvent.date.toLocaleDateString()}.`);

      // Mark other days in the week as 'Superseded'
      eventsByWeek[week].forEach(event => {
          if(event.rowIndex !== bestEvent.rowIndex) {
              const currentStatus = responseSheet.getRange(event.rowIndex, statusColumnIndex).getValue() || '';
              if (!currentStatus.startsWith('Event created') && !currentStatus.startsWith('Cancelled')) {
                responseSheet.getRange(event.rowIndex, statusColumnIndex).setValue('Superseded by other event');
              }
          }
      });

    } else {
      // No event could be scheduled. Now check why.
      if (failedReadyEvents.length > 0) {
        failedReadyEvents.forEach(rowIndex => {
          responseSheet.getRange(rowIndex, statusColumnIndex).setValue(`Failed: Duration < ${CONFIG.minEventDurationHours}h`);
        });
        Logger.log(`Marked ${failedReadyEvents.length} events as failed due to short duration for week ${week}.`);
      }

      // Collect reminder emails for "Awaiting" events (but don't send yet)
      eventsByWeek[week].forEach(event => {
        const status = event.rowData[statusColumnIndex - 1];
        if (status === "Awaiting responses") {
          const allResponses = responseSheet.getRange(event.rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();
          allResponses.forEach((response, i) => {
            const responseStr = response ? String(response).trim().toLowerCase() : '';
            if (responseStr === '?' || responseStr === '') {
              const playerName = allPlayerNames[i];
              if (playerInfo[playerName] && playerInfo[playerName].discordHandle) {
                globalReminderEmails.add(playerInfo[playerName].discordHandle);
              }
            }
          });
        }
      });
    }
  }

  // Send consolidated reminders once per participant after processing all weeks
  if (globalReminderEmails.size > 0) {
    const reminderSent = sendDiscordReminder(globalReminderEmails);

    if (reminderSent) {
      // Update status for all reminded rows across all weeks
      for (const week in eventsByWeek) {
        eventsByWeek[week].forEach(event => {
          const status = event.rowData[statusColumnIndex - 1];
          if (status === 'Awaiting responses') {
            // Check if any player in this row needed a reminder
            const allResponses = responseSheet.getRange(event.rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();
            let hasReminderRecipient = false;
            allResponses.forEach((response, i) => {
              const responseStr = response ? String(response).trim().toLowerCase() : '';
              if (responseStr === '?' || responseStr === '') {
                const playerName = allPlayerNames[i];
                if (playerInfo[playerName] && playerInfo[playerName].discordHandle) {
                  hasReminderRecipient = true;
                }
              }
            });

            if (hasReminderRecipient) {
              responseSheet.getRange(event.rowIndex, statusColumnIndex).setValue(`Reminder sent on ${today.toLocaleDateString()}`);
            }
          }
        });
      }
    }
  }

  // Update last run date after successful processing
  setLastRunDate(today);
  Logger.log(`Updated last run date to ${today.toLocaleDateString()}`);

  // --- NEW FEATURES: Archive old data and create new date rows ---
  try {
    archiveOldResponses(ss, processingStartDate);
    createFutureDateRows(ss, today);
  } catch (error) {
    Logger.log(`Error in archive/auto-create operations: ${error.toString()}`);
    // Don't fail the main function if these operations fail
  }
}
