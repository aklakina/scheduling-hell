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

// --- Configuration ---
// Adjust these settings to match your spreadsheet's layout.
const CONFIG = {
  responseSheetName: "Responses",
  rosterSheetName: "Player Roster",
  campaignDetailsSheetName: "Campaign details",
  archiveSheetName: "Archive",
  headerRow: 1,
  firstDataRow: 2,
  dateColumn: 1,
  firstPlayerColumn: 3,
  statusColumnName: "Status",
  minEventDurationHours: 2,
  shortEventWarningHours: 4,
  // Updated: Auto-scheduling configuration for 2 months ahead including today
  monthsToCreateAhead: 2,     // Always maintain 2 months of future dates including today
  weeksToKeepBeforeArchive: 1, // Keep last week's data before archiving
  // Discord webhook configuration
  discordWebhookUrl: PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK'), // Add your Discord webhook URL here (e.g., "https://discord.com/api/webhooks/...")
  discordChannelMention: "@everyone" // Change to specific role mention if needed (e.g., "<@&ROLE_ID>")
};
// --------------------


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

  // Apply header formatting
  const headerRange = sheet.getRange(CONFIG.headerRow, 1, 1, headers.length);
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
            .requireValueInList(['✅ y', '❌ n', '❓ ?', ''], true)
            .setAllowInvalid(true) // Allow time ranges and custom entries
            .setHelpText('Quick select: ✅ y (yes), ❌ n (no), ❓ ? (maybe), or enter time range (e.g., 18-22)')
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
  const allResponses = sheet.getRange(editedRow, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();
  const allPlayerNames = sheet.getRange(CONFIG.headerRow, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

  let yCount = 0;
  let questionMarkCount = 0;
  let timeResponsesCount = 0;
  let nFound = false;
  let blankCount = 0;
  let actualPlayerResponses = 0; // Count of responses from actual players

  allResponses.forEach((response, index) => {
    const playerName = allPlayerNames[index];
    // Only count responses for actual players (those in the roster)
    if (playerName && playerInfo[playerName]) {
      actualPlayerResponses++;

      const responseStr = response ? String(response).trim().toLowerCase() : '';
      if (responseStr === 'n') {
        nFound = true;
      } else if (responseStr === 'y') {
        yCount++;
      } else if (responseStr === '?') {
        questionMarkCount++;
      } else if (responseStr === '') {
        blankCount++;
      } else {
        // Check if it's a valid time range
        if (parseTimeRange(responseStr, date)) {
          timeResponsesCount++;
        } else {
          questionMarkCount++; // Treat invalid time formats as '?'
        }
      }
    }
  });

  const statusCell = sheet.getRange(editedRow, statusColIndex);

  if (nFound) {
    statusCell.setValue("Cancelled (No consensus)");
  } else if (yCount + timeResponsesCount === numPlayers && actualPlayerResponses === numPlayers) {
    statusCell.setValue("Ready for scheduling");
  } else if (blankCount > 0 || questionMarkCount > 0) {
    statusCell.setValue("Awaiting responses");
  } else {
    statusCell.setValue(""); // Clear status if state is indeterminate
  }
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
      // --- DISCORD WEBHOOK NOTIFICATION ---
      try {
        const webhookUrl = CONFIG.discordWebhookUrl;
        const channelMention = CONFIG.discordChannelMention;

        if (webhookUrl) {
          const eventDate = bestEvent.date.toLocaleDateString();
          const eventTime = bestEvent.start ? ` from ${bestEvent.start.toLocaleTimeString()}` : '';
          const eventEndTime = bestEvent.end ? ` to ${bestEvent.end.toLocaleTimeString()}` : '';
          const eventTitle = `${eventTitleFromSheet} - ${eventDate}${eventTime}${eventEndTime}`;
          const eventLinkMessage = eventLink ? `\nEvent Link: ${eventLink}` : '';

          const payload = JSON.stringify({
            content: `${channelMention} 🎉 **Event Scheduled**: ${eventTitle}!${eventLinkMessage}`,
            username: "Scheduler Bot",
            avatar_url: "https://example.com/bot-avatar.png"
          });

          const options = {
            method: "post",
            contentType: "application/json",
            payload: payload
          };

          UrlFetchApp.fetch(webhookUrl, options);
          Logger.log(`Discord event notification sent: ${eventTitle}`);
        }
      } catch (error) {
        Logger.log(`Error sending Discord event notification: ${error.toString()}`);
      }

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
    try {
      // --- DISCORD WEBHOOK REMINDER ---
      const webhookUrl = CONFIG.discordWebhookUrl;
      const channelMention = CONFIG.discordChannelMention;

      if (webhookUrl) {
        // Create Discord mentions for players who need reminders
        const discordMentions = [];
        [...globalReminderEmails].forEach(discordId => {
          if (discordId && discordId.trim() !== '') {
            discordMentions.push(`<@${discordId}>`);
          }
        });

        const mentionText = discordMentions.length > 0 ? discordMentions.join(' ') : channelMention;

        const payload = JSON.stringify({
          content: `${mentionText} 📅 **Reminder**: Please update your availability for upcoming events in the Google Sheet. We are trying to finalize the schedule for the next few weeks. Thanks!`,
          username: "Scheduler Bot",
          avatar_url: "https://example.com/bot-avatar.png" // Optional: Set a custom avatar for the bot
        });

        const options = {
          method: "post",
          contentType: "application/json",
          payload: payload
        };

        UrlFetchApp.fetch(webhookUrl, options);
        Logger.log(`Discord reminder sent to ${globalReminderEmails.size} participants: ${[...globalReminderEmails].join(', ')}`);
      } else {
        Logger.log(`Discord webhook URL is not set. Skipping reminder for ${globalReminderEmails.size} participants`);
      }

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
    } catch (error) {
      Logger.log(`Error sending Discord reminder: ${error.toString()}`);
      // Don't update status if Discord message failed to send
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

/**
 * Helper function to calculate the intersection of available times.
 * Returns an object with {intersectionStart, intersectionEnd}.
 * For all-day events (all 'Y'), returns {null, null}.
 * If no valid intersection, returns {undefined, undefined}.
 */
function calculateIntersection(responses, baseDate) {
    let intersectionStart = new Date(baseDate);
    intersectionStart.setHours(0, 0, 0, 0);
    let intersectionEnd = new Date(baseDate);
    intersectionEnd.setHours(23, 59, 59, 999);
    let hasTimeResponses = false;
    let allYResponses = true;

    for (const response of responses) {
        const responseStr = response ? String(response).trim().toLowerCase() : '';

        // Skip empty responses and question marks - they don't affect intersection
        if (responseStr === '' || responseStr === '?') {
            allYResponses = false;
            continue;
        }

        // Skip 'n' responses - they don't affect intersection but make event unschedulable
        if (responseStr === 'n') {
            allYResponses = false;
            continue;
        }

        // Handle 'y' responses
        if (responseStr === 'y') {
            allYResponses = allYResponses && true;
            continue;
        }

        // Try to parse as time range
        const parsedTime = parseTimeRange(responseStr, baseDate);
        if (parsedTime) {
            hasTimeResponses = true;
            allYResponses = false;
            if (parsedTime.start > intersectionStart) intersectionStart = parsedTime.start;
            if (parsedTime.end < intersectionEnd) intersectionEnd = parsedTime.end;
        } else {
            // Invalid format, treat as unavailable
            allYResponses = false;
        }
    }

    // If all valid responses are 'Y', it's an all-day event
    if (!hasTimeResponses && allYResponses) {
        return { intersectionStart: null, intersectionEnd: null };
    }

    // If no time responses but not all Y, then no valid intersection
    if (!hasTimeResponses) {
        return { intersectionStart: undefined, intersectionEnd: undefined };
    }

    // Check if intersection is valid and meets minimum duration
    if (intersectionStart >= intersectionEnd) {
        return { intersectionStart: undefined, intersectionEnd: undefined };
    }

    const durationInMs = intersectionEnd.getTime() - intersectionStart.getTime();
    if (durationInMs < (CONFIG.minEventDurationHours * 3600000)) {
        return { intersectionStart: undefined, intersectionEnd: undefined };
    }

    return { intersectionStart, intersectionEnd };
}


/**
 * Helper function to send Discord notifications for scheduled events.
 * Replaces the calendar event creation functionality.
 */
function sendDiscordEventNotification(date, start, end, eventTitle, eventLink) {
  try {
    const webhookUrl = CONFIG.discordWebhookUrl;
    const channelMention = CONFIG.discordChannelMention;

    if (!webhookUrl) {
      Logger.log(`Discord webhook URL is not set. Skipping event notification.`);
      return;
    }

    const eventDate = date.toLocaleDateString();
    const eventTime = start ? ` from ${start.toLocaleTimeString()}` : '';
    const eventEndTime = end ? ` to ${end.toLocaleTimeString()}` : '';
    const fullEventTitle = `${eventTitle} - ${eventDate}${eventTime}${eventEndTime}`;
    const eventLinkMessage = eventLink ? `\nEvent Link: ${eventLink}` : '';

    const payload = JSON.stringify({
      content: `${channelMention} 🎉 **Event Scheduled**: ${fullEventTitle}!${eventLinkMessage}`,
      username: "Scheduler Bot",
      avatar_url: "https://example.com/bot-avatar.png"
    });

    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload
    };

    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`Discord event notification sent: ${fullEventTitle}`);
  } catch (error) {
    Logger.log(`Error sending Discord event notification: ${error.toString()}`);
  }
}


/**
 * Parses a string to extract a start and end time.
 */
function parseTimeRange(timeStr, baseDate) {
  timeStr = String(timeStr).replace(/\s/g, '');

  // Updated regex patterns to handle seconds and be more flexible
  // Supports: HH, HH:MM for both single times and ranges
  const rangeRegex = /^(\d{1,2}(?::\d{2})?)-(\d{1,2}(?::\d{2})?)$/;
  const singleTimeRegex = /^(\d{1,2}(?::\d{2})?)$/;
  let match;

  const createDate = (timePart) => {
    const newDate = new Date(baseDate);
    const parts = timePart.split(':');
    const hours = parseInt(parts[0], 10);
    const minutes = parts.length > 1 ? parseInt(parts[1], 10) : 0;

    // Validate time components
    if (isNaN(hours) || isNaN(minutes) ||
        hours > 23 || minutes > 59 ||
        hours < 0 || minutes < 0) {
      return null;
    }

    newDate.setHours(hours, minutes, 0, 0);
    return newDate;
  };

  if ((match = timeStr.match(rangeRegex))) {
    const start = createDate(match[1]);
    const end = createDate(match[2]);
    if (start && end && start < end) return { start, end };
  } else if ((match = timeStr.match(singleTimeRegex))) {
    const start = createDate(match[1]);
    if (start) {
      const end = new Date(start);
      end.setHours(start.getHours() + CONFIG.shortEventWarningHours);
      return { start, end };
    }
  }
  return null;
}

/**
 * Helper function to get the week number for a given date.
 */
function getWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
    var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
    return d.getUTCFullYear() + '-' + weekNo;
}

/**
 * Gets the last run date from script properties.
 * Returns a Date object or null if not set.
 */
function getLastRunDate() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastRun = scriptProperties.getProperty('LAST_RUN_DATE');
  return lastRun ? new Date(lastRun) : null;
}

/**
 * Sets the last run date in script properties.
 */
function setLastRunDate(date) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('LAST_RUN_DATE', date.toISOString());
}

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
}

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
            .requireValueInList(['✅ y', '❌ n', '❓ ?', ''], true)
            .setAllowInvalid(true) // Allow time ranges and custom entries
            .setHelpText('Quick select: ✅ y (yes), ❌ n (no), ❓ ? (maybe), or enter time range (e.g., 18-22)')
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
                                   playerEndCol - playerStartCol + 1);

  // Yes responses (green) - matches cells containing 'y'
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('y')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([playerRange])
    .build());

  // No responses (red) - matches cells containing 'n'
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('n')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([playerRange])
    .build());

  // Maybe responses (yellow) - matches cells containing '?'
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('?')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([playerRange])
    .build());

  // Time range responses (blue) - matches cells that start with a number (time ranges)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(NOT(ISBLANK(INDIRECT(ADDRESS(ROW(),COLUMN())))), REGEXMATCH(INDIRECT(ADDRESS(ROW(),COLUMN())), "^[0-9]"))')
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
             .setValue('Date (Archived)');

  const updatedHeaders = archiveSheet.getRange(CONFIG.headerRow, 1, 1, lastCol).getValues()[0];
  const statusColIndex = updatedHeaders.indexOf(CONFIG.statusColumnName) > -1
    ? updatedHeaders.indexOf(CONFIG.statusColumnName) + 1
    : updatedHeaders.indexOf('Final Status') + 1;

  if (statusColIndex > 0) {
    archiveSheet.getRange(CONFIG.headerRow, statusColIndex)
               .setBackground('#212529')
               .setValue('Final Status');
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
