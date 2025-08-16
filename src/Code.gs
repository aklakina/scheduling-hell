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
    const message = CONFIG.messages.ui.setupSheet.rosterNotFound.replace('{rosterSheetName}', CONFIG.rosterSheetName);
    SpreadsheetApp.getUi().alert(message);
    return;
  }

  const rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 1).getValues();
  const playerNames = rosterData.map(row => row[0]).filter(name => name && name.toString().trim() !== '');

  if (playerNames.length === 0) {
    const message = CONFIG.messages.ui.setupSheet.noPlayersFound.replace('{rosterSheetName}', CONFIG.rosterSheetName);
    SpreadsheetApp.getUi().alert(message);
    return;
  }

  // Create or recreate the response sheet with proper structure
  if (sheet) {
    const confirmMessage = CONFIG.messages.ui.setupSheet.confirmRecreate.replace('{responseSheetName}', CONFIG.responseSheetName);
    const response = SpreadsheetApp.getUi().alert(
      'Setup Response Sheet',
      confirmMessage,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response === SpreadsheetApp.getUi().Button.YES) {
      ss.deleteSheet(sheet);
      sheet = ss.insertSheet(CONFIG.responseSheetName);
    } else {
      SpreadsheetApp.getUi().alert(CONFIG.messages.ui.setupSheet.setupCancelled);
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

  // Create future date rows for the next 2 months
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    createFutureDateRows(ss, today);
    Logger.log('Future date rows created successfully.');
  } catch (error) {
    Logger.log(`Error creating future date rows: ${error.toString()}`);
    SpreadsheetApp.getUi().alert(`Sheet setup completed, but there was an error creating date rows: ${error.toString()}`);
  }

  // Send Discord notification with sheet link
  try {
    const notificationSent = sendDiscordSheetSetupNotification();
    if (notificationSent) {
      Logger.log('Discord setup notification sent successfully.');
    }
  } catch (error) {
    Logger.log(`Error sending Discord setup notification: ${error.toString()}`);
    // Don't fail the setup if Discord notification fails
  }
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
  const todayColIndex = headers.findIndex(h => h.toString().includes(CONFIG.columns.today)) + 1;

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

  // Updated to read 3 columns instead of 2 to include mention preference
  const rosterData = rosterSheet.getRange(2, 1, rosterLastRow - 1, 3).getValues();
  const playerInfo = {};
  rosterData.forEach(row => {
    if (row[0]) {
      playerInfo[row[0]] = {
        discordHandle: row[1] || "",
        allowMention: Boolean(row[2])
      };
    }
  });
  const numPlayers = Object.keys(playerInfo).length;
  if (numPlayers === 0) return;

  // --- Validate the edited cell and provide UI feedback ---
  const ui = SpreadsheetApp.getUi();
  const value = e.value ? String(e.value).trim().toLowerCase() : '';
  const dateCell = sheet.getRange(editedRow, CONFIG.dateColumn).getValue();
  const date = new Date(dateCell);

  if (![CONFIG.responses.yes, CONFIG.responses.no, CONFIG.responses.maybe, CONFIG.responses.empty].includes(value)) {
    if (isNaN(date.getTime())) {
      ui.alert(CONFIG.messages.ui.invalidDate);
      return;
    }
    date.setHours(12, 0, 0, 0);
    const parsedTime = parseTimeRange(value, date);
    if (!parsedTime) {
      const message = CONFIG.messages.ui.invalidTimeFormat.message.replace('{userInput}', e.value);
      ui.alert(CONFIG.messages.ui.invalidTimeFormat.title, message, ui.ButtonSet.OK);
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

  // Updated to read 3 columns instead of 2 to include mention preference checkbox
  const rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 3).getValues();
  const playerInfo = {};
  rosterData.forEach(row => {
    if (row[0]) {
      playerInfo[row[0]] = {
        discordHandle: row[1] || "",
        allowMention: Boolean(row[2]) // New column for mention preference
      };
    }
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

  // Calculate processing window based on monthly trigger schedule (1st and 16th)
  const windowInfo = calculateMonthlyProcessingWindow(today);
  const processingStartDate = windowInfo.processingStartDate;
  const processingEndDate = windowInfo.processingEndDate;
  const windowDays = windowInfo.windowDays;

  Logger.log(`Processing window: ${processingStartDate.toLocaleDateString()} to ${processingEndDate.toLocaleDateString()} (${windowDays} days)`);
  Logger.log(`Current date: ${today.toLocaleDateString()}, trigger type: ${today.getDate() <= 15 ? '1st of month' : '16th of month'}`);

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

  // --- NEW: Create data structure to collect notifications by date ---
  const notificationsByDate = {};

  // --- Process each week ---
  // Replace global reminders with per-date tracking
  const remindersByDate = {};

  for (const week in eventsByWeek) {
    let bestEvent = null;
    let maxDuration = 0;
    const failedReadyEvents = []; // Keep track of events that were ready but unschedulable

    // Find the best schedulable event for the week
    eventsByWeek[week].forEach(event => {
      const status = event.rowData[statusColumnIndex - 1];
      if (status === CONFIG.messages.status.readyForScheduling) {
        const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
        const allResponses = responseSheet.getRange(event.rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

        // Use the new optimal combination logic to ensure we meet the 4-hour minimum
        const optimalCombination = findOptimalPlayerCombination(allResponses, allPlayerNames, playerInfo, eventDate);

        // Check if this event can actually be scheduled with the 4-hour minimum
        if (optimalCombination.duration >= CONFIG.minEventDurationHours &&
            optimalCombination.players.length === numPlayers) {
          // All players can participate for 4+ hours
          const duration = optimalCombination.duration * 3600000; // Convert to milliseconds
          if (duration > maxDuration) {
            maxDuration = duration;
            bestEvent = {
              date: eventDate,
              start: optimalCombination.intersectionStart,
              end: optimalCombination.intersectionEnd,
              rowIndex: event.rowIndex
            };
          }
        } else {
          // This event was "Ready" but doesn't meet the 4-hour minimum with all players
          failedReadyEvents.push(event.rowIndex);
        }
      }
    });

    // --- NEW: Check for events that could meet duration threshold with optimal player combination ---
    eventsByWeek[week].forEach((event, eventIndex) => {
      const status = event.rowData[statusColumnIndex - 1];

      // Check both "Awaiting responses" AND "Ready for scheduling" events
      // Ready events might still need notifications if they fail the 4-hour requirement
      if (status === CONFIG.messages.status.awaitingResponses || status === CONFIG.messages.status.readyForScheduling) {
        const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
        const allResponses = responseSheet.getRange(event.rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

        // Find optimal player combination
        const optimalCombination = findOptimalPlayerCombination(allResponses, allPlayerNames, playerInfo, eventDate);

        // Check if optimal combination meets the 60% threshold and has restricting players
        if (optimalCombination.players.length >= Math.ceil(numPlayers * CONFIG.playerCombinationThresholdPercentage) &&
            optimalCombination.duration >= CONFIG.minEventDurationHours &&
            optimalCombination.restrictingPlayers.length > 0) {

          // Instead of sending immediately, collect for grouped notification
          const dateString = eventDate.toLocaleDateString();
          if (!notificationsByDate[dateString]) {
            notificationsByDate[dateString] = {
              date: eventDate,
              messages: []
            };
          }

          // Add restriction message for this date
          notificationsByDate[dateString].messages.push({
            type: 'restriction',
            players: optimalCombination.restrictingPlayers,
            duration: optimalCombination.duration
          });

          Logger.log(`Queued duration restriction notification for ${dateString} to players: ${optimalCombination.restrictingPlayers.join(', ')}`);
        }
      }
    });

    // If a best event was found, schedule it and update status for the whole week
    if (bestEvent) {
      // Queue Discord notification for scheduled event instead of sending immediately
      const dateString = bestEvent.date.toLocaleDateString();
      if (!notificationsByDate[dateString]) {
        notificationsByDate[dateString] = {
          date: bestEvent.date,
          messages: []
        };
      }

      notificationsByDate[dateString].messages.push({
        type: 'event',
        start: bestEvent.start,
        end: bestEvent.end,
        eventTitle: eventTitleFromSheet,
        eventLink: eventLink
      });

      responseSheet.getRange(bestEvent.rowIndex, statusColumnIndex).setValue(`Event scheduled on ${today.toLocaleDateString()}`);
      Logger.log(`Scheduled best event for week ${week} on ${bestEvent.date.toLocaleDateString()}.`);

      // Mark other days in the week as 'Superseded'
      eventsByWeek[week].forEach(event => {
          if(event.rowIndex !== bestEvent.rowIndex) {
              const currentStatus = responseSheet.getRange(event.rowIndex, statusColumnIndex).getValue() || '';
              if (!currentStatus.startsWith(CONFIG.messages.status.eventCreated) && !currentStatus.startsWith('Cancelled')) {
                responseSheet.getRange(event.rowIndex, statusColumnIndex).setValue(CONFIG.messages.status.superseded);
              }
          }
      });

    } else {
      // No event could be scheduled. Now check why.
      if (failedReadyEvents.length > 0) {
        failedReadyEvents.forEach(rowIndex => {
          const failureMessage = CONFIG.messages.status.failedDuration.replace('{minHours}', CONFIG.minEventDurationHours);
          responseSheet.getRange(rowIndex, statusColumnIndex).setValue(failureMessage);
        });
        Logger.log(`Marked ${failedReadyEvents.length} events as failed due to short duration for week ${week}.`);
      }

      // Collect reminder emails for "Awaiting" events, separating by response type
      eventsByWeek[week].forEach(event => {
        const status = event.rowData[statusColumnIndex - 1];
        if (status === CONFIG.messages.status.awaitingResponses) {
          const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
          const dateString = eventDate.toLocaleDateString();
          const allResponses = responseSheet.getRange(event.rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

          // Initialize tracking for this date if needed
          if (!remindersByDate[dateString]) {
            remindersByDate[dateString] = {
              date: eventDate,
              maybeEmails: new Set(),
              noResponseEmails: new Set()
            };
          }

          // Count Y responses to determine if we should send reminders
          let yCount = 0;
          allResponses.forEach((response, i) => {
            const playerName = allPlayerNames[i];
            if (playerName && playerInfo[playerName]) {
              const responseStr = response ? String(response).trim().toLowerCase() : '';
              if (responseStr === CONFIG.responses.yes || isTime(responseStr).isValid) {
                yCount++;
              }
            }
          });

          // Calculate minimum Y responses needed based on percentage of total players
          const minYResponsesNeeded = Math.ceil(numPlayers * CONFIG.reminderThresholdPercentage);

          // Only collect reminders if there are enough Y or time responses
          if (yCount >= minYResponsesNeeded) {
            // Track maybe and no-response players separately for this date
            allResponses.forEach((response, i) => {
              const playerName = allPlayerNames[i];
              if (playerName && playerInfo[playerName] && playerInfo[playerName].discordHandle) {
                const responseStr = response ? String(response).trim().toLowerCase() : '';

                if (responseStr === CONFIG.responses.maybe) {
                  // Player answered "maybe"
                  remindersByDate[dateString].maybeEmails.add(playerInfo[playerName].discordHandle);
                } else if (responseStr === CONFIG.responses.empty) {
                  // Player didn't respond at all
                  remindersByDate[dateString].noResponseEmails.add(playerInfo[playerName].discordHandle);
                }
              }
            });

            // Mark this date for possible notifications
            if (!notificationsByDate[dateString]) {
              notificationsByDate[dateString] = {
                date: eventDate,
                messages: []
              };
            }
          } else {
            // NEW: Mark dates without enough responses with "Not enough responses" status
            responseSheet.getRange(event.rowIndex, statusColumnIndex).setValue(`Not enough responses (${yCount}/${minYResponsesNeeded} required)`);
          }
        }
      });
    }
  }

  // Add reminder messages to notification groups based on per-date tracking
  for (const dateString in remindersByDate) {
    const dateReminders = remindersByDate[dateString];

    // Only add notifications if we have this date in the notifications list
    if (notificationsByDate[dateString]) {
      // Add "maybe" reminder only if we have maybe responses AND NO empty responses
      if (dateReminders.maybeEmails.size > 0 && dateReminders.noResponseEmails.size === 0) {
        notificationsByDate[dateString].messages.push({
          type: 'reminder',
          reminderType: 'maybe',
          players: [...dateReminders.maybeEmails]
        });
      }

      // Always add "no response" reminder if needed
      if (dateReminders.noResponseEmails.size > 0) {
        notificationsByDate[dateString].messages.push({
          type: 'reminder',
          reminderType: 'noResponse',
          players: [...dateReminders.noResponseEmails]
        });
      }

      // Track if this date has any reminders for status updates later
      notificationsByDate[dateString].hasReminders =
        (dateReminders.maybeEmails.size > 0 && dateReminders.noResponseEmails.size === 0) ||
        dateReminders.noResponseEmails.size > 0;
    }
  }

  // Send consolidated notifications for each date
  if (Object.keys(notificationsByDate).length > 0) {
    const notificationSent = sendGroupedDiscordNotifications(notificationsByDate, playerInfo);

    // Update status for rows that got reminders
    if (notificationSent) {
      for (const week in eventsByWeek) {
        eventsByWeek[week].forEach(event => {
          const status = event.rowData[statusColumnIndex - 1];
          const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
          const dateString = eventDate.toLocaleDateString();

          if (status === CONFIG.messages.status.awaitingResponses &&
              notificationsByDate[dateString] &&
              notificationsByDate[dateString].hasReminders) {

            // Check if any player in this row needed a reminder
            const allResponses = responseSheet.getRange(event.rowIndex, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();
            let hasReminderRecipient = false;
            allResponses.forEach((response, i) => {
              const responseStr = response ? String(response).trim().toLowerCase() : '';
              if (responseStr === CONFIG.responses.maybe || responseStr === CONFIG.responses.empty) {
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
