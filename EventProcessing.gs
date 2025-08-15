/**
 * Event processing and scheduling logic
 */

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
 * Analyzes player responses and updates status for a specific row
 */
function analyzeRowResponses(sheet, editedRow, playerInfo, numPlayers, statusColIndex, numPlayerColumns, allPlayerNames) {
  const dateCell = sheet.getRange(editedRow, CONFIG.dateColumn).getValue();
  const date = new Date(dateCell);

  const allResponses = sheet.getRange(editedRow, CONFIG.firstPlayerColumn, 1, numPlayerColumns).getValues().flat();

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
