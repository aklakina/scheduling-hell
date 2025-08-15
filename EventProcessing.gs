/**
 * Event processing and scheduling logic
 */

/**
 * Finds the largest combination of players where the intersected duration
 * meets the minimum event duration threshold.
 * Returns an object with player combination details and restriction analysis.
 */
function findOptimalPlayerCombination(responses, allPlayerNames, playerInfo, baseDate) {
  const totalPlayers = Object.keys(playerInfo).length;
  const validPlayerResponses = [];

  // Filter responses to only include actual players from roster
  responses.forEach((response, index) => {
    const playerName = allPlayerNames[index];
    if (playerName && playerInfo[playerName]) {
      validPlayerResponses.push({
        playerName,
        response: response ? String(response).trim().toLowerCase() : '',
        index
      });
    }
  });

  // First, calculate what the intersection would be with ALL players
  const allPlayerResponses = validPlayerResponses.map(p => p.response);
  const { intersectionStart: allPlayersStart, intersectionEnd: allPlayersEnd } = calculateIntersectionForCombination(allPlayerResponses, baseDate);

  let allPlayersDuration = 0;
  if (allPlayersStart !== undefined && allPlayersEnd !== undefined) {
    allPlayersDuration = allPlayersEnd ? (allPlayersEnd.getTime() - allPlayersStart.getTime()) / 3600000 : 24;
  }

  // If all players together can achieve the minimum duration, no notifications needed
  if (allPlayersDuration >= CONFIG.minEventDurationHours) {
    return {
      players: validPlayerResponses.map(p => p.playerName),
      intersectionStart: allPlayersStart,
      intersectionEnd: allPlayersEnd,
      duration: allPlayersDuration,
      restrictingPlayers: []
    };
  }

  // Find the optimal combination that meets the threshold
  let bestCombination = {
    players: [],
    intersectionStart: null,
    intersectionEnd: null,
    duration: 0,
    restrictingPlayers: []
  };

  // Try all possible combinations of players (starting from largest)
  for (let size = validPlayerResponses.length; size >= Math.ceil(totalPlayers * CONFIG.playerCombinationThresholdPercentage); size--) {
    const combinations = generateCombinations(validPlayerResponses, size);

    for (const combination of combinations) {
      const combinationResponses = combination.map(p => p.response);
      const { intersectionStart, intersectionEnd } = calculateIntersectionForCombination(combinationResponses, baseDate);

      if (intersectionStart !== undefined && intersectionEnd !== undefined) {
        const duration = intersectionEnd ? (intersectionEnd.getTime() - intersectionStart.getTime()) / 3600000 : 24; // Convert to hours

        // Check if this combination meets the minimum event duration
        if (duration >= CONFIG.minEventDurationHours) {
          // Find restricting players - those whose time constraints limit the full group
          const restrictingPlayers = findRestrictingPlayers(validPlayerResponses, baseDate);

          bestCombination = {
            players: combination.map(p => p.playerName),
            intersectionStart,
            intersectionEnd,
            duration,
            restrictingPlayers
          };

          // Found the largest valid combination, return it
          return bestCombination;
        }
      }
    }
  }

  return bestCombination;
}

/**
 * Identifies players whose time constraints are restricting the event duration
 */
function findRestrictingPlayers(validPlayerResponses, baseDate) {
  const restrictingPlayers = [];

  for (const player of validPlayerResponses) {
    const response = player.response;

    // Skip players who said 'n' or have empty/invalid responses
    if (response === 'n' || response === '' || response === '?') {
      continue;
    }

    // Skip players who said 'y' (they're available all day)
    if (response === 'y') {
      continue;
    }

    // Check if this is a time range that's shorter than the minimum duration
    const parsedTime = parseTimeRange(response, baseDate);
    if (parsedTime) {
      const playerDuration = (parsedTime.end.getTime() - parsedTime.start.getTime()) / 3600000; // Convert to hours

      // If this player's availability is less than the minimum event duration, they're restricting
      if (playerDuration < CONFIG.minEventDurationHours) {
        restrictingPlayers.push(player.playerName);
      }
    }
  }

  return restrictingPlayers;
}

/**
 * Calculate intersection for a specific combination of responses
 */
function calculateIntersectionForCombination(responses, baseDate) {
  let intersectionStart = new Date(baseDate);
  intersectionStart.setHours(0, 0, 0, 0);
  let intersectionEnd = new Date(baseDate);
  intersectionEnd.setHours(23, 59, 59, 999);
  let hasTimeResponses = false;
  let allYResponses = true;

  for (const response of responses) {
    const responseStr = String(response).trim().toLowerCase();

    // Skip empty responses and question marks - they don't affect intersection
    if (responseStr === '' || responseStr === '?') {
      allYResponses = false;
      continue;
    }

    // Skip 'n' responses - they make combination invalid
    if (responseStr === 'n') {
      return { intersectionStart: undefined, intersectionEnd: undefined };
    }

    // Handle 'y' responses
    if (responseStr === 'y') {
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
      return { intersectionStart: undefined, intersectionEnd: undefined };
    }
  }

  // If all valid responses are 'Y', it's an all-day event
  if (!hasTimeResponses && allYResponses) {
    return { intersectionStart: null, intersectionEnd: null };
  }

  // If no time responses but not all Y, then use full day
  if (!hasTimeResponses) {
    return { intersectionStart: null, intersectionEnd: null };
  }

  // Check if intersection is valid and meets minimum consideration duration
  if (intersectionStart >= intersectionEnd) {
    return { intersectionStart: undefined, intersectionEnd: undefined };
  }

  const durationInMs = intersectionEnd.getTime() - intersectionStart.getTime();
  if (durationInMs < (CONFIG.minConsiderationDurationHours * 3600000)) {
    return { intersectionStart: undefined, intersectionEnd: undefined };
  }

  return { intersectionStart, intersectionEnd };
}

/**
 * Generate all combinations of a given size from an array
 */
function generateCombinations(arr, size) {
  if (size === 0) return [[]];
  if (size > arr.length) return [];

  const combinations = [];

  function backtrack(start, current) {
    if (current.length === size) {
      combinations.push([...current]);
      return;
    }

    for (let i = start; i < arr.length; i++) {
      current.push(arr[i]);
      backtrack(i + 1, current);
      current.pop();
    }
  }

  backtrack(0, []);
  return combinations;
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
    if (durationInMs < (CONFIG.minConsiderationDurationHours * 3600000)) {
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
