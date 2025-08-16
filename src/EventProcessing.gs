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


  // OPTIMIZATION: Find restricting players FIRST to exclude them from combination search
  const restrictingPlayers = findRestrictingPlayers(validPlayerResponses, baseDate);

  // Create a filtered list excluding restricting players for optimal combination search
  const nonRestrictingPlayers = validPlayerResponses.filter(p => !restrictingPlayers.includes(p.playerName));

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

  // If we have no non-restricting players, return early
  if (nonRestrictingPlayers.length === 0) {
    Logger.log(`All players are restricting - no valid combination possible`);
    return {
      players: [],
      intersectionStart: null,
      intersectionEnd: null,
      duration: 0,
      restrictingPlayers: restrictingPlayers
    };
  }

  // Find the optimal combination that meets the threshold - using ONLY non-RESTRICTING players
  let bestCombination = {
    players: [],
    intersectionStart: null,
    intersectionEnd: null,
    duration: 0,
    restrictingPlayers: restrictingPlayers
  };

  const minCombinationSize = Math.ceil(totalPlayers * CONFIG.playerCombinationThresholdPercentage);
  if (minCombinationSize > nonRestrictingPlayers.length) {
    return bestCombination; // No valid combinations possible
  }
  // Try all possible combinations of NON-RESTRICTING players (starting from largest)
  const maxSearchSize = Math.min(nonRestrictingPlayers.length, totalPlayers); // Don't search beyond total players
  for (let size = maxSearchSize; size >= minCombinationSize; size--) {
    // Skip if we don't have enough non-restricting players for this size
    if (size > nonRestrictingPlayers.length) {
      continue;
    }

    const combinations = generateCombinations(nonRestrictingPlayers, size);

    for (let combIndex = 0; combIndex < combinations.length; combIndex++) {
      const combination = combinations[combIndex];
      const combinationPlayerNames = combination.map(p => p.playerName);
      const combinationResponses = combination.map(p => p.response);


      const { intersectionStart, intersectionEnd } = calculateIntersectionForCombination(combinationResponses, baseDate);

      if (intersectionStart !== undefined && intersectionEnd !== undefined) {
        const duration = intersectionEnd ? (intersectionEnd.getTime() - intersectionStart.getTime()) / 3600000 : 24; // Convert to hours

        // Check if this combination meets the minimum event duration
        if (duration >= CONFIG.minEventDurationHours) {

          bestCombination = {
            players: combinationPlayerNames,
            intersectionStart,
            intersectionEnd,
            duration,
            restrictingPlayers: restrictingPlayers
          };

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
    if (response === CONFIG.responses.no || response === CONFIG.responses.empty || response === CONFIG.responses.maybe) {
      continue;
    }

    // Skip players who said 'y' (they're available all day)
    if (response === CONFIG.responses.yes) {
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
    if (responseStr === CONFIG.responses.empty || responseStr === CONFIG.responses.maybe) {
      allYResponses = false;
      continue;
    }

    // Skip 'n' responses - they make combination invalid
    if (responseStr === CONFIG.responses.no) {
      return { intersectionStart: undefined, intersectionEnd: undefined };
    }

    // Handle 'y' responses
    if (responseStr === CONFIG.responses.yes) {
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
        if (responseStr === CONFIG.responses.empty || responseStr === CONFIG.responses.maybe) {
            allYResponses = false;
            continue;
        }

        // Skip 'n' responses - they don't affect intersection but make event unschedulable
        if (responseStr === CONFIG.responses.no) {
            allYResponses = false;
            continue;
        }

        // Handle 'y' responses
        if (responseStr === CONFIG.responses.yes) {
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
      if (responseStr === CONFIG.responses.no) {
        nFound = true;
      } else if (responseStr === CONFIG.responses.yes) {
        yCount++;
      } else if (responseStr === CONFIG.responses.maybe) {
        questionMarkCount++;
      } else if (responseStr === CONFIG.responses.empty) {
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
    statusCell.setValue(CONFIG.messages.status.cancelled);
  } else if (yCount + timeResponsesCount === numPlayers && actualPlayerResponses === numPlayers) {
    statusCell.setValue(CONFIG.messages.status.readyForScheduling);
  } else if ((blankCount > 0 || questionMarkCount > 0) && yCount + timeResponsesCount < numPlayers && actualPlayerResponses > 0) {
    statusCell.setValue(CONFIG.messages.status.awaitingResponses);
  } else {
    statusCell.setValue(""); // Clear status if state is indeterminate
  }
}
