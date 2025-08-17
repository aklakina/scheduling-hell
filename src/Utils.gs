/**
 * Utility functions for date, time, and general operations
 */

/**
 * Checks if a string is a valid time format and returns regex patterns and matches.
 * @param {string} str - The string to check
 * @returns {Object} Object containing regex patterns and match results
 */
function isTime(str) {
  const rangeRegex = /^(\d{1,2}(?::\d{2})?)-(\d{1,2}(?::\d{2})?)$/;
  const singleTimeRegex = /^(\d{1,2}(?::\d{2})?)$/;

  str = String(str).replace(/\s/g, '');

  const rangeMatch = str.match(rangeRegex);
  const singleMatch = str.match(singleTimeRegex);

  return {
    isValid: Boolean(rangeMatch || singleMatch),
    rangeRegex,
    singleTimeRegex,
    rangeMatch,
    singleMatch
  };
}

/**
 * Parses a string to extract a start and end time.
 */
function parseTimeRange(timeStr, baseDate) {
  const timeCheck = isTime(timeStr);

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

  if (timeCheck.rangeMatch) {
    const start = createDate(timeCheck.rangeMatch[1]);
    const end = createDate(timeCheck.rangeMatch[2]);
    if (start && end && start < end) return { start, end };
  } else if (timeCheck.singleMatch) {
    const start = createDate(timeCheck.singleMatch[1]);
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
 * Calculates the processing window for monthly triggers that run on 1st and 16th.
 * Ensures full month coverage with minimal overlap between windows.
 *
 * @param {Date} today - Current date
 * @returns {Object} Object with processingStartDate and processingEndDate
 */
function calculateMonthlyProcessingWindow(today) {
  const currentDay = today.getDate();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();

  // Base processing start: 3 days from today
  const baseStartDate = new Date(today);
  baseStartDate.setDate(today.getDate() + 3);

  let processingStartDate, processingEndDate;

  // Determine which trigger we are (1st or 16th) and calculate appropriate window
  if (currentDay <= 15) {
    // We're running on the 1st (or close to it)
    // Cover from 3 days from now until mid-month of next month
    processingStartDate = new Date(baseStartDate);

    // End date: 15th of next month
    const nextMonth = currentMonth === 11 ? 0 : currentMonth + 1;
    const nextYear = currentMonth === 11 ? currentYear + 1 : currentYear;
    processingEndDate = new Date(nextYear, nextMonth, 15, 23, 59, 59, 999);

  } else {
    // We're running on the 16th (or close to it)
    // Cover from 3 days from now until mid-month of the month after next
    processingStartDate = new Date(baseStartDate);

    // End date: 15th of the month after next
    let targetMonth = currentMonth + 2;
    let targetYear = currentYear;

    if (targetMonth > 11) {
      targetMonth = targetMonth - 12;
      targetYear++;
    }

    processingEndDate = new Date(targetYear, targetMonth, 15, 23, 59, 59, 999);
  }

  return {
    processingStartDate,
    processingEndDate,
    windowDays: Math.ceil((processingEndDate - processingStartDate) / (24 * 60 * 60 * 1000))
  };
}

/**
 * Calculates the processing window for daily triggers.
 * Daily triggers process a smaller window (next 7 days) for immediate scheduling needs.
 *
 * @param {Date} today - Current date
 * @returns {Object} Object with processingStartDate and processingEndDate
 */
function calculateDailyProcessingWindow(today) {
  const processingStartDate = new Date(today);
  processingStartDate.setDate(today.getDate() + 1); // Start from tomorrow

  const processingEndDate = new Date(today);
  processingEndDate.setDate(today.getDate() + 7); // Process next 7 days
  processingEndDate.setHours(23, 59, 59, 999);

  return {
    processingStartDate,
    processingEndDate,
    windowDays: Math.ceil((processingEndDate - processingStartDate) / (24 * 60 * 60 * 1000))
  };
}

/**
 * Calculates the processing window for bi-weekly Monday notifications.
 * Checks the next 2 weeks shifted by 3 days from the notification date.
 *
 * @param {Date} today - Current date (should be a Monday for bi-weekly notifications)
 * @returns {Object} Object with processingStartDate and processingEndDate
 */
function calculateBiWeeklyProcessingWindow(today) {
  const processingStartDate = new Date(today);
  processingStartDate.setDate(today.getDate() + CONFIG.triggers.biWeeklyCheckDaysAhead);

  const processingEndDate = new Date(processingStartDate);
  processingEndDate.setDate(processingStartDate.getDate() + (CONFIG.triggers.biWeeklyCheckWindowWeeks * 7));
  processingEndDate.setHours(23, 59, 59, 999);

  return {
    processingStartDate,
    processingEndDate,
    windowDays: Math.ceil((processingEndDate - processingStartDate) / (24 * 60 * 60 * 1000))
  };
}

/**
 * Determines the trigger type and processing window based on current date and parameters.
 * @param {string} triggerType - 'daily', 'biWeekly', or 'monthly' (defaults to 'monthly' for backward compatibility)
 * @param {Date} today - Current date
 * @returns {Object} Object with triggerType, processingStartDate, processingEndDate, and windowDays
 */
function calculateProcessingWindow(triggerType = 'monthly', today = new Date()) {
  today = new Date(today);
  today.setHours(0, 0, 0, 0);

  switch (triggerType) {
    case 'daily':
      const dailyStart = new Date(today);
      dailyStart.setDate(today.getDate() + 1); // Start tomorrow
      const dailyEnd = new Date(today);
      dailyEnd.setDate(today.getDate() + CONFIG.triggers.daily.windowDays);
      dailyEnd.setHours(23, 59, 59, 999);
      return {
        triggerType: 'daily',
        processingStartDate: dailyStart,
        processingEndDate: dailyEnd,
        windowDays: CONFIG.triggers.daily.windowDays
      };

    case 'biWeekly':
      const biWeeklyStart = new Date(today);
      biWeeklyStart.setDate(today.getDate() + CONFIG.triggers.biWeekly.daysAhead);
      const biWeeklyEnd = new Date(biWeeklyStart);
      biWeeklyEnd.setDate(biWeeklyStart.getDate() + (CONFIG.triggers.biWeekly.windowWeeks * 7));
      biWeeklyEnd.setHours(23, 59, 59, 999);
      return {
        triggerType: 'biWeekly',
        processingStartDate: biWeeklyStart,
        processingEndDate: biWeeklyEnd,
        windowDays: CONFIG.triggers.biWeekly.windowWeeks * 7
      };

    case 'monthly':
    default:
      // Use existing monthly logic
      const monthlyWindow = calculateMonthlyProcessingWindow(today);
      return {
        triggerType: 'monthly',
        ...monthlyWindow
      };
  }
}

/**
 * Checks if today is a bi-weekly Monday (every other Monday).
 * @param {Date} today - Current date
 * @returns {boolean} True if today should run bi-weekly notifications
 */
function shouldRunBiWeeklyNotifications(today = new Date()) {
  // Check if today is Monday (1 = Monday)
  if (today.getDay() !== 1) {
    return false;
  }

  // Get week number and check if it's even (bi-weekly pattern)
  const weekNumber = getWeekNumber(today);
  const weekNum = parseInt(weekNumber.split('-')[1]);
  return weekNum % 2 === 0;
}

/**
 * Gets/sets the last bi-weekly notification date to prevent duplicate runs.
 */
function getLastBiWeeklyNotificationDate() {
  const properties = PropertiesService.getScriptProperties();
  const lastDate = properties.getProperty('LAST_BIWEEKLY_NOTIFICATION');
  return lastDate ? new Date(lastDate) : null;
}

function setLastBiWeeklyNotificationDate(date) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('LAST_BIWEEKLY_NOTIFICATION', date.toISOString());
}
