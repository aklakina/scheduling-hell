/**
 * Utility functions for date, time, and general operations
 */

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
