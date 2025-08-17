/**
 * Legacy utility functions - now delegates to UtilityService
 * Kept for backward compatibility
 */

/**
 * Checks if a string is a valid time format
 * @deprecated Use UtilityService.validateTimeFormat() instead
 */
function isTime(str) {
  const validation = UtilityService.validateTimeFormat(str);
  return {
    isValid: validation.isValid,
    rangeRegex: /^(\d{1,2}(?::\d{2})?)-(\d{1,2}(?::\d{2})?)$/,
    singleTimeRegex: /^(\d{1,2}(?::\d{2})?)$/,
    rangeMatch: validation.rangeMatch,
    singleMatch: validation.singleMatch
  };
}

/**
 * Parses a string to extract a start and end time
 * @deprecated Use UtilityService.parseTimeRange() instead
 */
function parseTimeRange(timeStr, baseDate) {
  const result = UtilityService.parseTimeRange(timeStr, baseDate);
  return {
    start: result.start,
    end: result.end
  };
}

/**
 * Creates future date rows
 * @deprecated Use UtilityService.createFutureDateRows() instead
 */
function createFutureDateRows(ss, startDate) {
  UtilityService.createFutureDateRows(ss, startDate);
}

/**
 * Get week number for a date
 * @deprecated Use UtilityService.getWeekNumber() instead
 */
function getWeekNumber(date) {
  return UtilityService.getWeekNumber(date);
}

/**
 * Calculate intersection for player combination
 * @deprecated Use UtilityService.calculateTimeIntersection() instead
 */
function calculateIntersectionForCombination(responses, baseDate) {
  const validResponses = responses.filter(r => r && r.toString().trim() !== '');
  const result = UtilityService.calculateTimeIntersection(validResponses, baseDate);

  return {
    intersectionStart: result.start,
    intersectionEnd: result.end
  };
}

/**
 * Generate combinations of specified size
 * @deprecated Use UtilityService.generateCombinations() instead
 */
function generateCombinations(array, size) {
  return UtilityService.generateCombinations(array, size);
}
