/**
 * Legacy Discord notification functions - now delegates to DiscordService
 * Kept for backward compatibility
 */

/**
 * @deprecated Use DiscordService.sendEventNotification() instead
 */
function sendDiscordEventNotification(date, start, end, eventTitle, eventLink, triggerType = 'monthly') {
  const discordService = new DiscordService();
  return discordService.sendEventNotification(date, start, end, eventTitle, eventLink, triggerType);
}

/**
 * @deprecated Use DiscordService.sendReminder() instead
 */
function sendDiscordReminder(reminderEmails, reminderType = 'mixed', triggerType = 'monthly') {
  const discordService = new DiscordService();
  const discordIds = Array.isArray(reminderEmails) ? reminderEmails : [...reminderEmails];
  return discordService.sendReminder(discordIds, reminderType, triggerType);
}

/**
 * @deprecated Use DiscordService.sendDurationWarning() instead
 */
function sendDiscordDurationRestrictionNotification(restrictingPlayers, eventDate, optimalDuration, playerInfo, triggerType = 'monthly') {
  const discordService = new DiscordService();
  return discordService.sendDurationWarning(restrictingPlayers, eventDate, playerInfo, triggerType);
}

/**
 * @deprecated Use DiscordService.sendSheetSetupNotification() instead
 */
function sendDiscordSheetSetupNotification() {
  const discordService = new DiscordService();
  return discordService.sendSheetSetupNotification();
}
