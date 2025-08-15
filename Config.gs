/**
 * Configuration settings for the scheduling notification system
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
  minConsiderationDurationHours: 2, // Minimum duration to consider an event possible
  minEventDurationHours: 4, // Minimum duration for actual event scheduling
  shortEventWarningHours: 2,
  // Updated: Auto-scheduling configuration for 2 months ahead including today
  monthsToCreateAhead: 2,     // Always maintain 2 months of future dates including today
  weeksToKeepBeforeArchive: 1, // Keep last week's data before archiving
  // Reminder threshold: percentage of players who must have Y responses before sending reminders
  reminderThresholdPercentage: 0.4, // 40% of players must respond with Y before reminders are sent
  // Player combination threshold: percentage of players needed for event duration notification
  playerCombinationThresholdPercentage: 0.6, // 60% of players must be available for minimum event duration
  // Discord webhook configuration
  discordWebhookUrl: PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK'), // Add your Discord webhook URL here (e.g., "https://discord.com/api/webhooks/...")
  discordChannelMention: "@everyone" // Change to specific role mention if needed (e.g., "<@&ROLE_ID>")
};
