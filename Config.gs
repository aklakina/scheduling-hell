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
  minEventDurationHours: 2,
  shortEventWarningHours: 4,
  // Updated: Auto-scheduling configuration for 2 months ahead including today
  monthsToCreateAhead: 2,     // Always maintain 2 months of future dates including today
  weeksToKeepBeforeArchive: 1, // Keep last week's data before archiving
  // Discord webhook configuration
  discordWebhookUrl: PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK'), // Add your Discord webhook URL here (e.g., "https://discord.com/api/webhooks/...")
  discordChannelMention: "@everyone" // Change to specific role mention if needed (e.g., "<@&ROLE_ID>")
};
