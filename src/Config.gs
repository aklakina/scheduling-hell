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

  // NEW: Trigger configuration
  triggers: {
    daily: {
      enabled: true,
      windowDays: 7,  // Check next 7 days for ready events
      allowEventNotifications: true,  // Can send event scheduled notifications
      allowReminders: false,  // Don't send reminders on daily
      allowDurationWarnings: false  // Don't send duration warnings on daily
    },
    biWeekly: {
      enabled: true,
      daysAhead: 3,  // Start checking 3 days from Monday
      windowWeeks: 2,  // Check next 2 weeks
      allowEventNotifications: false,  // Don't schedule events on bi-weekly
      allowReminders: true,  // Send reminders on bi-weekly
      allowDurationWarnings: true  // Send duration warnings on bi-weekly
    },
    monthly: {
      enabled: true,  // Keep existing monthly system
      allowEventNotifications: true,
      allowReminders: true,
      allowDurationWarnings: true
    }
  },

  // Discord webhook configuration
  discordWebhookUrl: PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK'), // Add your Discord webhook URL here (e.g., "https://discord.com/api/webhooks/...")
  discordChannelMention: "@everyone", // Change to specific role mention if needed (e.g., "<@&ROLE_ID>")

  // Player response constants
  responses: {
    yes: "y",
    no: "n",
    maybe: "?",
    empty: ""
  },

  // Column identifiers
  columns: {
    today: "Today",
    day: "Day"
  },

  // --- Message Templates ---
  // Discord notification messages
  messages: {
    discord: {
      // Event scheduled notification
      eventScheduled: "🎉 **Event Scheduled**: {eventTitle}!{eventLink}",

      // Reminder notification for players who responded with '?'
      reminder: "**?**: No pressure, but the event depends on you only.",

      // Reminder for players who haven't responded
      reminderNoResponse: "⏰ **Public Shaming announcement**: Fill out the sheet you lazy ahh!",

      // Duration restriction notification for players limiting event length
      durationRestriction: "⏰ **Event Duration Notice**: Please review the sheet to see we can match hours!",

      // Sheet setup notification with invite link
      sheetSetup: "🎮 **New scheduling sheet is ready!** Fill out the sheet here: {sheetUrl}",

      // Bot configuration
      botUsername: "Scheduler Bot",
      botAvatarUrl: "https://example.com/bot-avatar.png"
    },

    // UI alert messages
    ui: {
      invalidTimeFormat: {
        title: "Invalid Time Format",
        message: "Your entry \"{userInput}\" is not a valid time or time range. Please use formats like \"18:00\", \"18-22\", or \"18:30-22:00\"."
      },

      invalidDate: "Error: Could not find a valid date in column A for this row.",

      setupSheet: {
        rosterNotFound: "Error: The \"{rosterSheetName}\" sheet was not found or has no player data. Please create the roster sheet first with player names in column A.",
        noPlayersFound: "Error: No player names found in the \"{rosterSheetName}\" sheet. Please add player names in column A starting from row 2.",
        confirmRecreate: "The \"{responseSheetName}\" sheet already exists. Do you want to recreate it with the current roster structure? This will delete all existing data.",
        setupCancelled: "Setup cancelled. No changes were made."
      }
    },

    // Status messages for the sheet
    status: {
      cancelled: "Cancelled (No consensus)",
      readyForScheduling: "Ready for scheduling",
      awaitingResponses: "Awaiting responses",
      eventScheduled: "Event scheduled on {date}",
      superseded: "Superseded by other event",
      failedDuration: "Failed: Duration < {minHours}h",
      reminderSent: "Reminder sent on {date}",
      eventCreated: "Event created", // Added for checks in the code
      notEnoughResponses: "Not enough responses ({current}/{required} required)" // New status for insufficient response rate
    },

    // Data validation help text
    validation: {
      playerResponseHelp: "Quick select: Y (yes), N (no), ? (maybe), or enter time range (e.g., 18-22)"
    },

    // Archive sheet formatting
    archive: {
      dateColumnHeader: "Date (Archived)",
      statusColumnHeader: "Final Status"
    }
  }
};
