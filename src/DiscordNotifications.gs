/**
 * Discord notification functionality
 */

/**
 * Helper function to send Discord notifications for scheduled events.
 * Replaces the calendar event creation functionality.
 */
function sendDiscordEventNotification(date, start, end, eventTitle, eventLink) {
  try {
    const webhookUrl = CONFIG.discordWebhookUrl;
    const channelMention = CONFIG.discordChannelMention;

    if (!webhookUrl) {
      Logger.log(`Discord webhook URL is not set. Skipping event notification.`);
      return;
    }

    const eventDate = date.toLocaleDateString();
    const eventTime = start ? ` from ${start.toLocaleTimeString()}` : '';
    const eventEndTime = end ? ` to ${end.toLocaleTimeString()}` : '';
    const fullEventTitle = `${eventTitle} - ${eventDate}${eventTime}${eventEndTime}`;
    const eventLinkMessage = eventLink ? `\nEvent Link: ${eventLink}` : '';

    // Use message template from config
    const messageContent = CONFIG.messages.discord.eventScheduled
      .replace('{eventTitle}', fullEventTitle)
      .replace('{eventLink}', eventLinkMessage);

    const payload = JSON.stringify({
      content: `${messageContent} ${channelMention}`,
      username: CONFIG.messages.discord.botUsername,
      avatar_url: CONFIG.messages.discord.botAvatarUrl
    });

    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload
    };

    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`Discord event notification sent: ${fullEventTitle}`);
  } catch (error) {
    Logger.log(`Error sending Discord event notification: ${error.toString()}`);
  }
}

/**
 * Send Discord reminder notifications to players who haven't responded
 */
function sendDiscordReminder(reminderEmails, reminderType = 'mixed') {
  try {
    const webhookUrl = CONFIG.discordWebhookUrl;
    const channelMention = CONFIG.discordChannelMention;

    if (!webhookUrl) {
      Logger.log(`Discord webhook URL is not set. Skipping reminder for ${reminderEmails.size} participants`);
      return false;
    }

    // Create Discord mentions for players who need reminders
    const discordMentions = [];
    [...reminderEmails].forEach(discordId => {
      if (discordId && discordId.trim() !== '') {
        discordMentions.push(`<@${discordId}>`);
      }
    });

    const mentionText = discordMentions.length > 0 ? discordMentions.join(' ') : channelMention;

    // Choose message based on reminder type
    let message;
    switch (reminderType) {
      case 'maybe':
        message = CONFIG.messages.discord.reminder;
        break;
      case 'noResponse':
        message = CONFIG.messages.discord.reminderNoResponse;
        break;
      default:
        // Mixed or fallback - use the general reminder message
        message = CONFIG.messages.discord.reminder;
    }

    const payload = JSON.stringify({
      content: `${message} ${mentionText}`,
      username: CONFIG.messages.discord.botUsername,
      avatar_url: CONFIG.messages.discord.botAvatarUrl
    });

    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload
    };

    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`Discord reminder (${reminderType}) sent to ${reminderEmails.size} participants: ${[...reminderEmails].join(', ')}`);
    return true;
  } catch (error) {
    Logger.log(`Error sending Discord reminder: ${error.toString()}`);
    return false;
  }
}

/**
 * Send Discord notification to players who are restricting event duration below threshold
 */
function sendDiscordDurationRestrictionNotification(restrictingPlayers, eventDate, optimalDuration, playerInfo) {
  try {
    const webhookUrl = CONFIG.discordWebhookUrl;

    if (!webhookUrl) {
      Logger.log(`Discord webhook URL is not set. Skipping duration restriction notification.`);
      return false;
    }

    // Create Discord mentions for restricting players
    const discordMentions = [];
    restrictingPlayers.forEach(playerName => {
      if (playerInfo[playerName] && playerInfo[playerName].discordHandle) {
        discordMentions.push(`<@${playerInfo[playerName].discordHandle}>`);
      }
    });

    if (discordMentions.length === 0) {
      Logger.log('No Discord handles found for restricting players.');
      return false;
    }

    const mentionText = discordMentions.join(' ');
    const eventDateStr = eventDate.toLocaleDateString();

    // Use message template from config
    const messageContent = CONFIG.messages.discord.durationRestriction
      .replace('{eventDate}', eventDateStr)
      .replace('{minHours}', CONFIG.minEventDurationHours)
      .replace('{optimalHours}', optimalDuration.toFixed(1));

    const payload = JSON.stringify({
      content: `${messageContent} ${mentionText}`,
      username: CONFIG.messages.discord.botUsername,
      avatar_url: CONFIG.messages.discord.botAvatarUrl
    });

    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload
    };

    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`Discord duration restriction notification sent to ${restrictingPlayers.length} players: ${restrictingPlayers.join(', ')}`);
    return true;
  } catch (error) {
    Logger.log(`Error sending Discord duration restriction notification: ${error.toString()}`);
    return false;
  }
}

/**
 * Send Discord notification when sheet setup is completed
 */
function sendDiscordSheetSetupNotification() {
  try {
    const webhookUrl = CONFIG.discordWebhookUrl;
    const channelMention = CONFIG.discordChannelMention;

    if (!webhookUrl) {
      Logger.log(`Discord webhook URL is not set. Skipping sheet setup notification.`);
      return false;
    }

    // Get the current spreadsheet URL
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUrl = ss.getUrl();

    // Use message template from config
    const messageContent = CONFIG.messages.discord.sheetSetup
      .replace('{sheetUrl}', sheetUrl);

    const payload = JSON.stringify({
      content: `${messageContent} ${channelMention}`,
      username: CONFIG.messages.discord.botUsername,
      avatar_url: CONFIG.messages.discord.botAvatarUrl
    });

    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload
    };

    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`Discord sheet setup notification sent with URL: ${sheetUrl}`);
    return true;
  } catch (error) {
    Logger.log(`Error sending Discord sheet setup notification: ${error.toString()}`);
    return false;
  }
}
