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

    const payload = JSON.stringify({
      content: `${channelMention} 🎉 **Event Scheduled**: ${fullEventTitle}!${eventLinkMessage}`,
      username: "Scheduler Bot",
      avatar_url: "https://example.com/bot-avatar.png"
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
function sendDiscordReminder(reminderEmails) {
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

    const payload = JSON.stringify({
      content: `${mentionText} 📅 **Reminder**: Please update your availability for upcoming events in the Google Sheet. We are trying to finalize the schedule for the next few weeks. Thanks!`,
      username: "Scheduler Bot",
      avatar_url: "https://example.com/bot-avatar.png"
    });

    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload
    };

    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`Discord reminder sent to ${reminderEmails.size} participants: ${[...reminderEmails].join(', ')}`);
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

    const payload = JSON.stringify({
      content: `${mentionText} ⏰ **Event Duration Notice**: Your availability constraints are limiting the event on ${eventDateStr} to less than ${CONFIG.minEventDurationHours} hours. The optimal player combination could achieve ${optimalDuration.toFixed(1)} hours. Please consider adjusting your availability if possible to allow for a longer event. Thanks!`,
      username: "Scheduler Bot",
      avatar_url: "https://example.com/bot-avatar.png"
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
