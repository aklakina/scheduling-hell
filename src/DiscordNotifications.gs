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

/**
 * Send grouped Discord notifications organized by date
 * This prevents multiple notifications for the same date
 *
 * @param {Object} notificationsByDate - Object with dates as keys and message groups as values
 * @param {Object} playerInfo - Player information including discord handles and mention preferences
 * @returns {Boolean} - True if notifications were sent successfully
 */
function sendGroupedDiscordNotifications(notificationsByDate, playerInfo) {
  try {
    const webhookUrl = CONFIG.discordWebhookUrl;
    const channelMention = CONFIG.discordChannelMention;

    if (!webhookUrl) {
      Logger.log(`Discord webhook URL is not set. Skipping grouped notifications.`);
      return false;
    }

    // Process each date's notifications
    for (const dateString in notificationsByDate) {
      const dateInfo = notificationsByDate[dateString];
      const dateMessages = dateInfo.messages || [];

      if (dateMessages.length === 0) continue;

      // Build a consolidated message for this date
      let messageContent = `**📅 ${dateString}**\n\n`;
      let sentReminders = false;

      // Process each message type for this date
      dateMessages.forEach(msg => {
        switch (msg.type) {
          case 'event':
            const eventTime = msg.start ? ` from ${msg.start.toLocaleTimeString()}` : '';
            const eventEndTime = msg.end ? ` to ${msg.end.toLocaleTimeString()}` : '';
            const fullEventTitle = `${msg.eventTitle}${eventTime}${eventEndTime}`;
            const eventLinkMessage = msg.eventLink ? `\nEvent Link: ${msg.eventLink}` : '';

            messageContent += `🎉 **Event Scheduled**: ${fullEventTitle}${eventLinkMessage}\n`;
            // Add channel mention for event notifications
            messageContent += `${channelMention}\n\n`;
            break;

          case 'reminder':
            // Add different reminder messages based on type
            if (msg.reminderType === 'maybe') {
              messageContent += `${CONFIG.messages.discord.reminder}\n`;

              // Add player mentions directly to this message - check mention preference
              const maybePlayers = [];
              msg.players.forEach(player => {
                if (player && player.trim() !== '') {
                  const playerName = getPlayerNameFromDiscordId(player, playerInfo);
                  const shouldMention = shouldMentionPlayer(player, playerInfo);

                  if (shouldMention) {
                    maybePlayers.push(`<@${player}>`);
                  } else {
                    maybePlayers.push(playerName || player);
                  }
                }
              });

              if (maybePlayers.length > 0) {
                messageContent += `${maybePlayers.join(' ')}\n\n`;
              }
              sentReminders = true;
            } else if (msg.reminderType === 'noResponse') {
              messageContent += `${CONFIG.messages.discord.reminderNoResponse}\n`;

              // Add player mentions directly to this message - check mention preference
              const noResponsePlayers = [];
              msg.players.forEach(player => {
                if (player && player.trim() !== '') {
                  const playerName = getPlayerNameFromDiscordId(player, playerInfo);
                  const shouldMention = shouldMentionPlayer(player, playerInfo);

                  if (shouldMention) {
                    noResponsePlayers.push(`<@${player}>`);
                  } else {
                    noResponsePlayers.push(playerName || player);
                  }
                }
              });

              if (noResponsePlayers.length > 0) {
                messageContent += `${noResponsePlayers.join(' ')}\n\n`;
              }
              sentReminders = true;
            }
            break;

          case 'restriction':
            messageContent += `${CONFIG.messages.discord.durationRestriction}\n`;

            // Add player mentions for restriction notices directly to this message
            const restrictionMentions = [];
            if (msg.players && msg.players.length > 0) {
              msg.players.forEach(playerName => {
                const player = playerInfo[playerName];
                if (player && player.discordHandle) {
                  if (player.allowMention) {
                    restrictionMentions.push(`<@${player.discordHandle}>`);
                  } else {
                    restrictionMentions.push(playerName);
                  }
                }
              });
            }

            if (restrictionMentions.length > 0) {
              messageContent += `${restrictionMentions.join(' ')}\n\n`;
            } else {
              messageContent += '\n';
            }
            break;
        }
      });

      // Remove any trailing newlines if they exist
      messageContent = messageContent.trimEnd();

      // Send the consolidated message for this date
      const payload = JSON.stringify({
        content: messageContent,
        username: CONFIG.messages.discord.botUsername,
        avatar_url: CONFIG.messages.discord.botAvatarUrl
      });

      const options = {
        method: "post",
        contentType: "application/json",
        payload: payload
      };

      UrlFetchApp.fetch(webhookUrl, options);
      Logger.log(`Sent grouped Discord notification for date: ${dateString} with ${dateMessages.length} message types`);

      // Mark if we sent reminders for this date
      notificationsByDate[dateString].sentReminders = sentReminders;
    }

    return true;
  } catch (error) {
    Logger.log(`Error sending grouped Discord notifications: ${error.toString()}`);
    return false;
  }
}

/**
 * Check if player should be mentioned based on their preferences
 * @param {string} discordId - Discord ID of the player
 * @param {Object} playerInfo - Player info object containing preferences
 * @returns {boolean} - True if player should be mentioned
 */
function shouldMentionPlayer(discordId, playerInfo) {
  // Find player by discord ID
  for (const playerName in playerInfo) {
    const player = playerInfo[playerName];
    if (player.discordHandle === discordId) {
      // Only mention if they have both a discord handle AND have allowed mentions
      return player.discordHandle && player.allowMention;
    }
  }
  return false; // Default to false if player not found
}

/**
 * Get player name from their discord ID
 * @param {string} discordId - Discord ID of the player
 * @param {Object} playerInfo - Player info object
 * @returns {string} - Player name or null if not found
 */
function getPlayerNameFromDiscordId(discordId, playerInfo) {
  for (const playerName in playerInfo) {
    const player = playerInfo[playerName];
    if (player.discordHandle === discordId) {
      return playerName;
    }
  }
  return null;
}
