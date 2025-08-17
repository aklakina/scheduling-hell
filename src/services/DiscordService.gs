/**
 * Service for Discord notifications
 * Centralizes all Discord communication logic
 */

class DiscordService {
  constructor() {
    this.webhookUrl = CONFIG.discordWebhookUrl;
    this.channelMention = CONFIG.discordChannelMention;
    this.botConfig = CONFIG.messages.discord;
  }

  /**
   * Send Discord notification with error handling
   */
  async sendNotification(content, mentions = []) {
    if (!this.webhookUrl) {
      Logger.log('Discord webhook URL not configured');
      return false;
    }

    try {
      const mentionText = mentions.length > 0 ? mentions.join(' ') : this.channelMention;

      const payload = JSON.stringify({
        content: `${content} ${mentionText}`,
        username: this.botConfig.botUsername,
        avatar_url: this.botConfig.botAvatarUrl
      });

      const options = {
        method: "post",
        contentType: "application/json",
        payload: payload
      };

      UrlFetchApp.fetch(this.webhookUrl, options);
      Logger.log(`Discord notification sent: ${content}`);
      return true;
    } catch (error) {
      Logger.log(`Error sending Discord notification: ${error.toString()}`);
      return false;
    }
  }

  /**
   * Send event scheduled notification
   */
  async sendEventNotification(date, start, end, eventTitle, eventLink, triggerType = 'monthly') {
    const triggerConfig = CONFIG.triggers[triggerType];
    if (!triggerConfig?.allowEventNotifications) {
      Logger.log(`Event notifications disabled for trigger type: ${triggerType}`);
      return false;
    }

    const eventDate = date.toLocaleDateString();
    const eventTime = start ? ` from ${start.toLocaleTimeString()}` : '';
    const eventEndTime = end ? ` to ${end.toLocaleTimeString()}` : '';
    const fullEventTitle = `${eventTitle} - ${eventDate}${eventTime}${eventEndTime}`;
    const eventLinkMessage = eventLink ? `\nEvent Link: ${eventLink}` : '';

    const messageContent = this.botConfig.eventScheduled
      .replace('{eventTitle}', fullEventTitle)
      .replace('{eventLink}', eventLinkMessage);

    return await this.sendNotification(messageContent);
  }

  /**
   * Send reminder notification
   */
  async sendReminder(discordIds, reminderType = 'mixed', triggerType = 'monthly') {
    const triggerConfig = CONFIG.triggers[triggerType];
    if (!triggerConfig?.allowReminders) {
      Logger.log(`Reminder notifications disabled for trigger type: ${triggerType}`);
      return false;
    }

    const mentions = discordIds.map(id => `<@${id}>`).filter(mention => mention !== '<@>');

    let message;
    switch (reminderType) {
      case 'maybe':
        message = this.botConfig.reminder;
        break;
      case 'noResponse':
        message = this.botConfig.reminderNoResponse;
        break;
      default:
        message = this.botConfig.reminder;
    }

    return await this.sendNotification(message, mentions);
  }

  /**
   * Send duration restriction notification
   */
  async sendDurationWarning(restrictingPlayers, eventDate, playerInfo, triggerType = 'monthly') {
    const triggerConfig = CONFIG.triggers[triggerType];
    if (!triggerConfig?.allowDurationWarnings) {
      Logger.log(`Duration warning notifications disabled for trigger type: ${triggerType}`);
      return false;
    }

    const mentions = restrictingPlayers
      .filter(playerName => playerInfo[playerName]?.discordHandle)
      .map(playerName => `<@${playerInfo[playerName].discordHandle}>`);

    if (mentions.length === 0) {
      Logger.log('No Discord handles found for restricting players');
      return false;
    }

    const message = this.botConfig.durationRestriction;
    return await this.sendNotification(message, mentions);
  }

  /**
   * Send sheet setup notification
   */
  async sendSheetSetupNotification() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUrl = ss.getUrl();

    const message = this.botConfig.sheetSetup.replace('{sheetUrl}', sheetUrl);
    return await this.sendNotification(message);
  }
}
