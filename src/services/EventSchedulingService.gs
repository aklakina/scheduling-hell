/**
 * Service for event scheduling logic
 * Handles complex scheduling algorithms and business rules
 */

class EventSchedulingService {
  constructor(sheetService, discordService) {
    this.sheetService = sheetService;
    this.discordService = discordService;
  }

  /**
   * Main scheduling processing method
   */
  async processEvents(triggerType = 'monthly') {
    try {
      const playerInfo = this.sheetService.getPlayerRoster();
      const campaignDetails = this.sheetService.getCampaignDetails();
      const responseInfo = this.sheetService.getResponseSheetInfo();

      if (Object.keys(playerInfo).length === 0) {
        Logger.log('No players found in roster');
        return;
      }

      const windowInfo = this.calculateProcessingWindow();
      const eventsData = this.sheetService.getEventsInRange(
        windowInfo.processingStartDate,
        windowInfo.processingEndDate
      );

      if (eventsData.length === 0) {
        Logger.log(`No events to process in window: ${windowInfo.processingStartDate.toLocaleDateString()} to ${windowInfo.processingEndDate.toLocaleDateString()}`);
        return;
      }

      const eventsByWeek = this.groupEventsByWeek(eventsData);
      const notifications = new NotificationCollector();

      await this.processWeeklyEvents(
        eventsByWeek,
        playerInfo,
        campaignDetails,
        responseInfo,
        notifications,
        triggerType
      );

      await notifications.sendAll(this.discordService, triggerType);

    } catch (error) {
      Logger.log(`Error in processEvents: ${error.toString()}`);
      throw error;
    }
  }

  /**
   * Calculate dynamic processing window
   */
  calculateProcessingWindow() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Simple window calculation - can be made more complex if needed
    const processingStartDate = new Date(today);
    processingStartDate.setDate(today.getDate() + 3); // Start 3 days ahead

    const processingEndDate = new Date(today);
    processingEndDate.setDate(today.getDate() + 17); // 2 weeks window

    return {
      processingStartDate,
      processingEndDate,
      windowDays: 14
    };
  }

  /**
   * Group events by week number
   */
  groupEventsByWeek(eventsData) {
    const eventsByWeek = {};

    eventsData.forEach(event => {
      const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
      const weekNumber = this.getWeekNumber(eventDate);

      if (!eventsByWeek[weekNumber]) {
        eventsByWeek[weekNumber] = [];
      }
      eventsByWeek[weekNumber].push(event);
    });

    return eventsByWeek;
  }

  /**
   * Get week number for grouping
   */
  getWeekNumber(date) {
    const startOfYear = new Date(date.getFullYear(), 0, 1);
    const pastDaysOfYear = (date - startOfYear) / 86400000;
    return Math.ceil((pastDaysOfYear + startOfYear.getDay() + 1) / 7);
  }

  /**
   * Process events for each week
   */
  async processWeeklyEvents(eventsByWeek, playerInfo, campaignDetails, responseInfo, notifications, triggerType) {
    const numPlayers = Object.keys(playerInfo).length;

    for (const week in eventsByWeek) {
      const weekEvents = eventsByWeek[week];
      const bestEvent = this.findBestEventForWeek(weekEvents, playerInfo, responseInfo);

      if (bestEvent) {
        await this.scheduleEvent(bestEvent, weekEvents, campaignDetails, notifications);
      } else {
        await this.handleUnschedulableWeek(weekEvents, playerInfo, responseInfo, notifications, triggerType);
      }
    }
  }

  /**
   * Find the best schedulable event for a week
   */
  findBestEventForWeek(weekEvents, playerInfo, responseInfo) {
    let bestEvent = null;
    let maxDuration = 0;

    weekEvents.forEach(event => {
      const status = event.rowData[responseInfo.statusColumnIndex - 1];

      if (status === CONFIG.messages.status.readyForScheduling) {
        const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
        const responses = this.sheetService.getEventResponses(event.rowIndex, responseInfo.numPlayerColumns);

        const optimalCombination = this.findOptimalPlayerCombination(
          responses,
          responseInfo.allPlayerNames,
          playerInfo,
          eventDate
        );

        if (this.isEventSchedulable(optimalCombination, Object.keys(playerInfo).length)) {
          const duration = optimalCombination.duration * 3600000;
          if (duration > maxDuration) {
            maxDuration = duration;
            bestEvent = {
              date: eventDate,
              start: optimalCombination.intersectionStart,
              end: optimalCombination.intersectionEnd,
              rowIndex: event.rowIndex
            };
          }
        }
      }
    });

    return bestEvent;
  }

  /**
   * Check if event is schedulable based on combination results
   */
  isEventSchedulable(optimalCombination, totalPlayers) {
    return optimalCombination.duration >= CONFIG.minEventDurationHours &&
           optimalCombination.players.length === totalPlayers;
  }

  /**
   * Schedule an event and update related events
   */
  async scheduleEvent(bestEvent, weekEvents, campaignDetails, notifications) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Queue notification
    notifications.addEventNotification(
      bestEvent.date,
      bestEvent.start,
      bestEvent.end,
      campaignDetails.eventTitle,
      campaignDetails.eventLink
    );

    // Update status
    this.sheetService.updateEventStatus(
      bestEvent.rowIndex,
      `Event scheduled on ${today.toLocaleDateString()}`
    );

    // Mark other events in week as superseded
    weekEvents.forEach(event => {
      if (event.rowIndex !== bestEvent.rowIndex) {
        const currentStatus = event.rowData[this.sheetService.getResponseSheetInfo().statusColumnIndex - 1] || '';
        if (!currentStatus.startsWith(CONFIG.messages.status.eventCreated) &&
            !currentStatus.startsWith('Cancelled')) {
          this.sheetService.updateEventStatus(event.rowIndex, CONFIG.messages.status.superseded);
        }
      }
    });

    Logger.log(`Scheduled event for ${bestEvent.date.toLocaleDateString()}`);
  }

  /**
   * Handle week where no events could be scheduled
   */
  async handleUnschedulableWeek(weekEvents, playerInfo, responseInfo, notifications, triggerType) {
    const numPlayers = Object.keys(playerInfo).length;
    const minYResponsesNeeded = Math.ceil(numPlayers * CONFIG.reminderThresholdPercentage);

    weekEvents.forEach(event => {
      const status = event.rowData[responseInfo.statusColumnIndex - 1];

      if (status === CONFIG.messages.status.readyForScheduling) {
        // Mark as failed due to duration
        const failureMessage = CONFIG.messages.status.failedDuration
          .replace('{minHours}', CONFIG.minEventDurationHours);
        this.sheetService.updateEventStatus(event.rowIndex, failureMessage);
      }
      else if (status === CONFIG.messages.status.awaitingResponses) {
        this.handleAwaitingEvent(event, playerInfo, responseInfo, notifications, minYResponsesNeeded);
      }
    });
  }

  /**
   * Handle events that are awaiting responses
   */
  handleAwaitingEvent(event, playerInfo, responseInfo, notifications, minYResponsesNeeded) {
    const eventDate = new Date(event.rowData[CONFIG.dateColumn - 1]);
    const responses = this.sheetService.getEventResponses(event.rowIndex, responseInfo.numPlayerColumns);

    const yCount = this.countYesResponses(responses, responseInfo.allPlayerNames, playerInfo);

    if (yCount >= minYResponsesNeeded) {
      const reminders = this.collectReminders(responses, responseInfo.allPlayerNames, playerInfo);

      if (reminders.maybeIds.length > 0 && reminders.noResponseIds.length === 0) {
        notifications.addReminder(eventDate, reminders.maybeIds, 'maybe');
      }
      if (reminders.noResponseIds.length > 0) {
        notifications.addReminder(eventDate, reminders.noResponseIds, 'noResponse');
      }

      // Check for duration restrictions
      const optimalCombination = this.findOptimalPlayerCombination(
        responses,
        responseInfo.allPlayerNames,
        playerInfo,
        eventDate
      );

      if (this.shouldSendDurationWarning(optimalCombination, Object.keys(playerInfo).length)) {
        notifications.addDurationWarning(eventDate, optimalCombination.restrictingPlayers, playerInfo);
      }
    } else {
      // Not enough responses
      this.sheetService.updateEventStatus(
        event.rowIndex,
        `Not enough responses (${yCount}/${minYResponsesNeeded} required)`
      );
    }
  }

  /**
   * Count yes responses from valid players
   */
  countYesResponses(responses, allPlayerNames, playerInfo) {
    let yCount = 0;

    responses.forEach((response, i) => {
      const playerName = allPlayerNames[i];
      if (playerName && playerInfo[playerName]) {
        const responseStr = response ? String(response).trim().toLowerCase() : '';
        if (responseStr === CONFIG.responses.yes || this.isTimeFormat(responseStr)) {
          yCount++;
        }
      }
    });

    return yCount;
  }

  /**
   * Collect reminder information
   */
  collectReminders(responses, allPlayerNames, playerInfo) {
    const maybeIds = [];
    const noResponseIds = [];

    responses.forEach((response, i) => {
      const playerName = allPlayerNames[i];
      if (playerName && playerInfo[playerName]?.discordHandle) {
        const responseStr = response ? String(response).trim().toLowerCase() : '';

        if (responseStr === CONFIG.responses.maybe) {
          maybeIds.push(playerInfo[playerName].discordHandle);
        } else if (responseStr === CONFIG.responses.empty) {
          noResponseIds.push(playerInfo[playerName].discordHandle);
        }
      }
    });

    return { maybeIds, noResponseIds };
  }

  /**
   * Check if duration warning should be sent
   */
  shouldSendDurationWarning(optimalCombination, totalPlayers) {
    const thresholdMet = optimalCombination.players.length >=
      Math.ceil(totalPlayers * CONFIG.playerCombinationThresholdPercentage);

    return thresholdMet &&
           optimalCombination.duration >= CONFIG.minEventDurationHours &&
           optimalCombination.restrictingPlayers.length > 0;
  }

  /**
   * Find optimal player combination (simplified version)
   * This delegates to the existing complex logic in EventProcessing.gs
   */
  findOptimalPlayerCombination(responses, allPlayerNames, playerInfo, baseDate) {
    // This will call the existing function until we refactor that too
    return findOptimalPlayerCombination(responses, allPlayerNames, playerInfo, baseDate);
  }

  /**
   * Check if string is time format (simplified version)
   */
  isTimeFormat(str) {
    // This will call the existing function until we refactor Utils.gs
    return isTime(str).isValid;
  }
}

/**
 * Helper class to collect and batch notifications
 */
class NotificationCollector {
  constructor() {
    this.events = [];
    this.reminders = new Map(); // Map by date string
    this.durationWarnings = [];
  }

  addEventNotification(date, start, end, eventTitle, eventLink) {
    this.events.push({ date, start, end, eventTitle, eventLink });
  }

  addReminder(date, discordIds, reminderType) {
    const dateString = date.toLocaleDateString();
    if (!this.reminders.has(dateString)) {
      this.reminders.set(dateString, { date, maybeIds: [], noResponseIds: [] });
    }

    const reminder = this.reminders.get(dateString);
    if (reminderType === 'maybe') {
      reminder.maybeIds.push(...discordIds);
    } else if (reminderType === 'noResponse') {
      reminder.noResponseIds.push(...discordIds);
    }
  }

  addDurationWarning(date, restrictingPlayers, playerInfo) {
    this.durationWarnings.push({ date, restrictingPlayers, playerInfo });
  }

  async sendAll(discordService, triggerType) {
    // Send event notifications
    for (const event of this.events) {
      await discordService.sendEventNotification(
        event.date, event.start, event.end,
        event.eventTitle, event.eventLink, triggerType
      );
    }

    // Send reminders
    for (const [dateString, reminder] of this.reminders) {
      if (reminder.maybeIds.length > 0 && reminder.noResponseIds.length === 0) {
        await discordService.sendReminder(reminder.maybeIds, 'maybe', triggerType);
      }
      if (reminder.noResponseIds.length > 0) {
        await discordService.sendReminder(reminder.noResponseIds, 'noResponse', triggerType);
      }
    }

    // Send duration warnings
    for (const warning of this.durationWarnings) {
      await discordService.sendDurationWarning(
        warning.restrictingPlayers, warning.date,
        warning.playerInfo, triggerType
      );
    }
  }
}
