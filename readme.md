# Developer Documentation: Google Sheets Discord Scheduler

## 1. Architectural Overview

This Google Apps Script employs a dual-trigger, state-driven architecture to manage event scheduling with **Discord notifications**, automatic archiving, and intelligent player combination analysis. The system is designed to be robust, efficient, and to minimize API calls while providing rich Discord-based notifications for scheduling updates.

### Key System Changes

**MAJOR UPDATE**: This system now uses **Discord webhooks** for all notifications instead of Google Calendar events or email. The system sends Discord messages for:
- Event scheduling notifications
- Player reminders (differentiated by response type)
- Duration restriction warnings
- Sheet setup notifications

### State Machine

The primary state management mechanism is the **Status** column within the Google Sheet itself. The state of each row (a potential event) dictates how it is processed. Using the spreadsheet as a transparent "database of state" is a deliberate choice. It provides at-a-glance visibility into the scheduling pipeline, allows for easy manual overrides by the sheet owner, and avoids the opacity and storage limitations of `PropertiesService`.

### UI Trigger (onEditFeedback)

A lightweight, synchronous function that runs on every edit. Its sole responsibilities are to provide immediate input validation to the user and to update the state of the edited row in the "Status" column. It performs no external API calls to Discord or other services. This separation is critical to ensure a responsive user experience.

**Enhanced Features:**
- **Roster-aware validation**: Only counts responses from players listed in the "Player Roster" sheet with Discord handles
- **Advanced time format support**: Handles time formats with seconds (e.g., "18:23:00") and various combinations
- **Improved status calculation**: Accurately determines "Ready for scheduling" vs "Awaiting responses" based on actual player responses

### Processing Trigger (checkAndScheduleEvents)

A time-driven, asynchronous function that acts as the main processing engine with enhanced monthly trigger support. It reads the state of all relevant rows, applies the core scheduling and notification logic, and performs Discord notifications and automatic data management.

**Enhanced Features:**
- **Dynamic processing windows**: Adapts window size (14-35 days) based on time since last run
- **State persistence**: Tracks last run date to prevent gaps in coverage
- **Discord integration**: Sends targeted Discord notifications instead of emails
- **Intelligent reminder system**: Differentiated reminders for "maybe" vs "no response" players
- **Player combination analysis**: Identifies players restricting event duration
- **Automatic archiving**: Moves old data to archive sheet automatically
- **Future date generation**: Automatically creates new date rows to maintain 2-month scheduling window

---

## 2. Core Components & Logic

### 2.1. Configuration (CONFIG Object)

All configuration is stored in a global `CONFIG` object in `Config.gs`. This includes:

- **Sheet identifiers**: `responseSheetName`, `rosterSheetName`, `campaignDetailsSheetName`, `archiveSheetName`
- **Layout configuration**: `headerRow`, `firstDataRow`, `dateColumn`, `firstPlayerColumn`
- **Business logic parameters**: 
  - `minEventDurationHours` (4 hours) - Minimum for actual scheduling
  - `minConsiderationDurationHours` (2 hours) - Minimum to consider viable
  - `reminderThresholdPercentage` (0.4) - % of Y responses needed before sending reminders
  - `playerCombinationThresholdPercentage` (0.6) - % of players needed for duration notifications
- **Discord configuration**: Webhook URL, channel mentions, message templates
- **Automation settings**: Archive threshold, future date creation period

### 2.2. Discord Integration System

#### Discord Webhook Setup
The system requires a Discord webhook URL stored in Script Properties under the key `DISCORD_WEBHOOK`. This enables:

1. **Event Notifications**: Rich Discord messages when events are scheduled
2. **Targeted Reminders**: Separate messages for players who answered "?" vs those who didn't respond
3. **Duration Warnings**: Notifications to players whose time constraints are limiting event length
4. **Setup Notifications**: Messages when new scheduling sheets are created

#### Player Roster Integration
The "Player Roster" sheet now requires:
- **Column A**: Player names (matching response sheet headers)
- **Column B**: Discord user IDs (for mentions in notifications)
- **Column C**: Notification preferences (optional, for future use)

### 2.3. UI Trigger: onEditFeedback(e)

#### Execution Flow
1. **Guard Clauses**: Exits if edit is outside player response area or in system columns
2. **Roster Integration**: Cross-references with Discord-enabled players only
3. **Enhanced Input Validation**: Supports extended time formats with real-time feedback
4. **Intelligent State Analysis**: Processes only actual roster players, ignoring empty columns
5. **Accurate State Calculation**: Counts responses from Discord-enabled players only

### 2.4. Processing Trigger: checkAndScheduleEvents()

#### Enhanced Monthly Processing
The system now includes sophisticated monthly trigger handling with adaptive windows and comprehensive data management.

#### Player Combination Analysis
**NEW FEATURE**: The system analyzes optimal player combinations:
1. **Restricting Player Detection**: Identifies players whose time constraints limit event duration
2. **Threshold Analysis**: Checks if 60%+ of players can meet minimum duration
3. **Targeted Notifications**: Sends Discord messages to players restricting duration
4. **Optimal Scheduling**: Finds best player combinations for maximum event length

#### Discord Notification System
- **Event Scheduling**: Rich Discord messages with event details and mentions
- **Differentiated Reminders**: 
  - "Maybe" reminders for players who answered "?"
  - "No response" reminders for players who haven't responded
- **Duration Restrictions**: Targeted messages to players limiting event length
- **Throttling**: 7-day minimum between reminder campaigns

#### Automatic Data Management
**NEW FEATURES**:
1. **Auto-Archiving**: Moves old data (older than 1 week) to Archive sheet
2. **Future Date Creation**: Automatically generates dates for next 2 months
3. **Sheet Formatting**: Applies comprehensive formatting to both active and archive sheets

---

## 3. Enhanced Helper Functions & Algorithms

### 3.1. Player Combination Analysis

#### findOptimalPlayerCombination(responses, allPlayerNames, playerInfo, baseDate)
Advanced algorithm that:
1. **Identifies restricting players** whose time constraints are below minimum duration
2. **Finds optimal combinations** excluding restricting players
3. **Validates thresholds** ensuring sufficient player participation
4. **Returns detailed analysis** including duration and restricting player list

#### findRestrictingPlayers(validPlayerResponses, baseDate)
Specifically identifies players whose individual time constraints fall below the 4-hour minimum event duration, enabling targeted notifications.

### 3.2. Discord Notification Functions

#### sendDiscordEventNotification(date, start, end, eventTitle, eventLink)
Sends rich Discord messages for scheduled events with:
- Event title and date formatting
- Time range information
- Roll20/meeting links
- Channel mentions for visibility

#### sendDiscordReminder(reminderEmails, reminderType)
Sends targeted reminders with type-specific messaging:
- `'maybe'`: For players who answered "?"
- `'noResponse'`: For players who haven't responded
- Prevents spam through deduplication and throttling

#### sendDiscordDurationRestrictionNotification(restrictingPlayers, eventDate, optimalDuration, playerInfo)
Notifies specific players when their time constraints are preventing optimal scheduling.

### 3.3. Archive and Data Management

#### archiveOldResponses(ss, processingStartDate)
Automatically moves old response data to Archive sheet:
- **Threshold-based**: Archives data older than 1 week
- **Preserves structure**: Maintains formatting and column layout
- **Batch processing**: Efficient bulk operations
- **Automatic formatting**: Applies archive-specific styling

#### createFutureDateRows(ss, today)
Automatically generates future scheduling dates:
- **2-month window**: Always maintains dates for next 2 months
- **Daily scheduling**: Creates consecutive daily entries
- **Formula integration**: Adds day-of-week and "today" indicator formulas
- **Structure preservation**: Maintains all column formatting and validation

### 3.4. Enhanced Sheet Formatting

#### formatResponseSheet() / formatArchiveSheet()
Comprehensive formatting systems providing:
- **Conditional formatting**: Color-coded responses and statuses
- **Data validation**: Dropdown menus for quick response entry
- **Column optimization**: Proper widths and alignment
- **Visual hierarchy**: Clear distinction between data types
- **Archive styling**: Muted colors for historical data

---

## 4. Monthly Trigger System & Discord Setup

### 4.1. Discord Webhook Configuration

#### Required Setup Steps:
1. **Create Discord Webhook**: In your Discord server, create a webhook for the target channel
2. **Store Webhook URL**: Add to Script Properties as `DISCORD_WEBHOOK`
3. **Configure Mentions**: Set `CONFIG.discordChannelMention` to appropriate role or @everyone
4. **Test Notifications**: Use "Run Now" to verify Discord integration

#### Message Customization:
All Discord messages are template-based in `CONFIG.messages.discord`:
- `eventScheduled`: Event notification format
- `reminder` / `reminderNoResponse`: Different reminder types  
- `durationRestriction`: Duration warning messages
- `sheetSetup`: New sheet notifications

### 4.2. Monthly Trigger Configuration

#### Recommended Setup:
1. **First Trigger**: 1st of each month, 9:00 AM
2. **Second Trigger**: 16th of each month, same time
3. **Adaptive Windows**: System automatically adjusts processing window based on timing

#### How It Works:
- **1st of month trigger**: Processes ~3-45 days ahead (adaptive)
- **16th of month trigger**: Continues coverage with overlap prevention
- **State persistence**: Tracks execution history to prevent gaps
- **Discord throttling**: 7-day minimum between reminder campaigns

---

## 5. Sheet Structure & Setup

### 5.1. Required Sheets

#### "Player Roster" Sheet
| Column | Content | Purpose |
|--------|---------|---------|
| A | Player Name | Must match response sheet headers exactly |
| B | Discord User ID | For @mentions in notifications (e.g., "123456789012345678") |
| C | Notification Preferences | Optional, for future features |

#### "Campaign details" Sheet  
| Cell | Content | Purpose |
|------|---------|---------|
| A2 | Event Title | Used in Discord notifications |
| B2 | Roll20/Meeting Link | Included in event notifications |

#### "Responses" Sheet (Auto-generated)
- **Column A**: Dates (auto-generated)
- **Column B**: Day of week (formula-based)
- **Columns C+**: Player response columns (from roster)
- **Today Column**: Shows current date indicator
- **Status Column**: System-managed status tracking

#### "Archive" Sheet (Auto-created)
- **Same structure** as Responses sheet
- **Historical data** older than 1 week
- **Muted formatting** for archived content
- **No "Today" column** (removed during archiving)

### 5.2. Setup Process

1. **Create Player Roster**: Add player names and Discord IDs
2. **Create Campaign Details**: Set event title and meeting link
3. **Run Setup**: Use "Scheduler > Setup Sheet" menu item
4. **Configure Discord**: Add webhook URL to Script Properties
5. **Set Triggers**: Configure monthly triggers (1st and 16th)
6. **Test System**: Use "Scheduler > Run Now" for testing

---

## 6. Advanced Features & Enhancements

### 6.1. Player Combination Intelligence

The system now performs sophisticated analysis to optimize scheduling:

- **Minimum viable combinations**: Requires 60% of players for duration notifications
- **Restricting player identification**: Finds players limiting event length
- **Optimal duration calculation**: Maximizes session length within constraints
- **Targeted messaging**: Notifies only relevant players about constraints

### 6.2. Automated Data Lifecycle

**Archive Management**:
- Automatically moves data older than 1 week to Archive sheet
- Preserves complete response history
- Applies specialized formatting for historical data
- Maintains sheet performance by limiting active data

**Future Date Generation**:
- Always maintains 2 months of future dates
- Creates daily entries with proper formulas
- Integrates with existing formatting systems
- Ensures no gaps in scheduling availability

### 6.3. Enhanced User Experience

**Visual Feedback**:
- Real-time conditional formatting for responses
- Status-based color coding
- Data validation with helpful hints
- Clear visual hierarchy between active and archived data

**Smart Notifications**:
- Type-specific Discord reminders
- Threshold-based reminder triggering
- Duration constraint warnings
- Setup completion notifications

---

## 7. Troubleshooting & Common Issues

### 7.1. Discord Integration Issues

**Webhook Not Working**:
- Verify webhook URL in Script Properties under `DISCORD_WEBHOOK`
- Test webhook with external tool (e.g., curl)
- Check Discord server permissions

**Players Not Getting Mentioned**:
- Ensure Discord User IDs are correct in Player Roster (Column B)
- Verify ID format (should be numbers only, e.g., "123456789012345678")
- Check Discord privacy settings for mentioned users

**Messages Not Sending**:
- Review script execution logs for error messages
- Verify webhook URL hasn't expired
- Check Discord server status

### 7.2. Scheduling Issues

**Status Always "Awaiting Responses"**:
- Verify player names in roster exactly match response sheet headers
- Check that Discord User IDs are present for all players
- Ensure responses are in correct format (Y, N, ?, or time ranges)

**Events Not Being Scheduled**:
- Check that all players have responded
- Verify intersection meets 4-hour minimum duration
- Ensure event date is within processing window (3+ days ahead)
- Review logs for optimal player combination analysis

**Archive/Future Date Issues**:
- Check CONFIG settings for archive and future date parameters
- Verify sheet structure hasn't been manually modified
- Review execution logs for specific error messages

### 7.3. Performance Issues

**Slow Execution**:
- Large data sets may need archive threshold adjustment
- Consider reducing future date creation period
- Monitor Google Apps Script execution time limits

**Formatting Problems**:
- Re-run formatting functions manually via menu
- Check for manually modified sheet structure
- Verify conditional formatting rules aren't conflicting

---

## 8. Future Enhancement Roadmap

### 8.1. Short-term Improvements
- **Multi-timezone support**: Handle players in different time zones
- **Calendar integration**: Optional Google Calendar event creation alongside Discord
- **Advanced templates**: Customizable Discord message templates
- **Player analytics**: Availability pattern analysis and reporting

### 8.2. Medium-term Features
- **Multi-campaign support**: Handle multiple gaming groups/campaigns
- **Role-based permissions**: Different access levels for players vs organizers
- **Advanced scheduling**: Support for recurring events and flexible durations
- **Integration APIs**: Webhooks for external systems (Roll20, etc.)

### 8.3. Long-term Vision
- **Web interface**: Browser-based configuration and monitoring
- **Mobile optimization**: Better mobile spreadsheet experience
- **AI assistance**: Intelligent scheduling recommendations
- **Community features**: Shared availability calendars and group coordination

---

## 9. Technical Implementation Notes

### 9.1. Architecture Decisions

**Discord over Email/Calendar**:
- Immediate notification delivery
- Rich message formatting capabilities
- Better integration with gaming communities
- Reduced dependency on Google services

**Spreadsheet as Database**:
- Transparent state management
- Easy manual intervention capability
- Built-in version control through Google Sheets
- No additional database requirements

**Modular Code Structure**:
- Separated concerns across multiple .gs files
- Configurable message templates
- Extensible notification system
- Maintainable codebase organization

### 9.2. Performance Optimizations

**Batch Processing**:
- Single Discord message per participant per run
- Bulk archive operations
- Efficient sheet range operations
- Minimized API calls

**Intelligent Scheduling**:
- Adaptive processing windows
- State-based execution control
- Conditional processing logic
- Resource-conscious operation

This comprehensive documentation reflects the current state of the Discord-integrated scheduling system with automatic data management and enhanced player coordination features.
