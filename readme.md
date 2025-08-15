# Developer Documentation: Google Sheets Intelligent Scheduler

## 1. Architectural Overview

This Google Apps Script employs a dual-trigger, state-driven architecture to manage event scheduling with enhanced monthly trigger support and intelligent reminder consolidation. The system is designed to be robust, efficient, and to minimize API calls, which are subject to daily quotas and performance limitations within the Apps Script environment.

### State Machine

The primary state management mechanism is the **Status** column within the Google Sheet itself. The state of each row (a potential event) dictates how it is processed. Using the spreadsheet as a transparent "database of state" is a deliberate choice. It provides at-a-glance visibility into the scheduling pipeline, allows for easy manual overrides by the sheet owner, and avoids the opacity and storage limitations of `PropertiesService`.

### UI Trigger (onEditFeedback)

A lightweight, synchronous function that runs on every edit. Its sole responsibilities are to provide immediate input validation to the user and to update the state of the edited row in the "Status" column. It performs no external API calls to `MailApp` or `CalendarApp`. This separation is critical to ensure a responsive user experience; a single, monolithic `onEdit` trigger that also handled email and calendar events would be slow, prone to hitting the 30-second execution limit, and could result in redundant API calls if multiple cells are edited in quick succession.

**Enhanced Features:**
- **Roster-aware validation**: Only counts responses from players listed in the "Player Roster" sheet
- **Advanced time format support**: Handles time formats with seconds (e.g., "18:23:00") and various combinations
- **Improved status calculation**: Accurately determines "Ready for scheduling" vs "Awaiting responses" based on actual player responses

### Processing Trigger (checkAndScheduleEvents)

A time-driven, asynchronous function that acts as the main processing engine with enhanced monthly trigger support. It reads the state of all relevant rows, applies the core scheduling and notification logic, and performs all necessary API calls. This batch-processing approach is more efficient and respectful of API quotas.

**Enhanced Features:**
- **Dynamic processing windows**: Adapts window size (14-35 days) based on time since last run
- **State persistence**: Tracks last run date to prevent gaps in coverage
- **Consolidated reminder system**: Sends only one email per participant per run, regardless of how many events they need to respond to
- **Intelligent reminder throttling**: Prevents spam by enforcing minimum 7-day intervals between reminder emails

---

## 2. Core Components & Logic

### 2.1. Configuration (CONFIG Object)

All magic numbers and sheet identifiers are stored in a global `CONFIG` object. This allows for easy maintenance and adaptation without altering the core logic. Modifying a sheet name or row layout only requires a change in one location.

- **responseSheetName, rosterSheetName, campaignDetailsSheetName**: String identifiers for the required sheets.
- **headerRow, firstDataRow**: Defines the sheet structure, allowing the script to correctly differentiate between header information and schedulable data. For example, if headers are in row 1, `headerRow` is 1 and `firstDataRow` is 2.
- **dateColumn, firstPlayerColumn**: Defines the column layout for data parsing.
- **statusColumnName**: The string header for the state machine column.
- **minEventDurationHours, shortEventWarningHours**: Business logic parameters for event validation and notifications. These control the core scheduling rules.

### 2.2. UI Trigger: onEditFeedback(e)

#### Entry Point

Triggered by any `onEdit` event in the spreadsheet.

#### Event Object (e)

Utilizes the `e.range` and `e.value` properties to get the context of the edit.

#### Execution Flow

1. **Guard Clauses**: Immediately exits if the edit is outside the designated player response area or is an edit to the Status column itself. This is a crucial performance optimization, preventing the script from executing on irrelevant cell changes.
2. **Roster Integration**: Cross-references column headers with the "Player Roster" sheet to ensure only actual players' responses are counted.
3. **Enhanced Input Validation**: For non-standard inputs (i.e., anything other than `y`, `n`, `?`), it calls `parseTimeRange()` to validate the format with support for:
   - Single times: `18`, `18:30`, `18:30:00`
   - Time ranges: `18-22`, `18:30-20:00`, `18:30:00-22:00:00`
   - An invalid format results in a `SpreadsheetApp.getUi().alert()`, providing immediate, actionable feedback to the user.
4. **Intelligent State Analysis**: After any valid edit, the function reads the entire data for the edited row, but only processes columns corresponding to actual players in the roster.
5. **Accurate State Calculation**: It tallies the counts of `y`, `?`, `n`, blank, and valid time responses from actual players only. For example, in a 4-player game with empty columns, it ignores those empty columns and only considers the 4 actual player responses.
6. **State Commit**: Writes the calculated status (e.g., Ready for scheduling, Awaiting responses) to the Status column for that row.

### 2.3. Processing Trigger: checkAndScheduleEvents()

This is the primary asynchronous worker function with enhanced monthly trigger support.

#### Dynamic Time Window Calculation

**NEW FEATURE**: The function now calculates an intelligent processing window based on when it was last run:

1. **First Run**: Defaults to a 14-day window starting 3 days from now
2. **Regular Runs**: (≤14 days since last): Uses standard 14-day window
3. **Delayed Runs** (>14 days since last): Extends window up to 35 days to catch up
4. **State Persistence**: Uses `PropertiesService` to track last run date across executions

This adaptive approach ensures:
- No events are missed due to irregular monthly trigger timing
- Complete coverage without excessive overlap
- Graceful recovery from missed or delayed triggers

#### Data Grouping

1. It filters the rows further, excluding any that have already been successfully processed (e.g., status starts with "Event created" or "Superseded").
2. The remaining rows are grouped into an `eventsByWeek` object, using `getWeekNumber()` as a key. This is the core of the "weekly analysis" logic, allowing the script to make intelligent decisions for a whole week at once, rather than processing days in isolation.

#### Weekly Processing Loop

1. **Find Best Opportunity**: It iterates through all events in the week with a status of **Ready for scheduling**. For each, it calls `calculateIntersection()` to determine the common availability. The event with the longest resulting duration is flagged as the `bestEvent`. This prioritization of duration is a core business rule designed to maximize session quality.
2. **Schedule or Handle Failures**:
   - If `bestEvent` exists: `createCalendarEvent()` is called. The status of the scheduled event's row is updated to **Event created...**. All other rows in that same week are updated to **Superseded...**. This provides clear feedback and prevents those rows from being processed again. The loop then continues to the next week.
   - If no `bestEvent` exists: The script checks if any events were **Ready for scheduling** but failed the `calculateIntersection()` check (i.e., the common time was too short). If so, their status is updated to **Failed: Duration < 2h**. This is crucial for providing feedback on why an expected event was not scheduled.
3. **Consolidated Reminder Collection**: **NEW FEATURE**: Instead of sending reminders per week, the system now:
   - Collects all unique email addresses across ALL weeks in the processing window
   - Uses a `Set` to ensure each participant is only included once
   - Sends a single consolidated reminder email per participant per run
   - Only sends reminders if it's been at least 7 days since the last run

#### Enhanced Reminder System

**MAJOR IMPROVEMENT**: The new reminder system prevents notification fatigue:

- **Global Collection**: Gathers reminder recipients across all weeks in a single `Set`
- **Deduplication**: Each participant receives at most one email per run
- **Throttling**: 7-day minimum interval between reminder campaigns
- **Smart Status Updates**: Only marks rows as "Reminder sent" if their participants actually received emails

---

## 3. Helper Functions & Algorithms

### calculateIntersection(responses, baseDate)

**Enhanced with roster-aware logic**:

1. Initializes an intersection window spanning the entire day (00:00 to 23:59).
2. Iterates through player responses, now with improved handling for different response types.
3. **NEW**: Properly distinguishes between all-Y responses (all-day events) and mixed responses.

#### Return Logic

- Returns `{ start: null, end: null }` if all valid players responded `y`. This explicitly signals a valid all-day event.
- Returns `{ start: Date, end: Date }` for a valid, timed intersection that meets the `minEventDurationHours`.
- Returns `{ start: undefined, end: undefined }` if the final intersection is invalid (e.g., start time is after end time) or shorter than the minimum duration. This signals an unschedulable event.

### parseTimeRange(timeStr, baseDate)

**ENHANCED**: Now supports extended time formats with robust validation:

- **Single times**: `18`, `18:30`, `18:30:00`
- **Time ranges**: `18-22`, `18:30-20:00`, `18:30:00-22:00:00`
- **Mixed formats**: `18-20:30`, `18:30-22`
- Uses two robust regular expressions with non-capturing groups for optional components
- Validates all time components (hours 0-23, minutes 0-59, seconds 0-59)
- Single times are automatically converted into 4-hour blocks based on `shortEventWarningHours`

### getWeekNumber(d)

A standard helper to get the ISO 8601 week number for a given date. This is critical for ensuring that weekly groupings are consistent and handle year-end transitions correctly.

### State Persistence Functions (NEW)

#### getLastRunDate()
- Retrieves the last execution date from `PropertiesService`
- Returns `null` if never run before
- Used to calculate dynamic processing windows

#### setLastRunDate(date)
- Stores the current execution date in `PropertiesService`
- Called after successful processing to track execution history

#### shouldSendReminders(week, lastRunDate)
- Enforces 7-day minimum interval between reminder campaigns
- Prevents email spam from frequent or overlapping triggers

---

## 4. Monthly Trigger Setup & Configuration

### The Challenge

Google Apps Script doesn't support true bi-weekly triggers, only monthly ones. This creates potential gaps in coverage or excessive overlap.

### The Solution

**ENHANCED SYSTEM**: Dynamic processing windows with state persistence.

#### Recommended Monthly Trigger Setup

Set up **two monthly triggers** for optimal coverage:

1. **First Trigger:**
   - Type: Time-driven
   - Event source: Month timer  
   - Day of month: **1st**
   - Time of day: 9:00 AM (or your preferred time)

2. **Second Trigger:**
   - Type: Time-driven
   - Event source: Month timer
   - Day of month: **15th** 
   - Time of day: Same time as the first trigger

#### How the Enhanced System Works

1. **Adaptive Windows**: Processing window automatically adjusts from 14-35 days based on time since last run
2. **Complete Coverage**: Even if one trigger fails, the next run extends its window to compensate
3. **No Spam**: 7-day minimum between reminder emails prevents notification fatigue
4. **Fault Tolerance**: System gracefully handles missed or delayed triggers

---

## 5. Deployment & Maintenance

### Setup Process

1. **Install the Script**: Copy the code into a Google Apps Script project bound to your spreadsheet
2. **Run Setup**: Use the "Scheduler > Setup Sheet" menu item to add the Status column
3. **Configure Triggers**:
   - Set up an "On edit" trigger for `onEditFeedback`
   - Set up two monthly triggers for `checkAndScheduleEvents` (1st and 15th of each month)
4. **Test**: Use "Scheduler > Run Now" to manually test the system

### Required Sheets Structure

#### "Player Roster" Sheet
- Column A: Player names (must match column headers in response sheet)
- Column B: Email addresses
- Column C: Notification preferences (TRUE/FALSE)

#### "Campaign details" Sheet
- Cell A2: Event title for calendar events
- Cell B2: Roll20 or meeting link for event descriptions

#### Main Response Sheet (e.g., "2025")
- Column A: Dates
- Columns C+: Player name headers (matching roster)
- Last Column: "Status" (added automatically by setup)

### Scopes

The script requires authorization for:
- `SpreadsheetApp` (read/write)
- `MailApp` (send email on behalf of the authorizing user)
- `CalendarApp` (create and modify events in the user's default calendar)
- `PropertiesService` (store last run date for adaptive windows)

### Debugging

The enhanced system includes comprehensive logging:

- Processing window calculations and decisions
- Number of participants who would receive reminders
- Reminder throttling decisions
- Calendar event creation details
- Error handling for API failures

To debug, run `checkAndScheduleEvents` manually from the Apps Script editor and view the logs under "Executions".

---

## 6. Enhanced Features Summary

### What's New

1. **Monthly Trigger Support**: Dynamic processing windows that adapt to irregular trigger timing
2. **Roster Integration**: Only counts responses from actual players, ignoring empty columns
3. **Advanced Time Parsing**: Supports seconds and various time format combinations
4. **Consolidated Reminders**: One email per participant per run, preventing spam
5. **State Persistence**: Tracks execution history for intelligent decision-making
6. **Improved Error Handling**: Graceful handling of Google service outages
7. **Enhanced Logging**: Comprehensive debugging information
8. **Manual Trigger**: "Run Now" menu item for testing and manual execution

### Performance Improvements

- **Reduced API Calls**: Consolidated reminder emails
- **Intelligent Throttling**: Prevents unnecessary reminder campaigns
- **Roster-Aware Processing**: More accurate status calculations
- **Fault Tolerance**: Automatic recovery from missed triggers

### User Experience Enhancements

- **Immediate Validation**: Real-time feedback on time format entries
- **Transparent Status**: Clear indication of scheduling pipeline state
- **Flexible Input**: Support for various time formats
- **Reduced Spam**: Maximum one reminder email per participant per run
- **Easy Setup**: One-click Status column addition

---

## 7. Troubleshooting & Common Issues

### Monthly Triggers Not Working Consistently

**Solution**: The enhanced system automatically handles irregular trigger timing. Check the logs to see the adaptive window calculations.

### Players Receiving Multiple Reminder Emails

**Solution**: This has been fixed with the consolidated reminder system. Each participant receives at most one email per run.

### Time Formats Not Recognized

**Solution**: The system now supports:
- `18` (becomes 18:00-22:00)
- `18:30` (becomes 18:30-22:30)  
- `18:30:00` (becomes 18:30-22:30)
- `18-22` (18:00-22:00)
- `18:30-22:00` (18:30-22:00)

### Status Always Shows "Awaiting Responses"

**Solution**: Ensure your player column headers exactly match the names in the "Player Roster" sheet. The system now ignores columns that don't correspond to actual players.

### Events Not Being Scheduled

**Common causes**:
1. Check that all players have responded and status shows "Ready for scheduling"
2. Verify the time intersection meets the minimum duration requirement (default: 2 hours)
3. Ensure the event date is within the processing window (3+ days from now)
4. Check the logs for detailed processing information

---

## Potential Future Enhancements

1. **Dynamic Configuration**: Move the `CONFIG` object into a dedicated "Settings" sheet for non-developer customization
2. **Calendar Integration Options**: Support for multiple calendars or calendar selection
3. **Advanced Notification Templates**: Customizable email templates with event-specific information
4. **Player Availability Analytics**: Historical analysis and availability pattern reporting
5. **Timezone Support**: Enhanced handling for players in different timezones
6. **Mobile-Friendly Interface**: Optimized UI for mobile spreadsheet editing
