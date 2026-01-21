# google-calendar-migrate
Google CalendarEvents cannot be easily copied at bulk. In theory you can export into iCal format, and use that iCal to import into your new calendar. Try this first! (For me it did not work out at all.) Escpecially if you want to copy CalendarEvents cross-Google-Workspace, Google cannot help you out. (This has not been tested with personal Google accounts yet, let me know...). I change towards a new Google Workspace, and consider my CalendarEvents a valuable source of history, knowledge an some facts.

This script copies CalendarEvents from one Google Calendar to another Google Calendar of yours, which can be cross-Google-Workspace. You define the period of CalendarEvents to copy from your sourceCalendar into your targetCalendar. If an Event already exists, it will NOT be copied. It does understand recurring events, birthdays, Out-of-Office etc. Location events are NOT copied yet (For me they were too easy to recreate, and historically they do not provide much value (to me). 

## Todo
Currently, the script transforms all-day-events containing the substring 'jubileum' or 'verjaardag' into a yearly recurring all-day-event. These substrings are not easily/centrally configurable yet.

# Prerequisites
* Have your sourceCalendar AND your targetCalendar (having write permissions) into your Google Workspace account.
* It doens't matter what Workspace account, the source, the target or another, as long as your account has read permissions on the sourceCalendar, and write permissions on the targetCalendar.
* First time you run the script, it will ask for permissions against the targetCalendar, please grant these permissions.
* First time you run the script, it will ask for permissions against your GDrive. It needs these to store a GoogleDocument to persist your logging. You want to know when/what stuff fails. Please grant these permissions.

1. Open your webbrowser to https://script.google.com, create a new project, name it wisely, and copy the script ([[Code.js]]) into your editor. 
2. Change the values in the CONFIG object to match your case. See Configuration for details.
3. Test your settings, set CONFIG.IS_DRY_RUN to true, and validate things are running. Restore  CONFIG.IS_DRY_RUN into false again if your test is succesful.
4. Run the main() method.

# Configuration
The CONFIG object contains your configuration:
* SOURCE_CALENDAR_NAME: the id or name of the Calendar containing the CalendarEvents to copy from (e.g. sourceCalendar). Often this is an email address
* TARGET_CALENDAR_NAME: the id or name of the Calendar you want the CalendarItems to be copied into (e.g. targetCalendar). Often this is an email address
* START_DATE: the start date yyyy-MM-dd of the period to migrate (the day will start @ 00:00 hour)
* END_DATE: the end date yyyy-MM-dd of the period to migrate (the day will end @ 00:00 hour, meaning the next all-day event does is included)
* IS_DRY_RUN: do the work, but NOT do the commit. More of a running analysis of your calendar
* FETCH_NUMBER_OF_EVENTS: the number of events to read from sourceCalendar in a batch. Make this number larger than the number of invites you expect in 2 calendar days. Default: 100
* DEFAULT_EVENT_SETTINGS: default config to make migration feasible. It should NOT send invites to events (awkward wor historical events, inconvenient for future events the guest already have accepted). The second one is to allow for extended settings to copy GoogleMeet details or other conference/video settings {sendUpdates:'none', conferenceDataVersion:1}
* LOG_FILENAME: the name of the peristent log-file that will be stored in the root of your personal GDrive. This filename will be prefixed with a date-time mask (yyyyMMdd HHmmss)
* BATCH_PAUZE_INCREMENTS: Google ask you to be gentle with their API. After a batch read from sourceCalendar is processed, we take a break. A default break is BATCH_PAUZE_TIME seconds. This number defines how many of these breaks we concatinate as a period of rest. This is used to generate some kind of progress thingy in the on-screen log. (Default = 3)
* BATCH_PAUZE_TIME:  Google ask you to be gentle with their API. After a batch read from sourceCalendar is processed, we take a break. A default break is BATCH_PAUZE_TIME seconds. This number defines how many of these breaks we concatinate as a period of rest. This is used to generate some kind of progress thingy in the on-screen log. (Default is 5 (sec))
  

# Workings
## main()
1. The script will batch-read your sourceCalendar, usually in chunks of for example 50 to 100 CalendarEvents, until the batch is empty.
2. For each CalendarEvent in the batch:
   a. determine the type of event
   b. preprocess where needed
4. Create the event in the targetCalendar;
   a. persist the CalendarEvent
   b. recurring CalendarEvents will fail if a higher sequence number of that recurringEventId already exists (or even if the recurringEventId exists). If so, strip these recurring event details, and recreate the CalendarEvent using the same recurrance rule if possible.
     
## deleteTargetCalendar()
If your targetCalendar becomes polluted of trying and improving your script, you might want to bulk-clean your targetCalendar. In this method you can define a custom period to clean. It does obay the IS_DRY_RUN setting from the CONFIG though.

# Limitations
* Running an AppScript is limited in time. After some minutes your script will be auto-terminated by Google. This will be about 400+ CalendarEvents later in my experience.
* Google has limits on the use of its API: (Source: https://support.google.com/a/answer/2905486)
  
  | Limit | Penalty |
  |---|---|
  | ~10.000 invites sent to people outside primary/secondairy domain | possibly 1 month |
  | ~100.000 events created in a Calendar in short period | possibly several months |
  | ~2000 emails sent to external guests in a short period | possibly 24 hours |
  | Creation of >60 Calendars in short period | possibly several hours |
  | Sharing a Calendar with many users in a short period | possibly several hours |
  
  

