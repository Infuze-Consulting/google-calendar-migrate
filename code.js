
const CONFIG = {
  SOURCE_CALENDAR_NAME: "my.name@sourcedomain.ext",
  TARGET_CALENDAR_NAME: "my.name@targetdomain.ext", // must be your primary!!!
  START_DATE: new Date("2025-12-01"), // is the beginning of the start date YYYY-MM-dd  (@ 00:00 hour)
  END_DATE: new Date("2026-08-01"), // is the beginning of the end date YYYY-MM-dd  (@ 00:00 hour) (all-day event does is included on this end date)
  IS_DRY_RUN: false, // do the work, but not the commit. More of a running analysis of your calendar
  FETCH_NUMBER_OF_EVENTS: 100, // make this larger than the number of invites in 2 calendar days. Default: 100
  DEFAULT_EVENT_SETTINGS: {sendUpdates:'none', conferenceDataVersion:1},
  LOG_FILENAME: "MigrateCalendarEntriesLog", // will be prefixed with YYMMdd HHmmss
  BATCH_PAUZE_INCREMENTS: 3, // 3 increment introduces a delay of BATCH_PAUZE_TIME seconds, 
                              // (after each FETCH_NUMBER_OF_EVENTS to prevent flooding of the API. Default = 3
  BATCH_PAUZE_TIME: 5, // in sec. Default is 10 sec
  //MICRO_DELAY: 1500, //msec DISABLED
};
/********************************************************************/ 
/* There are limits to how hard you can push the API:               */
/* ~10.000 invites sent to people outside primary/secondairy domain */
/*    penalty: possibly 1 month                                     */
/* ~100.000 events created in a Calendar in short period            */
/*    penalty: possibly several months                              */
/* ~2000 emails sent to external guests in a short period           */
/*    penalty: possibly 24 hours                                    */
/* Creation of >60 Calendars in short period                        */
/*    penalty: possibly several hours                               */
/* Sharing a Calendar with many users in a short period             */
/*    penalty: possibly several hours                               */
/* Source: https://support.google.com/a/answer/2905486#             */
/********************************************************************/ 


function logLine(doc, line){
  const datetime = Utilities.formatDate(new Date(), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm:ss');
  doc.getBody().appendParagraph(datetime + " | " + line);
  Logger.log(line);
}


function main() {
  const doc = DocumentApp.create(Utilities.formatDate(new Date(), 'Europe/Amsterdam', 'YYYYMMdd HHmmss') + "_"+ CONFIG.LOG_FILENAME);
  const sourceCals = CalendarApp.getCalendarsByName(CONFIG.SOURCE_CALENDAR_NAME);
  const targetCals = CalendarApp.getCalendarsByName(CONFIG.TARGET_CALENDAR_NAME);
  var nextStartDate = CONFIG.START_DATE;

  if (sourceCals.length === 0) {
    Logger.log("Error: Source calendar '" + CONFIG.SOURCE_CALENDAR_NAME + "' not found.");
    return;
  }
  if (targetCals.length === 0) {
    Logger.log("Error: Target calendar '" + CONFIG.TARGET_CALENDAR_NAME + "' not found.");
    return;
  }

  const sourceCal = sourceCals[0];
  const targetCal = targetCals[0];
  
  logLine(doc, "Migrating Calendar events from "+ CONFIG.SOURCE_CALENDAR_NAME +" to " + CONFIG.TARGET_CALENDAR_NAME);
  logLine(doc, "Fetching events from " + CONFIG.START_DATE.toDateString() + " to " + CONFIG.END_DATE.toDateString());
  
  // fetch a number of invites, starting from nextStartDate. Process each event in this list. 
  // Return the last startDate as the new nextStartDate. Continue until there are no more calendar events in the list returned. 
  var loops=1;
  var returnObject = null; //{lastStart, lastEtag}
  var lastId = "";

  while (nextStartDate){
    Logger.log("Entering loop: "+ loops + " current startDate: " + nextStartDate + " lastId: " + lastId);
    returnObject = listNextEvents(doc, nextStartDate, CONFIG.END_DATE, targetCal, loops);
    if (returnObject){
      nextStartDate = returnObject.lastStart;
      if (lastId.includes(returnObject.lastId)) {
        nextStartDate=undefined;
      } else {
        lastId = returnObject.lastId;
      }
    } else {
      nextStartDate=undefined;
    }
    loops++;
    if (nextStartDate){
      Logger.log("Pauze for " + Math.round(CONFIG.BATCH_PAUZE_INCREMENTS * CONFIG.BATCH_PAUZE_TIME) + " seconds to prevent API locks by Google... (you loosing new Event inserts)");
      Logger.log("  progress:  0%");
      for (i=0;i<CONFIG.BATCH_PAUZE_INCREMENTS;i++){
        for (j=0;j<CONFIG.BATCH_PAUZE_TIME;j++){
          Utilities.sleep(1000);
        }
        Logger.log("  progress: " + Math.round((i+1)/CONFIG.BATCH_PAUZE_INCREMENTS*100)+ "%"); 
      }
    } // end if nextStartDate
  }
  logLine(doc, "Calendar events done copying, system ready");
}

// Fetch all events in the period startDate - endDate, limited by the max number to fetch.
// For each event from the sourceCal, duplicate this event in the targetCal.
function listNextEvents(doc, startDate, endDate, targetCal, loop) {
  const sourceCalId = CONFIG.SOURCE_CALENDAR_NAME;
  const targetCalId = CONFIG.TARGET_CALENDAR_NAME;
  const timeMinDate = Utilities.formatDate(new Date(startDate), 'Europe/Amsterdam', 'YYYY-MM-dd\'T\'HH:mm:ssZ');
  const timeMaxDate = Utilities.formatDate(new Date(endDate), 'Europe/Amsterdam', 'YYYY-MM-dd\'T\'HH:mm:ssZ');
  logLine(doc, "Fetching max " + CONFIG.FETCH_NUMBER_OF_EVENTS + " events from: " + timeMinDate + " to: "+timeMaxDate);
  
  const events = Calendar.Events.list(sourceCalId, {
    timeMin: timeMinDate,
    timeMax: timeMaxDate,
    singleEvents: true,
    orderBy: "startTime",
    maxResults: CONFIG.FETCH_NUMBER_OF_EVENTS,
  });
  if (!events.items || events.items.length === 0) {
    Logger.log("No events found.");
    return;
  } 
  
  var lastId="";     // contains the Id of the last event processed in the loop
  var start = null;  // contains the startTime of the last event processed in the loop
  var counter = 1;   // count the items in the loop
  for (const event of events.items) {
    lastId = event.id;
    const title = event.summary;
    var startDate;
    if (event.start.date){
      startDate = Utilities.formatDate(new Date(event.start.date), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm');
    } else { 
      startDate = Utilities.formatDate(new Date(event.start.dateTime), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm');
    }
    try{
      event.creator.email=CONFIG.TARGET_CALENDAR_NAME;
    } catch (e) {
      logLine(doc, "Warming:  "+  title + " (" + startDate + "): error setting event.creator.email");
    }
    // do not carry the event ID properties from the source Calendar into the target Calendar
    delete event.id;
    delete event.etag;
    delete event.iCalUID;
    delete event.htmlLink;

    
    switch (event.eventType){
      case "birthday":
          start = new Date(event.start.date);
          Logger.log(loop+"/"+counter + ") Birthday: %s (%s)", event.summary, Utilities.formatDate(new Date(start), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm') );
          createBirthday(doc,event, targetCal);
        break;

      case "focusTime":
        start = new Date(event.start.dateTime);
          Logger.log(loop+"/"+counter + ") Focus Time: %s (%s)", event.summary, Utilities.formatDate(new Date(start), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm') );
          createEvent(doc, event, targetCal);
        break;

      //case "fromGmail":
      //  break;

      case "outOfOffice": 
        event.reminders.useDefault='false';
        if (event.start.date) {
          // All-day event.
          start = new Date(event.start.date);
          Logger.log(loop+"/"+counter + ") Out of Office all day: %s (%s)", event.summary, Utilities.formatDate(new Date(start), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm') );
          createEvent(doc, event, targetCal);
        } else {
          start = new Date(event.start.dateTime);
          Logger.log(loop+"/"+counter + ") Out of Office: %s (%s)", event.summary, Utilities.formatDate(new Date(start), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm') );
          createEvent(doc, event, targetCal);
        }
        break;

      case "workingLocation": 
        Logger.log(loop+"/"+counter + ") Skipped: working location...");
        break; 

      default:
        if (event.start.date) {
          // All-day event.
          start = new Date(event.start.date);
          Logger.log(loop+"/"+counter + ") Yearly: %s (%s)", event.summary, start.toLocaleDateString());
          // @TODO add the substring below into the CONFIG
          if  (event.summary.includes("Jubileum") || event.summary.includes("Verjaardag") ) {
            createYearly(doc, event, targetCal);
          } else {
            Logger.log(loop+"/"+counter + ") DEFAULT all day: %s (%s)", event.summary, Utilities.formatDate(new Date(start), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm') );
            createEvent(doc, event, targetCal);
          }
        } else {
          start = new Date(event.start.dateTime);
          Logger.log(loop+"/"+counter + ") DEFAULT: %s (%s)", event.summary, Utilities.formatDate(new Date(start), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm') );
          createEvent(doc, event, targetCal);
        }
      break;
    } // end switch
    //Logger.log("Micro sleep...");
    //Utilities.sleep(CONFIG.MICRO_DELAY);
    //Logger.log("Done.");
    //doc.getBody().fl
    counter++;
  } // end for event
  var returnObject = new Object;
  returnObject.lastStart = start;
  returnObject.lastId = lastId; 
  return returnObject;
}


function createYearly(doc, sourceEvent, targetCal) {
  const eventJson = {
    start: { date: sourceEvent.start.date }, //toISOString()
    end: { date: sourceEvent.end.date },
    eventType: 'default',
    recurrence: ["RRULE:FREQ=YEARLY"],
    summary: sourceEvent.summary,
    transparency: "transparent",
    visibility: "private",
    reminders: {
      useDefault: false,
      overrides: [{
      method: 'popup',
      minutes: 720 }]  
    }
  }
  createEvent(doc, eventJson, targetCal);
}

function createBirthday(doc, event, targetCal) {
  const eventJson = {
    start: { date: event.start.date },
    end: { date: event.end.date },
    eventType: CalendarApp.EventType.BIRTHDAY,
    recurrence: ["RRULE:FREQ=YEARLY"],
    summary: event.summary,
    birthdayProperties: { type: event.birthdayProperties.type }, 
    transparency: "transparent",
    visibility: "private"
  }
  createEvent(doc, eventJson, targetCal);
}

function createEvent(doc, event, targetCal) {
  const title = event.summary;
  var startTime;
  if (event.start.date){
      startTime = Utilities.formatDate(new Date(event.start.date), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm');
    } else { 
      startTime = Utilities.formatDate(new Date(event.start.dateTime), 'Europe/Amsterdam', 'yyyy-MM-dd HH:mm');
    }
  if (!isDuplicate(event, targetCal)){
    if (!CONFIG.IS_DRY_RUN){
      try{ // try with recurring data on-board
        Calendar.Events.insert( event, targetCal.getId(), CONFIG.DEFAULT_EVENT_SETTINGS);
        //Logger.log("Insert Succes!");
      } catch (e) {
        if ((new String(e)).includes("API call to calendar.events.insert failed with error")){
          try{ // recurring event appears blocking, remove recurring shizzle
            //Logger.log("VOOR: "+ event);
            delete event.recurringEventId;
            //delete event.recurrence;
            delete event.recurringEvent;
            delete event.sequence;
            //Logger.log("NA: "+ event);
            Logger.log("REBOUND due to recurringEvent!");
            Calendar.Events.insert( event, targetCal.getId(), CONFIG.DEFAULT_EVENT_SETTINGS);
            logLine(doc, "[WARNING] Copied: " + title + " (" + startTime + ") without recurring event settings");
          } catch (err) {
            logLine(doc, "[ERROR] Failed to copy: " + title + " (" + startTime + "): " + err.message);
            Logger.log("CreateEvent: " + JSON.stringify(event) );
        }
        if ((new String(e)).includes("Calendar usage limits exceeded")){
          //WTF, Google reduces use of the API, let's terminate this run...
          logLine(doc, "[FATAL] Google says: \'Calendar usage limits exceeded\'. Failing event has timestamp: " + startTime + ". Use this as starting point for your next run");
          logLine(doc, "Script will terminate.");
          throw new Error("Calendar usage limits exceeded");
        }

      }  else {
          logLine(doc, "[ERROR] Failed to copy in rebound: " + title + " (" + startTime + "): " + e.message );
          logLine(doc, "Event JSON: " + event );
        }// end if
      }
    }
  } else {
    Logger.log("Duplicate: " + event.summary);
  }
}

function isDuplicate(event, targetCal){
  // DUPLICATE CHECK
  var existingEvents;
  if (event.start.date) { // all-day event
    existingEvents = targetCal.getEvents(new Date(event.start.date), new Date(event.end.date));
  } else { // regular event with time
    existingEvents = targetCal.getEvents(new Date(event.start.dateTime), new Date(event.end.dateTime));
  }
  return existingEvents.some(e => e.getTitle() === event.summary);
}

/**********************************/
/* Adjust the dates manually!     */
/* #HANDLE WITH CARE#             */
/**********************************/
function deleteAllTargetEvents(){
  const timeMinDate = new Date('2023-11-01').toISOString();
  const timeMaxDate = new Date('2023-11-01').toISOString();
 
  const events = Calendar.Events.list(CONFIG.TARGET_CALENDAR_NAME, {
    timeMin: timeMinDate,
    timeMax: timeMaxDate,
    singleEvents: true,
    orderBy: "startTime",
    maxResults: 100,
  });
  if (!events.items || events.items.length === 0) {
    Logger.log("No events found to delete.");
    return;
  } else {
    Logger.log("Found " + events.items.length + " calendar events to delete");
  }

  for (const event of events.items) {
    if (!CONFIG.IS_DRY_RUN){
      Logger.log("Deleting from "+ CONFIG.TARGET_CALENDAR_NAME +": %s (%s)", event.summary, new Date(event.start.dateTime).toLocaleDateString());
      Calendar.Events.remove(CONFIG.TARGET_CALENDAR_NAME, event.id, {sendUpdates: 'none'});
    } else {
      Logger.log("Dry run, otherwise would have deleted: %s (%s)", event.summary, new Date(event.start.dateTime).toLocaleDateString());
    }
  }
}
