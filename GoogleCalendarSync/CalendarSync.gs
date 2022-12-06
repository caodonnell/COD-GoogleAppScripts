// Code inspired by https://janelloi.com/auto-sync-google-calendar/
// which uses https://gist.github.com/ttrahan/a88febc0538315b05346f4e3b35997f2 

var id="christine.a.odon@gmail.com"; // CHANGE - id of the secondary calendar to pull events from
const daysToSync = 70; // how many days out from today should the function look for events
var travelBuffer = 30; // in minutes; make this a negative number to disable travel buffer events
var includeTitle = true; // whether the holds and travel buffers include the personal event title in the description; if false, it will use the event id from the secondary calendar


/*----- KNOWN (POTENTIAL) ISSUES -----
* 1. The amount of time for the travel buffer events is fixed (currently set by the variable in line 31).
* 2. The script requires that the description of the holds and travel buffers (if applicable) have a unique 
*     name/identifier. By default, this is the title of the original event in the personal/secondary 
*     calendar, but if includeTitle (line 50) is false, it will use the event id instead,
* 3. The script checks for Zoom URLs in the location field by looking for '.zoom.'. If that's present, it 
*     will ignore the location field in the personal/secondary calendar.
* 4. All events on the secondary calendar create holds on the primary calendar.
*     So, if there are 2 events at the same time, you will get 2 holds at the same time.
* 5. When checking whether something should be updated on the primary calendar 
*     (e.g., because it was deleted in the secondary calendar), it checks the event time 
*     and description(for the original personal/secondary calendar event time and title).
*     This means if there are 2 hold/events at the same time, and you delete one of those events 
*     from your personal calendar, the script will try to delete the corresponding hold 
*     and its travel time (if present). However, this is the most tenuous part of the process.
* 6. If an event is removed from the secondary/personal calendar, the script will try to remove
*     it (and travel buffers, if present) from the primary/work calendar. However, if somehow there
*     are orphan travel buffers in the primary/work calendar without a hold event (e.g., you deleted
*     the hold event manually), the script will not remove the orphan travle buffers.
*/

// NOTE: This script only uses data from the two calendars. 
// A more efficient script would create a database to store info on created events
// (which simplifies updating holds and buffers), but that means there would be 
// additional file(s) stored in Google Drive


/*----- CUSTOMIZEABLE OPTIONS -----*/
// some definitions...
// primary calendar = work calendar
// secondary calendar = personal calendar, which is the source of events to sync

var primaryEventTitle = "Busy (Synced Event)"; // update this to the text you'd like to appear in the new events created in primary calendar
var primaryEventTravelTitle = "Travel buffer"; // if a synced event has a location set, this will be the name for a travel buffer event
var primaryEventTravelDescription = ""; // if a synced event has a location set, this be the start of the description for the travel buffer event
var primaryEventLocation = "[Travel required]"; // to sync a location (that isn't a Zoom URL), but hide the details

var syncWeekdaysOnly = true; // only sync events on weekdays (i.e., skip weekends)
var skipAllDayEvents = true; // only sync events with scheduled times (i.e., skip events that are all day or multi-day)

// if out of office on primary/work calendar during a personal event, don't sync
// NOTE: REQUIRES USING the "Out of office" event type in Google Calendar
var checkOutOfOffice = true; 

var removeEventReminders = false; // if true, this will disable reminders on the primary calendar for all hold events
var removeTravelBeforeReminders = false; // if true, this will disable reminders for travel buffers BEFORE synced events
var removeTravelAfterReminders = true; // if true, this will disable reminders for travel buffers AFTER synced events
var removeEventReminderIfTravelReminder = true; // if true, this will disable reminders for hold events IF removeTravelBeforeReminders is set to true and travel does happen

// these two booleans determine whether holds and travel buffers use the default primary calendar color
var useDefaultEventColor = true;
var useDefaultTravelColor = true;
// if the above are false, the following colors will be used
// color names: https://developers.google.com/apps-script/reference/calendar/event-color
var customEventColor = 'GRAY';
var customTravelColor = 'PALE_RED';

var descriptionStart = ''; // if you want certain text in every primary calendar hold description, e.g., "This is synced from [email]"
var includeDescription = false; // whether holds in the primary calendar should include the description from the secondary calendar

// to prevent multiple runs at the same time (from updating multiple events on the personal calendar)
// the script uses LockServices to hold additional runs 
// this should be several times the estimated maximum time of a single run
var maxLockWaitTime = 600; // in seconds

/*----- OPTIONS YOU PROBABLY SHOULDN'T CHANGE -----*/

// DO NOT CHANGE THIS UNLESS YOU REALLY WANT TO TRY TO DELETE 
// ALL OF THE HOLDS AND TRAVEL BUFFERS ON YOUR PRIMARY CALENDAR
// note - this will fail if you changed things like the primaryEventTitle,
// primaryEventTravelTitle, and/or primaryEventTravelDescription
var clearAllHolds = false; // will try to delete all holds, travel buffers on the primary calendar

// DO NOT CHANGE THIS PARAMTER - required for creating travel buffers
const minsToMilliseconds = 60000; // convert minutes to milliseconds

// BE CAREFUL ABOUT CHANGING
// custom metadata in the primary calendar to link events to the secondary calendar
// for travel buffers, the tags need to be unique for before vs. after travel
var tagKey = 'syncID';
var travelBeforeTag = ' - before';
var travelAfterTag = ' - after';


/*----- MAIN FUNCTION -----*/
// calls the sync function, wrapped with LockService to avoid multiple 
// parallel iterations (which will cause errors, e.g., if something is
// deleted from the personal calendar during the run, trigger a second execution)
function mainFcn() {
  var lock = LockService.getScriptLock();
  lock.tryLock(maxLockWaitTime * 1000); // convert to milliseconds
  if (lock.hasLock()) {
    sync();
    lock.releaseLock();
  } else {
    Logger.log('Could not obtain lock after 60 seconds.');
    Error('Lock service failed. Maybe another execution is still running?');
  }
  lock.releaseLock();
}


/*----- HELPER FUNCTIONS FOR SYNC STATUS, DESCRIPTIONS, LOCATIONS, AND TITLES -----*/

// check whether an event is within the days that are being synced
// uses the syncWeekdaysOnly variable above
// true: event is on a weekday, or syncWeekdaysOnly not enabled
function checkEventDay(evi) {
  if (syncWeekdaysOnly) {
    var startDay = evi.getStartTime().getDay();
    var endDay = evi.getEndTime().getDay();
    // Logger.log(evi.getTitle()+': '+day)
    if (((startDay < 1) || (startDay > 5)) && ((endDay < 1) || (endDay > 5))) { // not a weekday
      return false;
    } 
  }
  // Logger.log(evi.getTitle()+': '+day)
  return true;
}

// check if marked out of office on the primary/work calendar
// if so, then don't need to sync the event
// true: primary/work calendar is OOO
function checkOOOstatus(evi) {
  if (checkOutOfOffice) {
    var resource = {
      timeMin: Utilities.formatDate(evi.getStartTime(), 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
      timeMax: Utilities.formatDate(evi.getEndTime(), 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
      // q: "Out of office"
    };
    var outOffice = Calendar.Events.list('primary', resource).items.filter(x => x.eventType == 'outOfOffice');
    // Logger.log(outOffice);
    if (outOffice.length > 0) {
      return true;
    }
  }
  return false;
}

// if there is a location in the secondary calendar event
// (and it's not a Zoom link)
// give the primary calendar hold the primaryEventLocation text
function createLocation(evi) {
  var eviLocation = evi.getLocation().trim(); // remove whitespace
  if ((eviLocation != '') && (eviLocation.indexOf('.zoom.') < 0)) {
    return primaryEventLocation;
  }
  return '';
}
//similar to above, but returns true if there is a non-Zoom location
// also checks if travelBufer is positive (i.e., whether there will be travel buffers)
function checkEventLocation(evi) {
  var eviLocation = evi.getLocation().trim(); // remove whitespace
  if ((eviLocation != '') && (eviLocation.indexOf('.zoom.') < 0) && (travelBuffer > 0)) {
    return true;
  }
  return false;
}

// set a primary event/hold description 
function createDescription(evi) {
  var str = descriptionStart;

  // add in the title, but first some newlines if there was a starting string
  if (str.length > 0) {
    str = str + '\n\n'
  }
  str = str + getEventTitlePrivate(evi);

  // add in the event description if requested
  if (includeDescription) {
    if (str.length > 0) {
      str = str + '\n\n'
    }
    str = str + evi.getDescription();
  }
  return str;
}

// get the original secondary/personal event title from the hold in the primary/work calendar
function getEventTitleDescription(pEvent) {
  var desc = pEvent.getDescription().replace(descriptionStart, '');
  return desc.split(/\r?\n/)[0];
}
// get the unique identifier (title or id) from the secondary/personal event
function getEventTitlePrivate(evi) {
  if (includeTitle) {
    return evi.getTitle();
  }
  return evi.getId().replace('@google.com', '');
}

// a subtlety here is how event series are handled
// if something is a recurring event, all instances have the same ID, so
// it won't have a unique ID in the personal/secondary calendar.
// this means if we want to have a unique ID for an event, 
// we also need to include the day it starts on
function createSyncTag(evi) {
  return evi.getId()+' - '+evi.getStartTime()+' - '+evi.getEndTime();
}

/*----- HELPER FUNCTIONS FOR CREATING, UPDATING, AND DELETING HOLDS -----*/

// function for creating events
// options can either be the event from the secondary calendar
// in which case the function will try to create the description and location from it
// but if that fails, then it assumes options is an object array that includes that info already (or null)
// removeFlag = boolean for whether to remove reminders
function eventCreate(cal, startTime, endTime, tagVal, evi) {

  var description = createDescription(evi);
  var location = createLocation(evi);
  // var tagVal = createSyncTag(evi);

  newEvent = cal.createEvent(primaryEventTitle, startTime, endTime, 
                              {description: description, location: location})
  newEvent.setTag(tagKey, tagVal);
  newEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
  
  
  if (checkEventLocation(newEvent)) {
    var newTravel = eventTravel(cal, getEventTitleDescription(newEvent), tagVal, startTime, endTime);
  } else {
    var newTravel = null;
  }
  
  if (removeEventReminders) {
    // Logger.log(newEvent.getDescription()+' - flag to remove notifications');
    newEvent.removeAllReminders();
  } else if (removeEventReminderIfTravelReminder && (newTravel != null) && !removeTravelBeforeReminders) {
    // Logger.log(newEvent.getDescription()+' - because of travel, removing notifications');
    newEvent.removeAllReminders();
  } //else {
    // Logger.log(newEvent.getDescription()+' - keeping notifications');
  //}

  //set the color if needed
  if (!useDefaultEventColor) {
    newEvent.setColor(customEventColor);
  }

  Logger.log('PRIMARY EVENT CREATED' + 
              '\nprimaryId: ' + newEvent.getId() + 
              '\nprimaryDesc ' + newEvent.getDescription() + 
              '\nprimaryTag '+newEvent.getTag(tagKey) + '\n');
  return {event: newEvent, travel: newTravel}; 
}

// function for creating travel buffers
// assumes opt_event is an object (or null) with a description
function eventTravelCreate(cal, startTime, endTime, eviTitle, eviTag, removeFlag) {
  newEventTravel = cal.createEvent(primaryEventTravelTitle, startTime, endTime,
                                    {description: primaryEventTravelDescription+eviTitle});
  newEventTravel.setTag(tagKey, eviTag)
  newEventTravel.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
  if (removeFlag) {
    newEventTravel.removeAllReminders();
  } 
  if (!useDefaultTravelColor) {
    newEventTravel.setColor(customTravelColor);
  }
  return newEventTravel; 
}

// create travel buffer events
function eventTravel(cal, eviTitle, eviTag, eviStartTime, eviEndTime) {
  var travelStartTime = new Date(eviStartTime.getTime() - travelBuffer*minsToMilliseconds);
  var travelEndTime = new Date(eviEndTime.getTime() + travelBuffer*minsToMilliseconds);

  var travelBefore = eventTravelCreate(cal, travelStartTime, eviStartTime, 
                                        eviTitle, eviTag+travelBeforeTag,
                                        removeTravelBeforeReminders);
  Logger.log('TRAVEL BUFFER CREATED (BEFORE)' + 
                '\ntravelId: ' + travelBefore.getId() + 
                '\ntravelDesc ' + travelBefore.getDescription() + 
                '\ntravelTag ' + travelBefore.getTag(tagKey) + '\n');

  var travelAfter = eventTravelCreate(cal, eviEndTime, travelEndTime, 
                                        eviTitle, eviTag+travelAfterTag,
                                        removeTravelAfterReminders);
  Logger.log('TRAVEL BUFFER CREATED (AFTER)' + 
                '\ntravelId: ' + travelAfter.getId() + 
                '\ntravelDesc ' + travelAfter.getDescription() +  
                '\ntravelTag ' + travelAfter.getTag(tagKey) + '\n');

  // return {travelBefore: travelBefore.getId(), travelAfter: travelAfter.getId()};
  return [travelBefore.getId(), travelAfter.getId()];
}

// function for updating events with location, description, and ensuring privacy
// pEvent = hold that's being put in the primary calendar
// secEvent = original event in the secondary/personal calendar
function eventUpdate(cal, eventsTravel, pEvent, secEvent) {

  // make sure the description is accurate
  var newDescription = createDescription(secEvent);
  pEvent.setDescription(newDescription);

  pEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar

  // make sure start and end times are accurate
  pEvent.setTime(secEvent.getStartTime(), secEvent.getEndTime());

  //location - check if it's changed
  var pLocation = pEvent.getLocation().trim();
  var secLocation = createLocation(secEvent); 

  if ((pLocation == '') && (secLocation != '') && (travelBuffer > 0)) { //new location added
    var pTravel = eventTravel(cal, getEventTitleDescription(pEvent), pEvent.getTag(tagKey),
                                secEvent.getStartTime(), secEvent.getEndTime());
    var pTravelDeleted = null;
    pEvent.setLocation(secLocation);
  } 

  if (removeEventReminders) {
    pEvent.removeAllReminders();
  } else if (removeEventReminderIfTravelReminder && ((pTravel != null) || ((pLocation != '')&&(travelBuffer > 0))) && !removeTravelBeforeReminders) {
    pEvent.removeAllReminders();
  }
  
  if ((pLocation != '') && (secLocation == '') && (travelBuffer > 0)) { //location removed
    var pTravel = null;
    var pTravelDeleted = [travelDelete(eventsTravel[pEvent.getTag(tagKey)+travelBeforeTag]), 
                            travelDelete(eventsTravel[pEvent.getTag(tagKey)+travelAfterTag])]
    pEvent.setLocation(secLocation);
    // check if need to restore notifications
    if (removeEventReminderIfTravelReminder && !removeTravelBeforeReminders) {
      pEvent.resetRemindersToDefault();
    }
  }

  Logger.log('PRIMARY EVENT UPDATED' + 
              '\nprimaryId: ' + pEvent.getId() + 
              '\nprimaryDesc: ' + newDescription +  
              '\nprimaryTag ' + pEvent.getTag(tagKey) + '\n');
  return {event: pEvent, newTravel: pTravel, deleteTravel: pTravelDeleted};
}


// function to delete an event
function deleteFcn(evi, deleteStr) {
   try {
     var eviId = evi.getId();
    Logger.log(deleteStr + 
              '\nprimaryId: ' + eviId + 
              '\nprimaryDesc ' + evi.getDescription() + 
              '\nprimaryTag ' + evi.getTag(tagKey) + '\n');
    evi.deleteEvent();
    return eviId;
   } catch (e) {
     Logger.log('ERROR: COULD NOT DELETE:' +
                  '\nprimaryId: ' + evi.getId() + 
                  '\nprimaryDesc ' + evi.getDescription() + 
                  '\nprimaryTag ' + evi.getTag(tagKey) + '\n');
   }
   return;
}

// function to delete a hold on the primary calendar 
function eventDelete(evi) {
  return deleteFcn(evi, 'EVENT DELETED');
}

// function to delete a travel event
// because 2 travel buffers can be mapped to a single hold
// this needs its own wrapper function
function travelDelete(evi) {
  return deleteFcn(evi, 'TRAVEL BUFFER DELETED');
}


/*----- NOW THE ACTUAL FUNCTION -----*/

function sync() {
  var today=new Date();
  var enddate=new Date();
  enddate.setDate(today.getDate()+daysToSync); // how many days in advance to monitor and block off time
  
  var secondaryCal=CalendarApp.getCalendarById(id);
  var secondaryEvents=secondaryCal.getEvents(today,enddate);
  if (clearAllHolds) { // THIS WILL DELETE ALL HOLDS AND TRAVEL BUFFERS
    var secondaryEvents=[]; 
  }
  
  var primaryCal=CalendarApp.getDefaultCalendar();
  var primaryEvents=primaryCal.getEvents(today,enddate); // all primary calendar events
  
  var evi; 
  var primaryEventsFiltered = new Object(); // to contain primary calendar events that were previously created from secondary calendar
  var primaryEventsTravel = new Object();  // to contain primary calendar events that are travel buffers for the secondary calendar
  var primaryEventsUpdated = []; // to contain IDs of primary calendar events that were updated from secondary calendar
  var primaryEventsCreated = []; // to contain IDs of primary calendar events that were created from secondary calendar
  var primaryEventsTravelCreated = []; // to IDs of contain primary calendar travel buffer events that were created
  var primaryEventsDeleted = []; // to contain IDs of primary calendar events previously created that have been deleted from secondary calendar
  var primaryEventsTravelDeleted = []; // to IDs of contain primary calendar travel buffer events that were deleted
  
  // Logger.log('Number of primaryEvents: ' + primaryEvents.length);  
  Logger.log('Number of secondaryEvents: ' + secondaryEvents.length);
  
  // create filtered list of existing primary calendar events that were previously created from the secondary calendar
  for (pev in primaryEvents)
  {
    var pEvent = primaryEvents[pev];
    var pTitle = pEvent.getTitle();
    var pTag = pEvent.getTag(tagKey);

    if (pTitle === primaryEventTitle) { 
      // Logger.log('Hold event\n'+pTag+'\n'+pEvent.getStartTime()+pEvent.getEndTime());
      primaryEventsFiltered[pTag] = pEvent;
      // Logger.log(Object.keys(primaryEventsFiltered));
      // Logger.log(primaryEventsFiltered[pTag])
    }
    if (pTitle === primaryEventTravelTitle) { 
      // Logger.log('Travel event\n'+pTag+'\n'+pEvent.getStartTime()+pEvent.getEndTime());
      primaryEventsTravel[pTag] = pEvent;
      // Logger.log(Object.keys(primaryEventsTravel));
      // Logger.log(primaryEventsTravel[pTag]);
    }
  }
  
  Logger.log('Number of primaryEvents (holds): ' + Object.keys(primaryEventsFiltered).length);
  Logger.log('Number of primaryEvents (travel buffers): ' + Object.keys(primaryEventsTravel).length);

  // process all events in secondary calendar
  for (sev in secondaryEvents)
  {
    evi=secondaryEvents[sev];

    // skip events you RSVP'ed "no" to
    if (evi.getMyStatus()=="NO") {
      continue;
    }

    // Do nothing if the event is an all-day event, if that setting is enabled
    if (skipAllDayEvents && evi.isAllDayEvent())
    {
      continue; 
    }
    
    // if the secondary event is on a day/time that we're syncing events for, update/create the hold(s)
    if (checkEventDay(evi) && !checkOOOstatus(evi)) {
      eviTag = createSyncTag(evi);
      
      // try to update the existing event
      // if this fails, it'll assume it's a new event to sync
     if (eviTag in primaryEventsFiltered) {
        var pEvent = primaryEventsFiltered[eviTag];
        var pEventTravel = eventUpdate(primaryCal, primaryEventsTravel, pEvent, evi);
        primaryEventsUpdated.push(eviTag);
        if (pEventTravel.newTravel != null) { primaryEventsTravelCreated = primaryEventsTravelCreated.concat(pEventTravel.newTravel); }
        if (pEventTravel.deleteTravel != null) { primaryEventsTravelDeleted = primaryEventsTravelDeleted.concat(pEventTravel.deleteTravel); }
      } else {
        var eviStartTime = evi.getStartTime();
        var eviEndTime = evi.getEndTime();
        var newEventTravel = eventCreate(primaryCal, eviStartTime, eviEndTime, eviTag, evi);
        primaryEventsCreated.push(newEventTravel.event.getTag(tagKey));
        if (newEventTravel.travel != null) { primaryEventsTravelCreated = primaryEventsTravelCreated.concat(newEventTravel.travel); }   
      }
    }
  }

  // if a primary event previously created no longer exists in the secondary calendar, delete it
  var delEv, deleteEvents;
  deleteEvents = Object.keys(primaryEventsFiltered).filter(x => primaryEventsUpdated.indexOf(x) < 0);
  for (delEv in deleteEvents) {
    primaryEventsDeleted.push(eventDelete(primaryEventsFiltered[deleteEvents[delEv]]));
  }
  // repeat with the travel buffers
  deleteEvents = Object.keys(primaryEventsTravel).filter(x => primaryEventsUpdated.findIndex(y => x.indexOf(y) > -1) < 0);
  for (delEv in deleteEvents) {
    primaryEventsTravelDeleted = primaryEventsTravelDeleted.concat(travelDelete(primaryEventsTravel[deleteEvents[delEv]]));
  }


  // Logger.log('Primary events previously created: ' + primaryEventsFiltered.length);
  Logger.log('Primary events updated: ' + primaryEventsUpdated.length);
  Logger.log('Primary events deleted: ' + primaryEventsDeleted.length);
  Logger.log('Primary events created: ' + primaryEventsCreated.length);
  Logger.log('Primary travel buffers deleted: ' + primaryEventsTravelDeleted.length);
  Logger.log('Primary travel buffers created: ' + primaryEventsTravelCreated.length);

  return 0;
}  
