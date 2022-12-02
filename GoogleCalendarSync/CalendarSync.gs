// Code inspired by https://janelloi.com/auto-sync-google-calendar/
// which uses https://gist.github.com/ttrahan/a88febc0538315b05346f4e3b35997f2 

var id="XXX"; // CHANGE - id of the secondary calendar to pull events from
const daysToSync = 70; // how many days out from today should the function look for events
var travelBuffer = 30; // in minutes; make this a negative number to disable travel buffer events
var includeTitle = true; // whether the holds and travel buffers include the personal event title in the description; if false, it will use the event id from the secondary calendar


/*----- KNOWN (POTENTIAL) ISSUES -----
* 1. The amount of time for the travel buffer events is fixed (currently set by the variable in line 6).
*     If you change the travel time value, the script will not change existing travel buffers.
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
// (which simplifies updating holds and buffers), but that means there would be an
// additional file stored in Google Drive


/*----- CUSTOMIZEABLE OPTIONS -----*/
// primary calendar = work calendar
// secondary calendar = personal calendar, which is the source of events to sync

var primaryEventTitle = "Busy (Synced Event)"; // update this to the text you'd like to appear in the new events created in primary calendar
var primaryEventTravelTitle = "Travel buffer"; // if a synced event has a location set, this will be the name for a travel buffer event
var primaryEventTravelDescription = ""; // if a synced event has a location set, this be the start of the description for the travel buffer event
var primaryEventLocation = "[Travel required]"; // to sync a location (that isn't a Zoom URL), but hide the details

var syncWeekdaysOnly = true; // only sync events on weekdays (i.e., skip weekends)
var skipAllDayEvents = true; // only sync events with scheduled times (i.e., skip events that are all day or multi-day)

var removeEventReminders = false; // if true, this will disable reminders on the primary calendar for all hold events
var removeTravelBeforeReminders = false; // if true, this will disable reminders for travel buffers BEFORE synced events
var removeTravelAfterReminders = true; // if true, this will disable reminders for travel buffers AFTER synced events
var removeEventReminderIfTravelReminder = true; // if true, this will disable reminders for hold events IF removeTravelBeforeReminders is set to true and travel does happen

var descriptionStart = ''; // if you want certain text in every primary calendar hold description, e.g., "This is synced from [email]"
var includeDescription = false; // whether holds in the primary calendar should include the description from the secondary calendar


/*----- OPTIONS YOU PROBABLY SHOULDN'T CHANGE -----*/

// DO NOT CHANGE THIS UNLESS YOU REALLY WANT TO TRY TO DELETE 
// ALL OF THE HOLDS AND TRAVEL BUFFERS ON YOUR PRIMARY CALENDAR
// note - this will fail if you changed things like the primaryEventTitle,
// primaryEventTravelTitle, and/or primaryEventTravelDescription
var clearAllHolds = false; // will try to delete all holds, travel buffers on the primary calendar

// DO NOT CHANGE THIS PARAMTER - required for creating travel buffers
const minsToMilliseconds = 60000; // convert minutes to milliseconds


/*----- CUSTOMIZEABLE HELPER FUNCTIONS -----*/
// check if an event in the primary calendar is an existing hold for
// and event in the secondary calendar 
function compareEvents(primaryEvent, secondaryEvent) {
  var primaryStart = primaryEvent.getStartTime().getTime();
  var primaryEnd = primaryEvent.getEndTime().getTime();
  var secondaryStart = secondaryEvent.getStartTime().getTime();
  var secondaryEnd = secondaryEvent.getEndTime().getTime();
  var secondaryDescription = createDescription(secondaryEvent);
  if ((primaryStart === secondaryStart) && (primaryEnd === secondaryEnd) && (primaryEvent.getDescription() == secondaryDescription)) {
    return true;
  }
  return false;
}

// check whether an event is within the days that are being synced
// uses the syncWeekdaysOnly variable above
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
  str = str + getEventTitle(evi);

  // add in the event description if requested
  if (includeDescription) {
    if (str.length > 0) {
      str = str + '\n\n'
    }
    str = str + evi.getDescription();
  }
  return str;
}

// function for creating events
// options can either be the event from the secondary calendar
// in which case the function will try to create the description and location from it
// but if that fails, then it assumes options is an object array that includes that info already (or null)
// removeFlag = boolean for whether to remove reminders
function eventCreate(cal, title, startTime, endTime, opt_event) {
  try {
    var description = createDescription(opt_event);
    var location = createLocation(opt_event);
    newEvent = cal.createEvent(title, startTime, endTime, {description: description, location: location})
  } catch (e) {
    newEvent = cal.createEvent(title, startTime, endTime, opt_event);
  }

  newEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
  
  
  if (checkEventLocation(newEvent)) {
    var newTravel = eventTravel(cal, getEventTitleDescription(newEvent), startTime, endTime);
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

  return {event: newEvent, travel: newTravel}; 
}

// function for creating travel buffers
// assumes opt_event is an object (or null) with a description
function eventTravelCreate(cal, title, startTime, endTime, removeFlag, opt_event) {
  newEventTravel = cal.createEvent(title, startTime, endTime, opt_event);
  newEventTravel.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
  if (removeFlag) {
    newEventTravel.removeAllReminders();
  } 
  return newEventTravel; 
}

// function for updating events with location, description, and ensuring privacy
// pEvent = hold that's being put in the primary calendar
// secEvent = original event in the secondary/personal calendar
// removeFlag = boolean for whether to remove reminders
function eventUpdate(cal, eventsTravel, pEvent, secEvent) {
  pEvent.setDescription(createDescription(secEvent));
  pEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
  //location - check if it's changed
  var pLocation = pEvent.getLocation().trim();
  var secLocation = createLocation(secEvent);
  
  if ((pLocation == '') && (secLocation != '') && (travelBuffer > 0)) { //new location added
    var pTravel = eventTravel(cal, getEventTitleDescription(pEvent), secEvent.getStartTime(), secEvent.getEndTime());
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
    var pTravelDeleted = travelDelete(pEvent,getEventTitleDescription(pEvent), eventsTravel);
    pEvent.setLocation(secLocation);
    // check if need to restore notifications
    if (removeEventReminderIfTravelReminder && !removeTravelBeforeReminders) {
      pEvent.resetRemindersToDefault();
    }
  }
  
  return {event: pEvent, newTravel: pTravel, deleteTravel: pTravelDeleted};
}

// create travel buffer events
function eventTravel(cal, eviTitle, eviStartTime, eviEndTime) {
  var travelStartTime = new Date(eviStartTime.getTime() - travelBuffer*minsToMilliseconds);
  var travelEndTime = new Date(eviEndTime.getTime() + travelBuffer*minsToMilliseconds);

  var travelBefore = eventTravelCreate(cal, primaryEventTravelTitle, travelStartTime, eviStartTime, 
                                        removeTravelBeforeReminders, {description: primaryEventTravelDescription+eviTitle});
  Logger.log('TRAVEL BUFFER CREATED (BEFORE)' + 
                '\ntravelId: ' + travelBefore.getId() + 
                '\ntravelTitle: ' + travelBefore.getTitle() + 
                '\ntravelDesc ' + travelBefore.getDescription() + '\n');

  var travelAfter = eventTravelCreate(cal, primaryEventTravelTitle, eviEndTime, travelEndTime, 
                                        removeTravelAfterReminders, {description: primaryEventTravelDescription+eviTitle});
  Logger.log('TRAVEL BUFFER CREATED (AFTER)' + 
                '\ntravelId: ' + travelAfter.getId() + 
                '\ntravelTitle: ' + travelAfter.getTitle() + 
                '\ntravelDesc ' + travelAfter.getDescription() + '\n');

  // return {travelBefore: travelBefore.getId(), travelAfter: travelAfter.getId()};
  return [travelBefore.getId(), travelAfter.getId()];
}


// function to delete an event
function eventDelete(evi, deleteStr) {
  var eviId = evi.getId();
  Logger.log(deleteStr + 
            '\nprimaryId: ' + eviId + '\nprimaryTitle: ' + evi.getTitle() + '\nprimaryDesc ' + evi.getDescription() + '\n');
  evi.deleteEvent();
  return eviId;
}

// function to find and delete travel
function travelDelete(pEvent, eventTravelTitle, eventsTravel) {
  // Logger.log(eventTravelTitle);
  pStartTime = pEvent.getStartTime().getTime();
  pEndTime = pEvent.getEndTime().getTime();
  var pTravelDeleted = [];
  for (pev2 in eventsTravel) {
    var pTravel = eventsTravel[pev2];
    if (pTravel.getDescription() == primaryEventTravelDescription+eventTravelTitle) {
      if ((pTravel.getEndTime().getTime() == pStartTime) || (pTravel.getStartTime().getTime() == pEndTime)) {
        pTravelDeleted.push(eventDelete(pTravel, 'TRAVEL BUFFER DELETED'));
      }
    }
  }
  return pTravelDeleted;
}

// get the original secondary/personal event title from the hold in the primary/work calendar
function getEventTitleDescription(pEvent) {
  var desc = pEvent.getDescription().replace(descriptionStart, '');
  return desc.split(/\r?\n/)[0];
}
// get the unique identifier (title or id) from the secondary/personal event
function getEventTitle(evi) {
  if (includeTitle) {
    return evi.getTitle();
  }
  return evi.getId().replace('@google.com', '');
}


/*----- NOW THE MAIN FUNCTION -----*/
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
  
  var stat=1;
  var evi, existingEvent; 
  var primaryEventsFiltered = []; // to contain primary calendar events that were previously created from secondary calendar
  var primaryEventsTravel = [];  // to contain primary calendar events that are travel buffers for the secondary calendar
  var primaryEventsUpdated = []; // to contain primary calendar events that were updated from secondary calendar
  var primaryEventsCreated = []; // to contain primary calendar events that were created from secondary calendar
  var primaryEventsTravelCreated = []; // to contain primary calendar travel buffer events that were created
  var primaryEventsDeleted = []; // to contain primary calendar events previously created that have been deleted from secondary calendar
  var primaryEventsTravelDeleted = []; // to contain primary calendar travel buffer events that were deleted

  Logger.log('Number of primaryEvents: ' + primaryEvents.length);  
  Logger.log('Number of secondaryEvents: ' + secondaryEvents.length);
  
  // create filtered list of existing primary calendar events that were previously created from the secondary calendar
  for (pev in primaryEvents)
  {
    var pEvent = primaryEvents[pev];
    var pTitle = pEvent.getTitle();

    if (pTitle === primaryEventTitle) { 
      // Logger.log('found event: '+ pEvent.getId());
      primaryEventsFiltered.push(pEvent); 
    }
    if (pTitle === primaryEventTravelTitle) { 
      // Logger.log('found travel: '+ pEvent.getId());
      primaryEventsTravel.push(pEvent); 
    }
  }
  
  // process all events in secondary calendar
  for (sev in secondaryEvents)
  {
    stat=1;
    evi=secondaryEvents[sev];

    // skip events you RSVP'ed "no" to
    if (evi.getMyStatus()=="NO") {
      continue;
    }

    // Do nothing if the event is an all-day or multi-day event. This script only syncs hour-based events
    if (skipAllDayEvents && evi.isAllDayEvent())
    {
      continue; 
    }
    
    // if the secondary event has already been blocked in the primary calendar, update it
    for (existingEvent in primaryEventsFiltered)
      {
        var pEvent = primaryEventsFiltered[existingEvent];
        if (compareEvents(pEvent, evi)) {
          stat=0;
          var pEventTravel = eventUpdate(primaryCal, primaryEventsTravel, pEvent, evi);
          var pEvent = pEventTravel.event;
          primaryEventsUpdated.push(pEvent.getId());
          if (pEventTravel.newTravel != null) { primaryEventsTravelCreated = primaryEventsTravelCreated.concat(pEventTravel.newTravel); }
          if (pEventTravel.deleteTravel != null) { primaryEventsTravelDeleted = primaryEventsTravelDeleted.concat(pEventTravel.deleteTravel); }
          Logger.log('PRIMARY EVENT UPDATED'
                     + '\nprimaryId: ' + pEvent.getId() + ' \nprimaryTitle: ' + pEvent.getTitle() + ' \nprimaryDesc: ' + pEvent.getDescription());
        } 
      }

    if (stat==0) {
      continue;    
    }

    // if the secondary event does not exist in the primary calendar, create it
    // as long as it's on a day that we're syncing events for
    if (checkEventDay(evi)) {
      var eviStartTime = evi.getStartTime();
      var eviEndTime = evi.getEndTime();
      var newEventTravel = eventCreate(primaryCal, primaryEventTitle, eviStartTime, eviEndTime, evi);
      var newEvent = newEventTravel.event;
      if (newEventTravel.travel != null) { primaryEventsTravelCreated = primaryEventsTravelCreated.concat(newEventTravel.travel); }
      primaryEventsCreated.push(newEvent.getId());
      Logger.log('PRIMARY EVENT CREATED'
                 + '\nprimaryId: ' + newEvent.getId() + '\nprimaryTitle: ' + newEvent.getTitle() + '\nprimaryDesc ' + newEvent.getDescription() + '\n');
    }
  }

  // if a primary event previously created no longer exists in the secondary calendar, delete it
  for (pev in primaryEventsFiltered)
  {
    var pEvent = primaryEventsFiltered[pev];
    var pevIsUpdatedIndex = primaryEventsUpdated.indexOf(pEvent.getId());
    if (pevIsUpdatedIndex == -1) // it's not among the updated events
    { 
      if (checkEventLocation(pEvent)) { // if it had a travel buffer, find those too
        var deletedTravel = travelDelete(pEvent, getEventTitleDescription(pEvent), primaryEventsTravel);
        primaryEventsTravelDeleted = primaryEventsTravelDeleted.concat(deletedTravel);
      }
      primaryEventsDeleted.push(eventDelete(pEvent, 'EVENT DELETED'));
    }
  }  

  Logger.log('Primary events previously created: ' + primaryEventsFiltered.length);
  Logger.log('Primary events updated: ' + primaryEventsUpdated.length);
  Logger.log('Primary events deleted: ' + primaryEventsDeleted.length);
  Logger.log('Primary events created: ' + primaryEventsCreated.length);
  Logger.log('Primary travel buffers deleted: ' + primaryEventsTravelDeleted.length);
  Logger.log('Primary travel buffers created: ' + primaryEventsTravelCreated.length);
}  
