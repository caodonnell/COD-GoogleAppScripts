// Code inspired by https://janelloi.com/auto-sync-google-calendar/
// which uses https://gist.github.com/ttrahan/a88febc0538315b05346f4e3b35997f2 

var id="XXX"; // CHANGE - id of the secondary calendar to pull events from

/*----- KNOWN (POTENTIAL) ISSUES -----
* 1. Travel buffer time is fixed (currently set by the variable in line 24)
* 2. Skips all-day events
* 3. All events on the secondary calendar create holds on the primary calendar.
*     So, if there are 2 events at the same time, you will get 2 holds at the same time.
* 3. When checking whether something should be updated on the primary calendar 
*     (e.g., because it was deleted in the secondary calendar), it ONLY checks the event time.
*     This means if there are 2 things in your personal calendar at the same time, you will get 
*     2 holds on your primary, but if you delete one of those things, neither hold will be deleted.
*/

/*----- CUSTOMIZEABLE OPTIONS -----*/
// primary calendar = work calendar
// secondary calendar = personal calendar, which is the source of events to sync

var primaryEventTitle = "Busy (Synced Event)"; // update this to the text you'd like to appear in the new events created in primary calendar
var primaryEventTravelTitle = "Travel buffer"; // if a synced event has a location set, this will automatically add in travel time
var primaryEventLocation = "Travel required"; // to sync a location, but hide the details
var travelBuffer = 30; // in minutes
const minsToMilliseconds = 60000; // convert minutes to milliseconds

const daysToSync = 70; // how many days out from today should the function look for events

var syncWeekdaysOnly = true; // only sync events on weekdays (i.e., skip weekends)
var skipAllDayEvents = true; // only sync events with scheduled times (i.e., skip events that are all day or multi-day)
var removeEventReminders = false; // if true, this will disable reminders on the primary calendar for all hold events
var removeTravelReminders = false; // if true, this will disable reminders for travel buffers on the primary calendar

// some setup for what's included in the event description
// setting these to false makes the information from the personal/secondary calendar more private
var includeTitle = true; // description in the primary calendar will include the event title/name from the secondary calendar
var includeDescription = false; // description in the primary calendar will include the description from the secondary calendar
var descriptionStart = ''; // if you want certain text in every primary calendar hold description, e.g., "This is synced from [email]"

/*----- CUSTOMIZEABLE HELPER FUNCTIONS -----*/
// check if an event in the primary calendar is an existing hold for
// and event in the secondary calendar 
// not a very interesting function, but the hopt is that this could eventually be done
// better to resolve potential issue #3 above
function compareEvents(primaryEvent, secondaryEvent) {
  var primaryStart = primaryEvent.getStartTime().getTime();
  var primaryEnd = primaryEvent.getEndTime().getTime();
  var secondaryStart = secondaryEvent.getStartTime().getTime();
  var secondaryEnd = secondaryEvent.getEndTime().getTime();
  if ((primaryStart === secondaryStart) && (primaryEnd === secondaryEnd)) {
    return true;
  }
  return false;
}

// check whether an event is within the days that are being synced
// uses the syncWeekdaysOnly variable above
function checkEventDay(evi) {
  if (syncWeekdaysOnly) {
    var day = evi.getStartTime().getDay();
    // Logger.log(evi.getTitle()+': '+day)
    if ((day < 1) || (day > 5)) { // not a weekday
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

// set a primary event/hold description 
// optional argument: include what the "start" of the description should say
// e.g., if you want some defined text, like "This is a synced event"
function createDescription(evi, str) {
  if (str == null) {
    str = '';
  }
  if (includeTitle) {
    if (str.length > 0) {
      str = str + '\n\n'
    }
    str = str + evi.getTitle();
  }
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
function eventCreate(cal, title, startTime, endTime, removeFlag, opt_event) {
  try {
    var description = createDescription(opt_event, descriptionStart);
    var location = createLocation(opt_event);
    newEvent = cal.createEvent(title, startTime, endTime, {description: description, location: location})
  } catch (e) {
    newEvent = cal.createEvent(title, startTime, endTime, opt_event);
  }
  if (removeFlag) {newEvent.removeAllReminders;}
  newEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
  return newEvent; 
}

// function for updating events with location, description, and ensuring privacy
// newEvent = hold that's being put in the primary calendar
// secEvent = original event in the secondary/personal calendar
// removeFlag = boolean for whether to remove reminders
function eventUpdate(newEvent, removeFlag, secEvent) {
  newEvent.setDescription(createDescription(secEvent, descriptionStart));
  newEvent.setLocation(createLocation(secEvent));
  newEvent.setVisibility(CalendarApp.Visibility.PRIVATE); // set blocked time as private appointments in work calendar
  if (removeFlag) {
    newEvent.removeAllReminders();
  }
  return newEvent;
}

// function to delete an event
function eventDelete(evi, deleteStr) {
  var eviId = evi.getId();
  Logger.log(deleteStr + 
            '\nprimaryId: ' + eviId + '\nprimaryTitle: ' + evi.getTitle() + '\nprimaryDesc ' + evi.getDescription() + '\n');
  evi.deleteEvent();
  return eviId;
}



/*----- NOW THE MAIN FUNCTION -----*/
function sync() {

  var today=new Date();
  var enddate=new Date();
  enddate.setDate(today.getDate()+daysToSync); // how many days in advance to monitor and block off time
  
  var secondaryCal=CalendarApp.getCalendarById(id);
  var secondaryEvents=secondaryCal.getEvents(today,enddate);
  
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
          pEvent = eventUpdate(pEvent, removeEventReminders, evi);
          primaryEventsUpdated.push(pEvent.getId());
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
      var eviTitle = evi.getTitle();
      var newEvent = eventCreate(primaryCal, primaryEventTitle, eviStartTime, eviEndTime, removeEventReminders, evi);
      primaryEventsCreated.push(newEvent.getId());
      Logger.log('PRIMARY EVENT CREATED'
                 + '\nprimaryId: ' + newEvent.getId() + '\nprimaryTitle: ' + newEvent.getTitle() + '\nprimaryDesc ' + newEvent.getDescription() + '\n');

      if (newEvent.getLocation().trim() != '') { // if a location is set, block of 30 minutes before/after
        var travelStartTime = new Date(eviStartTime.getTime() - travelBuffer*minsToMilliseconds);
        var travelEndTime = new Date(eviEndTime.getTime() + travelBuffer*minsToMilliseconds);

        //travel time BEFORE
        var newEvent = eventCreate(primaryCal, primaryEventTravelTitle, travelStartTime, eviStartTime, 
                                    removeTravelReminders, {description: 'Travel buffer: '+eviTitle});
        primaryEventsTravelCreated.push(newEvent.getId());
        Logger.log('TRAVEL BUFFER CREATED (BEFORE)'
                 + '\nprimaryId: ' + newEvent.getId() + '\nprimaryTitle: ' + newEvent.getTitle() + '\nprimaryDesc ' + newEvent.getDescription() + '\n');

        //travel time AFTER
        var newEvent = eventCreate(primaryCal, primaryEventTravelTitle, eviEndTime, travelEndTime, 
                                    removeTravelReminders, {description: 'Travel buffer: '+eviTitle});
        primaryEventsTravelCreated.push(newEvent.getId());
        Logger.log('TRAVEL BUFFER CREATED (AFTER)'
                 + '\nprimaryId: ' + newEvent.getId() + '\nprimaryTitle: ' + newEvent.getTitle() + '\nprimaryDesc ' + newEvent.getDescription() + '\n');
      }
      
    }
  }

  // if a primary event previously created no longer exists in the secondary calendar, delete it
  for (pev in primaryEventsFiltered)
  {
    var pEvent = primaryEventsFiltered[pev];
    var pevIsUpdatedIndex = primaryEventsUpdated.indexOf(pEvent.getId());
    if (pevIsUpdatedIndex == -1) // it's not among the updated events
    { 
      if (primaryEventsFiltered[pev].getLocation().trim() != '') { // if it had a travel buffer, find those too
        pStartTime = pEvent.getStartTime().getTime();
        pEndTime = pEvent.getEndTime().getTime();
        for (pev2 in primaryEventsTravel) {
          var pTravel = primaryEventsTravel[pev2];
          if ((pTravel.getEndTime().getTime() == pStartTime) || (pTravel.getStartTime().getTime() == pEndTime)) {
            primaryEventsTravelDeleted.push(eventDelete(pTravel, 'TRAVEL BUFFER DELETED'));
          }
        }
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
