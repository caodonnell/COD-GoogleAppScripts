# Google Calendar Sync script
I have a work calendar and a personal calendar. I don't want to put everything into a single calendar, and having to create holds
on my work calendar every time I addded something to my personal calendar was a pain. I tried a few different apps that can sync
calendars, but was never happy - either they only pulled events for the next 2 weeks (and I want to be able to schedule things
further out than that), or they didn't keep everything as private as I wanted. So, I wrote a script based on 
[this blog post](https://janelloi.com/auto-sync-google-calendar/) which uses 
[this script](https://gist.github.com/ttrahan/a88febc0538315b05346f4e3b35997f2).

## What does this script do? What does it not do?
This scripts takes data from a personal/secondary calendar and creates events/holds on a primary/work calendar to block off time 
for personal events and activities. It will also create events for travel/buffer time in the primary/work calendar for events 
that have a location. 

There are a few limits on the script as written:
* Travel buffer time is a fixed amount of time.
* All holds and travel buffers on the primary/work calendar require a unique identifier from the personal/secondary calendar. Otherwise, the script's methodology for recognizing holds are already present and for deleting holds will not function. The default is to use the event title from the personal/secondary calendar, but if the relevant flag is set to false (`includeTitle`), it will use the event id instead.
* The script currently ignores locations in the personal/secondary calendar that are Zoom URLs (which is determined by the presence of `.zoom.` in the location text). 
* All events on the secondary/personal calendar will produce a hold on the primary/work calendar. So, if you have 2 events at the same time listed on your secondary calender, there will be 2 holds at the same time on your primary calendar as well.
* The script will try to remove holds if events are removed from the secondary/personal calendar. To do this, the script checks event start/end times and the description to confirm if a hold on the primary calendar corresponds to something on your secondary/personal calendar. This means that if the scenario above happens (you have 2 or more events at the same time on your secondary/personal calendar, resulting in multiple holds at the same time on your primary/work calendar), and you delete one of the events from the secondary/personal calendar, it will try to delete the corresponding hold and travel buffers (if present) on the primary/work calendar by matching the time and the event description. However, I would consider this the most tenuous part of the script (i.e., the bit that's most likely to break).
* If there are orphan travel buffers in the primary/work calendar (e.g., you manually deleted the hold event it was associated with), the script will *not* find the orphan travel buffers and delete them. In this situation, if the original event is still on the personal/secondary calendar, the script will generate a new hold on the primary/work calendar (since there's a personal event without a matching hold) and will also create new travel buffers.

Finally, because this script just parses information from the two calendars, it's not the fastest script. A more efficient approach would create a
database that stores info about created holds/travel buffers, but that requires creating and managing an additional file in the
user's Google Drive.

## How to use

Note - if you have multiple personal Google calendars you want to sync, you'll have to repeat the steps below for each calendar.

### A. Set up your personal Google calendar so that it can be synced
1. In your personal calendar, access your [settings](https://calendar.google.com/calendar/u/0/r/settings).
2. Scroll down in the left sidebar and click on the calendar you want to sync.

> ![Google calendar settings - click on the calendar you want to sync](images/personal-cal-settings.png)

3. In that calendar's settings, find the "Share with specific people" section, and click on "Add people".
4. Enter your work email and make sure "See all event details" is the permission setting. Click "Send".

> ![Share your calendar with your work email](images/personal-cal-sharing.png)

5. We'll also need the Calendar ID for the next steps, so scroll through the calendar's settings until you see "Integrate calendar". The first item in that section will be "Calendar ID"; if this is the main calendar for a personal Gmail account, the ID will probably be your email address. Save your ID.

> ![Get your personal calendar ID](images/personal-cal-id.png)

### B. Set up the integration in your work Google calendar
1. Open your work email, which should now have something from Google notifying you that your personal email shared a calendar. Click on the "Add this calendar" link.
2. Open your work Google calendar. You should now see all of your personal calendar events listed in the calendar. If you're the only one looking at your work calendar, this is probably enough. However, in my organization, we often schedule events with other employees by looking at their Google calendars, which only includes their main calendar, and NOT synced calendars. If that's true for you, continue on with these steps.
3. Since we'll be adding events/holds to your main work Google calendar, you should hide the events from your personal calendar (otherwise, everything will show up twice, which can be annoying/overwhelming). Open the left sidebar (click the hamburger menu in the upper left corner if it isn't visible). Scroll down to the "Other calendars" section and uncheck the box next to the name of your personal calendar. I found that I also had to click on the three dots to the right of the calendar name and select "Hide from list" - otherwise, when I reloaded my work calendar, my personal calendar would still show up.

> ![Hide your personal calendar in your work calendar to avoid duplicate events being visible](images/work-cal-view.png)

### C. Adding the script to your Google account
1. Open up [My Drive](https://drive.google.com/).
2. In the upper left, click on "New" and click on "More" then "Google Apps Script".

> ![Create a new Google Apps Script in Google Drive](images/create-script.png)

3. Rename the script by clicking on the title "Untitled Project" and rename it something like "Calendar Sync" (or whatever makes the most sense to you - this process does not depend on the script name).
4. You should see `function myFunction()` and some brackets in the main window. Delete all of this text.

> ![Name the script, and delete the existing text](images/script-setup.png)

5. Copy and paste the script from [my code](https://github.com/caodonnell/COD-GoogleAppScripts/blob/main/GoogleCalendarSync/CalendarSync.gs) into the window. 
6. **YOU NEED TO ENTER YOUR CALENDAR ID IN LINE 4**. This is the ID you saved from A.5 above when checking out the settings of your personal calendar. **Replace the XXX, but make sure you don't remove the quotation marks or the semicolon at the end of the line.**
7. *(Optional)* Edit lines 5-7, taking care to ensure any semicolons are not deleted.
  - Line 5 is a variable for how many days out from today should the script sync events from your personal calendar. I set it to 70 (so, 10 weeks), which does mean the script can take a minute to run. Fewer days means that the script is faster, but that means if you're trying to plan something further out, things may not have synced yet from the personal calendar. 
  - Line 6 is a variable for how long of a travel buffer should be set before and after events with a location. The value of the variable should be in minutes. If you do not want travel buffers, make this a negative number (e.g., `-30` or `-1`).
  - Line 7 has a variable for whether the hold/buffer event descriptions in the primary/work calendar include the title of the event from the personal/secondary calendar (`true` means it will include the title, `false` means it will use an event id instead). 
8. Press the floppy disk icon to Save the code.
9. Click "Run" to execute the script. Google will ask you for permission to run the script, since it does read and edit your calendar. Please grant it all of the permissions it requires. As the script runs, there will be some output at the bottom of the screen in the Execution Log. If it completes successfully, the last line in the log should be yellow and say "Execution completed" (you may need to scroll down to see it). If the log ends with a red line and an error message, check that you copied the entire script and entered your calendar ID correctly.

> ![Copy and paste the calendar sync script, and edit line 4 to include your personal calendar ID. Then save and run the script before clicking on the clock icon to set up an automation](images/script-edit-v4.png)

10. Check that everything worked correctly by going through your calendar to see if events and travel buffers appear as expected.

### D. Set up automation so the script will run every time you update your personal calendar
1. In the Apps Script window, click the clock icon on the left sidebar. There should be a "Add trigger" button in the bottom right of the main window. 
2. Click on "Add Trigger".
3. Edit the trigger settings to be the following:
  - Choose which function to run: **mainFcn**
  - Which runs at deployment: **Head** (this should be the default setting)
  - Select event source: **From calendar**
  - Enter calendar details: **Calendar updated** and then enter the **email associated with your personal calendar**
  - The notification settings aren't super critical... I like it to notify me immediately if there's an issue, but a daily email is probably enough. If you do get this warnings, I'd recommend opening up the script from your Google Drive and running it to check out what's going on (like you did in C.9-10). For me, the most common reason I'd get an error was from doing too many sequential updates to my personal calendar (e.g., because I was rearranging multiple events), and since the script is triggered every time the personal calendar is updated, multiple changes mean multiple executions that can create conflicts. Manually running it afterwards generally resolved whatever was going on.
 4. As a caveat to the above, if you set the number of days to sync from your personal calendar to be a small number, you may want to add in a second trigger that runs on a time basis (e.g., a second trigger where the event source is "Time-driven", the type of time is a "Day timer", and it runs everyday "Midnight to 1am").
 
> ![Parameters for setting up a trigger so that the script will automatically run when you update your personal calendar](images/trigger-setup-v2.png)

### E. Optional: Change settings in the script
Lines 42-73 include a variety of things you can change, including
* Default titles for holds and travel buffers on the primary/work calendar 
* Default text for events with a location (that isn't a Zoom url)
* Set the travel buffer time (DO NOT CHANGE THE `const minsToMilliseconds` in line 31)
* Whether the script should skip all-day events and/or events on weekends. Note that setting `skipAllDayEvents = false` will only exclude multi-day events if they do not have set start/end times (e.g., something from 3pm on Wednesday through 4pm on Friday will still sync). Whether something occurs on a weekend is determined based on whether *both* the start and end times occurs on weekends. So, if `syncWeekdaysOnly = false`, something that starts on 3pm on Friday and ends on 3pm on Saturday will still sync, as will an event that is 3pm Sunday through 3pm Monday, but an event from 3pm Saturday to 3pm Sunday will *not* sync (since it's entirely over the weekend).
* Whether the script should sync events from the personal calendar during "Out of office" events on the primary/work calendar. Note that this setting requires OOO events to be created in the "Out of office" type.
* Whether the primary calendar should have default notifications enabled for holds and travel buffers
* Calendar event colors for holds and travel buffers in the primary/work calendar
* Whether the description for a hold on the primary/work calendar should have some default text (e.g., start with something about "This is a synced event"), or whether the hold description should also include the description from the secondary/personal calendar event. Note that setting `includeDescription = true` means that more information from your personal calendar will be present in your work calendar.
* The estimated maximum time it'll take for the script to run. If you set up the trigger so that the script runs every time the personal calendar is updated, then if you delete 2 separate events back-to-back (because you're rearranging a bunch of stuff), the script will run 2 times. However, that can cause an error at the end because things will be deleted and thus not exist by the time the 2nd execution works its way through. 
