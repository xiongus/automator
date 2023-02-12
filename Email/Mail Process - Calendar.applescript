##############################################
# Title: Mail Process - Calendar
##############################################

# Iain Dunn
# Logic2Design
# www.logic2design.com
# logic2design@icloud.com

# Last update: 7 August 2022
# Version: 3 - Added "Now" option, Alerts and cleaned and limited Message content

# Contributors and sources
# http://www.michaelkummer.com/2014/03/18/how-to-create-a-reminder-from-an-e-mail/

##############################################
# Configuration
##############################################
use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

#Calendar
set the Calender_Lists to {"Business", "Home", "Finance", "Personal", "Recreation", "Work"} -- Add the Calendars that you want to schedule
set defaultCalendar to "Personal"

set defaultCalendarTime to "9"
set defaultCalendarDay to "Today"

set mailAccount to "Fastmail" -- Mail account

set alert1 to 1440 -- minutes
set alert2 to 15 -- minutes
##############################################
# Code
##############################################
set now to (current date)
set today to now - (time of now)
set tomorrow to (today) + (24 * 60 * 60)
set thisWeek to (current date) + (7 * days)
### Process selected Mail messages
tell application "Mail"
	set theMessages to the selection
	set theFolder to (POSIX path of (path to temporary items))
	set noSubjectString to "Mail Event - No Subject"
	repeat with theMessage in theMessages
		
		tell theMessage
			set {theDateReceived, theDateSent, theSender, theSubject, theContentText, theSource, theReadFlag} to {the date received, the date sent, the sender, the subject, the content, the source, the read status}
			
			set theURL to "message://%3C" & message id of theMessage & "%3E"
			if theSubject is equal to "" then set theSubject to noSubjectString
		end tell
		
		# Clean and limit Content length
		set tid to AppleScript's text item delimiters
		set AppleScript's text item delimiters to space
		set theText to (words of theContentText as rich text)
		set AppleScript's text item delimiters to tid
		if length of theText is less than 500 then
			set theContent to theText
		else
			set theContent to rich text 1 thru 500 of theText
		end if
		
		tell me to activate
		
		# Set Event Date
		(choose from list {"Now", "2 Hours", "Tonight", "Today", "Tomorrow", "2 Days", "3 Days", "4 Days", "5 Days", "6 Days", "Saturday", "Sunday", "Next Monday", "1 Week", "2 weeks", "1 Month", "2 Months", "3 Months", "Custom"} default items defaultCalendarDay OK button name "Select" with prompt "Set Event Date" with title "Time Block ")
		set reminderDate to result as rich text
		# exit if user clicks Cancel or Escape
		if reminderDate is "false" then return
		
		set theStartDate to my chooseRemindMeDate(reminderDate)
		
		# Set Event Time
		if reminderDate is "Now" then
			defaultCalendarTime
		else if reminderDate is "2 Hours" then
			defaultCalendarTime
		else if reminderDate is "Tonight" then
			defaultCalendarTime
		else
			set defaultCalendarTime to text returned of (display dialog "what time do you want to start? (Answer in decimal ie 2:30pm is 14.5)" default answer "9")
		end if
		
		if reminderDate is "Now" then
			set theStartDate to (current date)
		else if reminderDate is "2 Hours" then
			set theStartDate to (current date) + 2 * hours
		else if reminderDate is "Tonight" then
			set time of theStartDate to 60 * 60 * 17
		else
			set time of theStartDate to 60 * 60 * defaultCalendarTime
		end if
		
		#Set Event Length
		display dialog "How long is the Event? (minutes) " default answer 30
		
		set appt_length to text returned of result
		
		if appt_length < 1 then
			set appt_mins to (0)
			set theEndDate to theStartDate + (appt_mins * minutes)
			set allDay to true
		else
			set appt_mins to (appt_length)
			set theEndDate to theStartDate + (appt_mins * minutes)
			set allDay to false
		end if
		
		(choose from list Calender_Lists default items defaultCalendar OK button name "Select" with prompt "Pick Calendar" with title "Create Calendar Event")
		set Cal to result as rich text
		tell application "Calendar"
			
			tell calendar Cal
				make new event with properties {summary:theSubject, start date:theStartDate, end date:theEndDate, description:theContent, url:theURL}
			end tell
		end tell
		
		#Set Event Alerts
		tell application "Calendar"
			tell calendar Cal
				set theEvent to (first event where its summary = theSubject)
				tell theEvent
					--Alert 1
					make new display alarm at end of display alarms with properties {trigger interval:-alert1}
					--Alert 2
					make new sound alarm at end of sound alarms with properties {trigger interval:-alert2, sound name:"Sosumi"}
				end tell
			end tell
			reload calendars
		end tell
		
		#Move the Message to Archive
		set flagged status of theMessage to false -- Remove Flag
		set mailbox of theMessage to mailbox "Archive" of account mailAccount
		
	end repeat
end tell

##############################################
# Functions
##############################################

## Set Reminders Date
on chooseRemindMeDate(selectedDate)
	if selectedDate = "Now" then
		set remindMeDate to (current date)
		
	else if selectedDate = "2 Hours" then
		set remindMeDate to (current date)
		
	else if selectedDate = "Tonight" then
		set remindMeDate to (current date)
		
	else if selectedDate = "Today" then
		set remindMeDate to (current date)
		
	else if selectedDate = "Tomorrow" then
		set remindMeDate to (current date) + 1 * days
		
	else if selectedDate = "2 Days" then
		set remindMeDate to (current date) + 2 * days
		
	else if selectedDate = "3 Days" then
		set remindMeDate to (current date) + 3 * days
		
	else if selectedDate = "4 Days" then
		set remindMeDate to (current date) + 4 * days
		
	else if selectedDate = "5 Days" then
		set remindMeDate to (current date) + 5 * days
		
	else if selectedDate = "6 Days" then
		set remindMeDate to (current date) + 6 * days
		
	else if selectedDate = "Saturday" then
		# get the current day of the week
		set curWeekDay to weekday of (current date) as string
		if curWeekDay = "Monday" then
			set remindMeDate to (current date) + 5 * days
		else if curWeekDay = "Tuesday" then
			set remindMeDate to (current date) + 4 * days
		else if curWeekDay = "Wednesday" then
			set remindMeDate to (current date) + 3 * days
			# if it's Thursday, I'll set the reminder for Friday
		else if curWeekDay = "Thursday" then
			set remindMeDate to (current date) + 2 * days
			# if it's Friday I'll set the reminder for Thursday next week
		else if curWeekDay = "Friday" then
			set remindMeDate to (current date) + 1 * days
		else if curWeekDay = "Saturday" then
			set remindMeDate to (current date) + 7 * days
		else if curWeekDay = "Sunday" then
			set remindMeDate to (current date) + 6 * days
		end if
		
	else if selectedDate = "Sunday" then
		# end of week means Sunday in terms of reminders
		# get the current day of the week
		set curWeekDay to weekday of (current date) as string
		if curWeekDay = "Monday" then
			set remindMeDate to (current date) + 6 * days
		else if curWeekDay = "Tuesday" then
			set remindMeDate to (current date) + 5 * days
		else if curWeekDay = "Wednesday" then
			set remindMeDate to (current date) + 4 * days
			# if it's Thursday, I'll set the reminder for Friday
		else if curWeekDay = "Thursday" then
			set remindMeDate to (current date) + 3 * days
			# if it's Friday I'll set the reminder for Thursday next week
		else if curWeekDay = "Friday" then
			set remindMeDate to (current date) + 2 * days
		else if curWeekDay = "Saturday" then
			set remindMeDate to (current date) + 1 * days
		else if curWeekDay = "Sunday" then
			set remindMeDate to (current date) + 7 * days
		end if
		
	else if selectedDate = "Next Monday" then
		set curWeekDay to weekday of (current date) as string
		if curWeekDay = "Monday" then
			set remindMeDate to (current date) + 7 * days
		else if curWeekDay = "Tuesday" then
			set remindMeDate to (current date) + 6 * days
		else if curWeekDay = "Wednesday" then
			set remindMeDate to (current date) + 5 * days
		else if curWeekDay = "Thursday" then
			set remindMeDate to (current date) + 4 * days
		else if curWeekDay = "Friday" then
			set remindMeDate to (current date) + 3 * days
		else if curWeekDay = "Saturday" then
			set remindMeDate to (current date) + 2 * days
		else if curWeekDay = "Sunday" then
			set remindMeDate to (current date) + 1 * days
		end if
		
	else if selectedDate = "1 Week" then
		set remindMeDate to (current date) + 7 * days
		
	else if selectedDate = "2 Weeks" then
		set remindMeDate to (current date) + 14 * days
		
	else if selectedDate = "1 Month" then
		set remindMeDate to (current date) + 28 * days
		
	else if selectedDate = "2 Months" then
		set remindMeDate to (current date) + 56 * days
		
	else if selectedDate = "3 Months" then
		set remindMeDate to (current date) + 84 * days
		
	else if selectedDate = "Custom" then
		tell me to activate
		# adapt the date format suggested with what is configured in the user's 'Language/Region'-Preferences
		#Set Current Date & Time
		set todayDate to current date
		set reminderDate to short date string of todayDate as text
		set remindMeDate to (date (text returned of (display dialog (localized string "Enter the due date (e.g. 1/4/2021)") default answer reminderDate)))
		
	end if
	
	return remindMeDate
end chooseRemindMeDate