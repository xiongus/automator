##############################################
# Title: Mail Process - Reminders
##############################################

# Iain Dunn
# Logic2Design
# www.logic2design.com
# logic2design@icloud.com

# Last update: 5 August 2022
# Version: 1

# Contributors and sources
# Rick0713 at https://discussions.apple.com/thread/3435695?start=30&tstart=0
# http://www.macosxautomation.com/applescript/sbrt/sbrt-06.html
# http://www.michaelkummer.com/2014/03/18/how-to-create-a-reminder-from-an-e-mail/

##############################################
# Configuration
##############################################
--Mail Account
set mailAccount to "Fastmail"
-- Move to Archive "on" or "off" 
set switchArchive to "on"
--Archive
set Archive to "Archive"
--Flag choose something between 1 and 6- x-devonthink-item://907C008A-4695-4FD9-AED0-A026A1604AD6?reveal=1
set FlagIndex to 1

# Set the default reminder date
# these are the possible choices: "2hours", "Tonight", "Tomorrow", "2 Days", "3 Days", "4 Days", "5 Days", "6 Days","Saturday", "Sunday", ""Next Monday", "1 Week", "2 Weeks", "1 Month", "2 Months", "3 Months", "Custom"
set defaultReminder to "1 Week"

# Set the default reminder time in hours after midnight
# for a reminder at "8:00 am" set "8", for "3 PM" or "15:00" set "15", for "8h45" set "8,75"
set defaultReminderTime to "9"

# Shortcut to set Tag
-- https://www.icloud.com/shortcuts/e4cd641d071147d78a8ff0609fc24709
##############################################
# Code
##############################################
tell application "Mail"
	set theSelection to selection as list
	# do nothing if no email is selected in Mail
	try
		set theMessage to item 1 of theSelection
	on error
		return
	end try
	
	set theSubject to theMessage's subject
	set theUrl to "message://%3c" & (message id of theMessage) & "%3e"
	
	# Set Reminder Title
	if flag index of theMessage is not FlagIndex then
		
	else
		# Look for existing Reminder
		tell application "Reminders"
			set reminderCompleted to name of reminders whose name is theSubject and completed is true
			set reminderOpen to name of reminders whose name is theSubject and completed is false
			
			#Open Reminder found	
			if reminderOpen is not {} then
				set theButton to button returned of (display dialog "The Reminder is still active, would you like to mark it as complete or leave it open? " with title "Existing Reminder" buttons {"Complete", "Leave Open"} default button 2)
				if theButton is "Complete" then
					tell application "Mail"
						# unflag email/message
						set flag index of theMessage to -1
					end tell
					set theReminder to last reminder whose name is theSubject and completed is false
					# just in case 2 Reminders exist for the email
					set theReminderf to first reminder whose name is theSubject and completed is false
					set completed of theReminder to true
					set completed of theReminderf to true
					return
				else if theButton is "Leave Open" then
					return
				end if
				
				# Completed Reminder found
			else if reminderCompleted is not {} then
				tell me
					activate
				end tell
				set theButton to button returned of (display dialog "The selected email matches a completed Reminder, would you like to clear the flag of this message or create a new Reminder?" with title "Clear Reminder Flag?" buttons {"Mark complete", "Create new", "Cancel"} default button 1)
				
				if theButton is "Mark complete" then
					tell application "Mail"
						# unflag email/message
						set flag index of theMessage to -1
					end tell
					return
				else if theButton is "Cancel" then
					return
				end if
				
				# No Reminder found - will add new Reminder or Clear Flag
			else
				set theButton to button returned of (display dialog "No Reminder was found, do you want to set one?" with title "Reminder Check" buttons {"Create new", "Remove Flag"} default button 1)
				if theButton is "Remove Flag" then
					tell application "Mail"
						# unflag email/message
						set flag index of theMessage to -1
					end tell
					return
				end if
			end if
		end tell
	end if
	
	# Select Reminders List
	tell application "Reminders"
		set lName to name of every list
		set dName to name of default list
	end tell
	tell me to activate
	set lName to choose from list lName with prompt "Select Reminder List" default items {dName} without empty selection allowed
	if lName is false then
		return 1
	else
		set lName to lName as string
	end if
	set RemindersList to lName as rich text
	
	# Set Followup Date/Time
	(choose from list {"2 Hours", "Tonight", "Tomorrow", "2 Days", "3 Days", "4 Days", "5 Days", "6 Days", "Saturday", "Sunday", "Next Monday", "1 Week", "2 Weeks", "1 Month", "2 Months", "3 Months", "Custom"} default items defaultReminder OK button name "Create" with prompt "Set follow-up time" with title "Create Reminder from E-Mail")
	set reminderDate to result as rich text
	
	# exit if user clicks Cancel or Escape
	if reminderDate is "false" then return
	# for all the other options, calculate the date based on the current date
	set remindMeDate to my chooseRemindMeDate(reminderDate)
	
	# set the time for on the reminder date
	if reminderDate is "2 Hours" then
		set remindMeDate to (current date) + 2 * hours
	else if reminderDate is "Tonight" then
		set time of remindMeDate to 60 * 60 * 17
	else
		set time of remindMeDate to 60 * 60 * defaultReminderTime
	end if
	
	# Flag selected email/message in Mail
	set flag index of theMessage to FlagIndex
	
	# Create Mail URL
	set theUrl to "message://%3c" & (message id of theMessage) & "%3e"
	
	# Move to Archive if variable = "On"
	if switchArchive is "on" then
		move theMessage to mailbox Archive of account mailAccount
	end if
end tell

# Create the new Reminder
tell application "Reminders"
	
	tell list RemindersList
		# create new reminder with due date, subject name and the URL linking to the email in Mail
		make new reminder with properties {name:theSubject, remind me date:remindMeDate, body:theUrl}
		
	end tell
	
end tell

# Set Reminders Tag
tell application "Shortcuts Events"
	run the shortcut named "Reminders set Purchases Tag"
end tell

tell application "Reminders" to activate

##############################################
# Functions
##############################################
# date calculation with the selection from the dialogue
# use to set the initial and the re-scheduled date
on chooseRemindMeDate(selectedDate)
	if selectedDate = "2 Hours" then
		set remindMeDate to (current date) + 0 * days
		--(current date) + 2 * hours
		--set time of remindMeDate to 120 * minutes
		
	else if selectedDate = "Tonight" then
		# add 0 day and set time to 17h into the day = 5pm
		set remindMeDate to (current date) + 0 * days
		--set time of remindMeDate to 60 * 60 * 17
		
	else if selectedDate = "Tomorrow" then
		# add 1 day and set time to 9h into the day = 9am
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
		# adapt the date format suggested with what is configured in the user's 'Language/Region'-Preferences
		#Set Current Date & Time
		set todayDate to current date
		set reminderDate to short date string of todayDate as text
		set remindMeDate to (date (text returned of (display dialog (localized string "Enter the due date (e.g. 1/4/2021)") default answer reminderDate)))
		
	end if
	
	return remindMeDate
end chooseRemindMeDate
