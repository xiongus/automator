##############################################
# Title: Create Reminder from a selected Mail message, check for existing Reminders
##############################################

# Iain Dunn
# Logic2Design
# www.logic2design.com
# logic2design@icloud.com

# Last update: 17 January 2022
# Version: 2

# Contributors and sources
# Rick0713 at https://discussions.apple.com/thread/3435695?start=30&tstart=0
# http://www.macosxautomation.com/applescript/sbrt/sbrt-06.html
# http://www.michaelkummer.com/2014/03/18/how-to-create-a-reminder-from-an-e-mail/

##############################################
# Configuration
##############################################
use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

# Flag colour, choose something between 1 and 7
# None = 0 | Orange = 1 | Red = 2 | Yellow = 3 | Blue = 4 | Purple = 5 | Green = 6 | Gray = 7
set flagIndex to 1

# set Archive Account Name
set accountName to "iCloud"

# Set to "on" to move the message to Archive
set switchArchive to "on"

#Set the Archive target mailbox
set Archive to "Archive"

# Set the default reminder date
# these are the possible choices: "2hours", "Tonight", "Tomorrow", "2 Days", "3 Days", "4 Days", "End of Week", "Next Monday", "1 Week", "2 Weeks", "1 Month", "2 Months", "3 Months", "Custom"
set defaultReminder to "1 Week"

# Set the default reminder time in hours after midnight
# for a reminder at "8:00 am" set "8", for "3 PM" or "15:00" set "15", for "8h45" set "8,75"
set defaultReminderTime to "9"

##############################################
# Code
##############################################
tell application "Mail"
	set selectedMessages to (selected messages of front message viewer)
	if (count of selectedMessages) is greater than 0 then
		set theMessage to item 1 of selectedMessages
		try
			tell theMessage
				set {theSubject, theFlag} to {the subject, the flag index}
				set theURL to "message://%3C" & message id of theMessage & "%3E"
			end tell
			
			# Check for existing Reminder
			if theFlag is not flagIndex then
				set theSubject to theMessage's subject
				
			else
				# Look for existing Reminder
				tell application "Reminders"
					
					set reminderCompleted to (name of reminders whose name is theSubject and completed is true)
					set reminderOpen to (name of reminders whose name is theSubject and completed is false)
					
					#Open Reminder found	
					if reminderOpen is not {} then
						set theButton to button returned of (display dialog "The Reminder is still active, would you like to mark it as complete or leave it open? " with title "Existing Reminder" buttons {"Complete", "Leave Open"} default button 2)
						#display dialog "Reminder still active" with title "Reminder Check" buttons {"OK"} default button 1
						if theButton is "Complete" then
							tell application "Mail"
								# unflag email/message
								set flag index of theMessage to -1
							end tell
							set theReminder to (last reminder whose body is theURL and completed is false)
							set completed of theReminder to true
							return
						else if theButton is "Leave Open" then
							return
						end if
						
						# Completed Reminder found
					else if reminderCompleted is not {} then
						tell me
							activate
						end tell
						set theButton to button returned of (display dialog "The selected email matches a completed reminder, would you like to clear the flag of this message?" with title "Clear Reminder Flag" buttons {"Mark complete", "Reopen Reminder"} default button 1)
						
						if theButton is "Mark complete" then
							tell application "Mail"
								# unflag email/message
								set flag index of theMessage to -1
							end tell
							return
						else
							# Set Followup Date/Time
							(choose from list {"2 Hours", "Tonight", "Tomorrow", "2 Days", "3 Days", "4 Days", "End of Week", "Next Monday", "1 Week", "2 Weeks", "1 Month", "2 Months", "3 Months", "Custom"} default items defaultReminder OK button name "Create" with prompt "Set follow-up time" with title "Create Reminder from E-Mail")
							
							set reminderDate to result as text
							
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
							
							set theReminder to (last reminder whose name is theSubject and completed is true)
							set completed of theReminder to false
							set remind me date of theReminder to remindMeDate
							return
						end if
					end if
				end tell
			end if
			
			# No Reminder found - will add new Reminder
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
			(choose from list {"2 Hours", "Tonight", "Tomorrow", "2 Days", "3 Days", "4 Days", "End of Week", "Next Monday", "1 Week", "2 Weeks", "1 Month", "2 Months", "3 Months", "Custom"} default items defaultReminder OK button name "Create" with prompt "Set follow-up time" with title "Create Reminder from E-Mail")
			
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
			#Create Reminder
			tell application "Reminders"
				
				tell list RemindersList
					# create new reminder with proper due date, subject name and the URL linking to the email in Mail
					make new reminder with properties {name:theSubject, remind me date:remindMeDate, body:theURL}
					
				end tell
				
				#	tell application "Reminders" to activate
			end tell
			tell application "Mail"
				# Flag selected Message in Mail
				set flag index of theMessage to flagIndex
				#Archive Message
				move theMessage to mailbox Archive of account accountName
			end tell
		end try
		
		
	end if
end tell

##############################################
# Functions
##############################################
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
		
	else if selectedDate = "End of Week" then
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
		set remindMeDate to (date (text returned of (display dialog (localized string "Enter the due date (e.g. 1/4/2017)") default answer "")))
		
	end if
	
	return remindMeDate
end chooseRemindMeDate

