##############################################
# Title: Create Contact(s) from Mail Message
##############################################

# Iain Dunn
# Logic2Design
# www.logic2design.com
# logic2design@icloud.com

# Last update: 13 April 2022
# Version: 2

# Contributors and sources
# 

##############################################
# Configuration
##############################################

##############################################
# Code
##############################################
tell application "Contacts" to set the_list to name of every group
set old_delims to AppleScript's text item delimiters
set AppleScript's text item delimiters to {ASCII character 10} -- always a linefeed
set list_string to (the_list as string)
set new_string to do shell script "echo " & quoted form of list_string & " | sort -f"
set new_list to (paragraphs of new_string)
set AppleScript's text item delimiters to old_delims


set thePrompt to "Select the group(s) to which to add the sender of the selected message(s)."
my Simple_sort(new_list)
set R to choose from list new_list with prompt thePrompt with multiple selections allowed
if R is false then return
if R contains {"Family"} or R contains {"Friends"} then
	
	# Personal Email
	tell application "Mail"
		set theMessages to selection
		--repeat with a in (theMessages)
		
		if theMessages is not {} then -- check empty list
			set loop to 1
			repeat with i from 1 to the number of items in theMessages
				--repeat with selectedmessage in items of the theMessages
				--set flagged status of item loop of theMessages to false
				set theSenderName to extract name from sender of item loop of theMessages
				set nameArray to my split(theSenderName, " ")
				set theFirstName to item 1 of nameArray
				set theLastName to last item of nameArray
				set theEmail to extract address from sender of item loop of theMessages
				tell application "Contacts"
					set thePerson to make new person with properties {first name:theFirstName, last name:theLastName}
					make new email at end of emails of thePerson with properties {label:"home", value:theEmail}
					
					repeat with theGroupName in items of R
						add (item 1 of thePerson) to group theGroupName
					end repeat
					save
					set selected of group theGroupName to true
				end tell
				set loop to loop + 1
			end repeat
		end if
		
	end tell
	
	#Business Email	
else
	tell application "Mail"
		set theMessages to selection
		if theMessages is not {} then -- check empty list
			--set loop to 1
			repeat with i from 1 to the number of items in theMessages
				--set flagged status of item loop of theMessages to false
				set theSenderName to extract name from sender of item i of theMessages
				set theEmail to extract address from sender of item i of theMessages
				
				tell application "Contacts"
					
					set thePerson to make new person with properties {organization:theSenderName}
					make new email at end of emails of thePerson with properties {label:"work", value:theEmail}
					
					try
						repeat with theGroupName in items of R
							add (item 1 of thePerson) to group theGroupName
						end repeat
					end try
					save
					set selected of group theGroupName to true
				end tell
				
				
				--set loop to loop + 1
			end repeat
		end if
		
	end tell
	
end if

# Run Mail Rules
tell application "System Events"
	tell application "Mail" to activate
	keystroke "a" using {command down}
	keystroke "l" using {command down, option down}
end tell

##############################################
# Functions
##############################################
# Sort the Group list
on Simple_sort(my_list)
	set the index_list to {}
	set the sorted_list to {}
	repeat (the number of items in my_list) times
		set the low_item to ""
		repeat with i from 1 to (number of items in my_list)
			if i is not in the index_list then
				set this_item to item i of my_list as text
				if the low_item is "" then
					set the low_item to this_item
					set the low_item_index to i
				else if this_item comes before the low_item then
					set the low_item to this_item
					set the low_item_index to i
				end if
			end if
		end repeat
		set the end of sorted_list to the low_item
		set the end of the index_list to the low_item_index
	end repeat
	return the sorted_list
end Simple_sort

#Split Sender name
on split(theString, theDelimiter)
	set oldDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to theDelimiter
	set theArray to every text item of theString
	set AppleScript's text item delimiters to oldDelimiters
	return theArray
end split