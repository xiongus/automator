set message_dates to {}
set message_subjects to {}
set message_senders to {}
set message_ids to {}
set message_urls to {}

tell application "Mail"
	--Get the messages currently selected in Mail
	set selected_messages to selection as list
	#set selected_messages to selected messages of message viewer 0
	--If no messages are selected, alert the user via a dialogue box and quit
	if (selected_messages is equal to missing value) then
		set alert_text to "No messages selected!"
		set alert_message to "Please select at least one message to send to copy and try again."
		display alert alert_text message alert_message as critical buttons {"OK"}
		return
	end if
	--Iterate over the selected messages and store their data in the lists
	repeat with selected_message in selected_messages
		set end of message_dates to date received of selected_message
		set end of message_subjects to subject of selected_message
		set end of message_senders to sender of selected_message
		set end of message_ids to selected_message's message id
		set end of message_urls to "message://%3C" & message id of selected_message & "%3e"
		#set end of message_urls to {"message://%3C" & my replaceText(message_ids, "%", "%25") & "%3E"}
	end repeat
end tell

set note_text to ""
--Iterate over the messages obtained from Mail
repeat with n from 1 to count of message_subjects
	--First line: Link to message with text based on message subject
	--Second line: message sender
	set message_note_text to ((item n of message_dates) & return & (item n of message_subjects) & return & (item n of message_senders) & return & (item n of message_urls) & return)
	--New line separator between messages
	set note_text to note_text & message_note_text & return
end repeat

#set my_list to note_text
#my Simple_sort(the my_list)

#set the clipboard to my_list as text
set the clipboard to note_text as text

on replaceText(subject, find, replace)
	set prevTIDs to text item delimiters of AppleScript
	set text item delimiters of AppleScript to find
	set subject to text items of subject
	
	set text item delimiters of AppleScript to replace
	set subject to "" & subject
	set text item delimiters of AppleScript to prevTIDs
	
	return subject
end replaceText

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