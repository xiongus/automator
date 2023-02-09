###################################################################
# Title: Create Link(s) from Mail, Safari, Things 3, Devonthink or Finder
###################################################################

# Iain Dunn
# Logic2Design
# www.logic2design.com

# Last update: 2020-12-01
# Version: 1

####################################################################
# Configuration
####################################################################

set tabStart to 9 -- enter Safari tab number to start from, stops Pinned Tabs being selected

####################################################################
# Code
####################################################################
tell application "System Events"
	set activeApp to name of first application process whose frontmost is true
	
	-- Mail items
	
	if "Mail" is in activeApp then
		
		set message_subjects to {}
		set message_ids to {}
		set message_urls to {}
		
		tell application "Mail"
			--Get the messages currently selected in Mail
			set selected_messages to selection as list
			--Iterate over the selected messages and store their data in the lists
			repeat with selected_message in selected_messages
				set end of message_subjects to subject of selected_message
				set end of message_ids to selected_message's message id
				set end of message_urls to "message://%3C" & message id of selected_message & "%3e"
			end repeat
		end tell
		
		set note_text to ""
		--Iterate over the messages obtained from Mail
		repeat with n from 1 to count of message_subjects
			
			set message_note_text to ((item n of message_subjects) & " - " & (item n of message_urls) & return)
			--New line separator between messages
			set note_text to note_text & message_note_text & return
		end repeat
		
		#set the clipboard to my_list as text
		set the clipboard to note_text as text
		
		-- Web Pages
		
	else if "Safari" is in activeApp then
		
		set link_question to display dialog "Set link to current Tab, Not Pinned or All Tabs?" buttons {"Current", "Not Pinned", "All"} default button 1
		set link_answer to button returned of link_question as text
		
		if link_answer is equal to "Current" then
			
			tell application "Safari"
				set theURL to URL of front document
				set theTitle to name of front document
				set the clipboard to theTitle & " - " & theURL & return as string
			end tell
			
		else if link_answer is equal to "Not Pinned" then
			tell application "Safari"
				set docText to ""
				set windowCount to count (every window where visible is true)
				repeat with x from 1 to windowCount
					set tabCount to number of tabs in window x
					repeat with y from tabStart to tabCount
						set tabName to name of tab y of window x
						set tabURL to URL of tab y of window x as string
						set docText to docText & tabName & " Ð " & tabURL & return & return as string
					end repeat
					set the clipboard to the docText
				end repeat
			end tell
			
		else
			tell application "Safari"
				set docText to ""
				set windowCount to count (every window where visible is true)
				repeat with x from 1 to windowCount
					set tabCount to number of tabs in window x
					repeat with y from 1 to tabCount
						set tabName to name of tab y of window x
						set tabURL to URL of tab y of window x as string
						set docText to docText & tabName & " Ð " & tabURL & return & return as string
					end repeat
					set the clipboard to the docText
				end repeat
			end tell
		end if
		
		-- Things
		
	else if "Things3" is in activeApp then
		
		tell application "Things3"
			set selectToDos to selected to dos
			
			set linkToDo to {}
			
			repeat with r from 1 to length of selectToDos
				set this_todo to item r of selectToDos
				set todoName to (get name of this_todo) as text
				set todoID to (get id of this_todo as text)
				set todoURL to todoName & " - " & "things:///show?id=%22" & todoID & "%22" & return & return as text
				copy (todoURL) to end of linkToDo
			end repeat
			
			set the clipboard to linkToDo as text
			
		end tell
		
		-- Devonthink
		
	else if "DEVONThink 3" is in activeApp then

tell application id "DNtp"
	
	set newLink to {}
	set theSelection to selection
	repeat with theRecord in theSelection
		
		set recType to (type of theRecord) as string
		set recURL to reference URL of theRecord
		set recName to (name of theRecord)
		if recType = "PDF document" then
			set currentPage to current page of think window 1
			-- This is a zero-based index, so if you're on page 30, it will report 29
			set currentAttribute to "&?page=" & currentPage
		else if recType = "quicktime" then
			set currentTime to round (current time of think window 1 as real)
			set currentAttribute to "&?time=" & currentTime
		else
			set currentAttribute to ""
			-- Sets currentAttribute to a null value so nothing is appended to the URL.
		end if
		set the end of newLink to recName & " - " & recURL & "?reveal=1" & currentAttribute & return & return
		-- Note you could also add "?reveal=1" if you wanted to reveal the file in the database
		
	end repeat
	
	set the clipboard to newLink as text
	
end tell

		--Files & Folders	
		
	else
		set theSeparator to " - "
		
		set niceURL to {}
		
		tell application "Finder"
			# Create aliases from the Finder selection in the alias hub
			set theSelection to the selection as alias list
			repeat with i in theSelection
				set {theName, theURL} to {displayed name of i, URL of i}
				set the end of niceURL to theName & theSeparator & theURL & return & return
			end repeat
		end tell
		
		# Put it to the clipboard as text
		set {saveTID, AppleScript's text item delimiters} to {AppleScript's text item delimiters, {linefeed}}
		set the clipboard to niceURL as text
		set AppleScript's text item delimiters to saveTID
	end if
	
	--Files & Folders
	
end tell