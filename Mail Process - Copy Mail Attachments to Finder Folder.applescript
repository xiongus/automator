##############################################
# Title: Copy Mail Attachments to Finder Folder
##############################################

# Iain Dunn
# Logic2Design
# www.logic2design.com
# logic2design@icloud.com

# Last update: 1 August 2022
# Version: 1

# Contributors and sources
# 

##############################################
# Configuration
##############################################

##############################################
# Code
##############################################
tell application "Finder"
	set theOutputFolder to choose folder with prompt "Please select an output folder:"
	
	-- Select Mail Messages
	tell application "Mail"
		--activate
		set ListMessage to selection -- take all emails selected
		--set ListMessage to (every message in inbox whose subject contains ("motion" as list))
		repeat with aMessage in ListMessage -- loop through each message
			set AList to every mail attachment of aMessage
			repeat with aFile in AList --loop through each files attached to an email
				if (downloaded of aFile) then -- check if file is already downloaded
					
					set Filepath to theOutputFolder & (name of aFile) as rich text
					save aFile in file Filepath as native format
				end if
				delay 1
			end repeat -- next file
		end repeat
		
	end tell
	reveal theOutputFolder
	activate
end tell

tell application "System Events" to tell process "Dock"
	set frontmost to true
	tell list 1
		tell UI element "Finder"
			perform action "AXShowExpose"
		end tell
	end tell
end tell
##############################################
# Functions
##############################################
