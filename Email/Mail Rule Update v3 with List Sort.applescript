##############################################
# Title: Mail Rules Updater
##############################################

# Iain Dunn
# Logic2Design
# www.logic2design.com
# logic2design@icloud.com

# Last update: 4 August 2022
# Version: 3 - Added create new Rule option

# Contributors and sources
# 

##############################################
# Configuration
##############################################
set mailAccount to "Fastmail"

##############################################
# Code
##############################################
# Select Mail rule
tell application "Mail"
	set theList to name of every rule as list
end tell
bubblesort(theList)
copy "** New Rule ***" to the end of the theList
set rName to choose from list theList with prompt "Select Mail Rule" without empty selection allowed
if rName is false then
	return
else
	set theRule to result as text
end if

--Create New Rule
if theRule = "** New Rule ***" then
	set theResponse to display dialog "What's the name of the new Rule?" default answer "" with icon note buttons {"Continue"} default button "Continue"
	set theRuleName to text returned of theResponse as string
	
	set theResponse to display dialog "What's the name of the new Mail Folder? use / to create sub folders" default answer "" with icon note buttons {"Continue"} default button "Continue"
	set theFolderName to text returned of theResponse as string
	
	tell application "Mail"
		tell account mailAccount
			make new mailbox with properties {name:theFolderName}
		end tell
		set newRule to make new rule at beginning of rules with properties {name:theRuleName, enabled:true, should move message:true}
		set the theRule to theRuleName
		tell newRule
			make new rule condition at end of rule conditions with properties {rule type:from header, expression:"zzzz", qualifier:does contain value}
			set move message to (mailbox theFolderName of account mailAccount of application "Mail")
		end tell
	end tell
end if

# Select Rule type
set defaultTarget to "Sender"
(choose from list {"Domain", "Keyword", "Recipient", "Sender", "Subject"} default items defaultTarget OK button name "Select" with prompt "Select the Rule type" with title "Add to Mail Rule")
set theTarget to result as text

-- Update Rule
if theTarget = "Recipient" then
	tell application "Mail"
		# Start by getting the recipient's address and the message's account
		set theMessages to selection
		if theMessages is not {} then -- check for empty list
			try
				repeat with theSelectedMessage in theMessages
					set acct to account of mailbox of theSelectedMessage
					set emailAddr to address of first recipient of theSelectedMessage
					get acct
					# Add that address to a new condition of the rule
					set updateRule to rule theRule
					set ruleQualifier to expression of rule conditions of rule theRule # Check for existing rule for the email address
					if ruleQualifier does not contain emailAddr then
						tell updateRule
							set newCondition to make new rule condition at beginning of rule conditions
							tell newCondition
								set rule type to to header
								set qualifier to equal to value
								set expression to emailAddr
							end tell
						end tell
					end if
				end repeat
			on error errText
				display dialog "Warning: " & return & return & errText
			end try
		end if
	end tell
	# Sender
else if theTarget = "Sender" then
	tell application "Mail"
		# Start by getting the sender's address and the message's account
		set theMessages to selection
		if theMessages is not {} then -- check for empty list
			try
				repeat with theSelectedMessage in theMessages
					set acct to account of mailbox of theSelectedMessage
					set emailAddr to extract address from sender of theSelectedMessage
					get acct
					# Add that address to a new condition of the rule
					set updateRule to rule theRule
					set ruleQualifier to expression of rule conditions of rule theRule # Check for existing rule for the email address
					if ruleQualifier does not contain emailAddr then
						tell updateRule
							set newCondition to make new rule condition at beginning of rule conditions
							tell newCondition
								set rule type to from header
								set qualifier to equal to value
								set expression to emailAddr
							end tell
						end tell
					end if
				end repeat
			on error errText
				display dialog "Warning: " & return & return & errText
			end try
		end if
	end tell
	
	# Domain
else if theTarget = "Domain" then
	tell application "Mail"
		set theMessages to selection
		if theMessages is not {} then # check empty list
			try
				repeat with theSelectedMessage in theMessages
					-- Get the sender's domain and the message's account
					set acct to account of mailbox of theSelectedMessage
					set emailAddr to extract address from sender of theSelectedMessage
					set normDelims to AppleScript's text item delimiters
					set AppleScript's text item delimiters to "@"
					set theDomain to text item 2 of emailAddr
					set AppleScript's text item delimiters to normDelims
					get acct
					# Add that domain to a new condition of the rule
					set updateRule to rule theRule
					# Check for exisitung rule for the email/domain
					set ruleQualifier to expression of rule conditions of rule theRule
					if ruleQualifier does not contain theDomain then -- Check for existing rule for the domain
						tell updateRule
							set newCondition to make new rule condition at beginning of rule conditions
							tell newCondition
								set rule type to from header
								set qualifier to does contain value
								set expression to theDomain
							end tell
						end tell
					end if
				end repeat
			on error errText
				display dialog "Warning: " & return & return & errText
			end try
		end if
	end tell
	
	# Subject	
else if theTarget = "Subject" then
	tell application "Mail"
		-- Start by getting the message's subject
		set theMessages to selection
		if theMessages is not {} then -- check empty list
			try
				repeat with theSelectedMessage in theMessages
					set emailSubject to subject of theSelectedMessage
					-- Add that Subject to a new condition of the rule
					set updateRule to rule theRule
					set ruleQualifier to expression of rule conditions of rule theRule -- Check for existing rule for the Subject
					if ruleQualifier does not contain emailSubject then
						tell updateRule
							set newCondition to make new rule condition at beginning of rule conditions
							tell newCondition
								set rule type to subject header
								set qualifier to equal to value
								set expression to emailSubject
							end tell
						end tell
					end if
				end repeat
			on error errText
				display dialog "Warning: " & return & return & errText
			end try
		end if
	end tell
	
	# Keyword	
else if theTarget = "Keyword" then
	display dialog "What is the word or phrase to search for" default answer ""
	set theKeyword to text returned of result as text
	tell application "Mail"
		set theMessages to selection
		if theMessages is not {} then -- check empty list
			try
				repeat with theSelectedMessage in theMessages
					set acct to account of mailbox of theSelectedMessage
					set emailContent to content of theSelectedMessage
					get acct
					set updateRule to rule theRule
					set ruleQualifier to expression of rule conditions of rule theRule -- Check for existing rule for the content
					if ruleQualifier does not contain theKeyword then
						tell updateRule
							set newCondition to make new rule condition at beginning of rule conditions
							tell newCondition
								set rule type to message content
								set qualifier to does contain value
								set expression to theKeyword
							end tell
							set newCondition1 to make new rule condition at beginning of rule conditions
							tell newCondition1
								set rule type to subject header
								set qualifier to does contain value
								set expression to theKeyword
							end tell
						end tell
					end if
				end repeat
			on error errText
				display dialog "Warning: " & return & return & errText
			end try
		end if
	end tell
end if

# Process Messages with the new Rule
tell application "System Events"
	tell application "Mail" to activate
	keystroke "a" using {command down}
	keystroke "l" using {command down, option down}
end tell
##############################################
# Functions
##############################################
on bubblesort(theList)
	script o
		property lst : theList
	end script
	
	repeat with i from (count theList) to 2 by -1
		set a to beginning of o's lst
		repeat with j from 2 to i
			set b to item j of o's lst
			if (a > b) then
				set item (j - 1) of o's lst to b
				set item j of o's lst to a
			else
				set a to b
			end if
		end repeat
	end repeat
end bubblesort
