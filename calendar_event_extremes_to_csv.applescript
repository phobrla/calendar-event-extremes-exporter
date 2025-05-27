-- AppleScript: Export Calendar Event Extremes to Timestamped CSV
-- Author: Copilot, customized for general use
-- Description: This script exports the earliest and latest events in each Apple Calendar to a timestamped CSV file in the user's Documents folder.
-- Output columns: Email For Account, Calendar Name, Earliest Timestamp, Earliest Name, Latest Timestamp, Latest Name

-- Helper: Get current user's home directory
set homeDir to (POSIX path of (path to home folder))

-- Helper: Get current date and time for the filename (yyyyMMdd_HHmmss)
set currentDate to current date
set y to year of currentDate as string
set m to text -2 thru -1 of ("0" & ((month of currentDate as integer) as string))
set d to text -2 thru -1 of ("0" & ((day of currentDate as integer) as string))
set h to text -2 thru -1 of ("0" & ((hours of currentDate as integer) as string))
set min to text -2 thru -1 of ("0" & ((minutes of currentDate as integer) as string))
set s to text -2 thru -1 of ("0" & ((seconds of currentDate as integer) as string))
set timestamp to y & m & d & "_" & h & min & s

-- Construct output file path in the user's Documents folder
set filePath to homeDir & "Documents/calendar_event_extremes_" & timestamp & ".csv"

-- Prepare CSV headers
set csvText to "Email For Account,Calendar Name,Earliest Timestamp,Earliest Name,Latest Timestamp,Latest Name" & linefeed

-- Main logic: Gather calendar event extremes
tell application "Calendar"
	set allCalendars to calendars
	
	repeat with cal in allCalendars
		set calName to name of cal
		-- Try to get the account email, fallback to blank if not available
		try
			set acct to email of account of cal
			if acct is missing value then set acct to ""
		on error
			set acct to ""
		end try
		
		set calEvents to every event of cal
		if (count of calEvents) is 0 then
			set csvText to csvText & my csvRow(acct, calName, "", "", "", "") & linefeed
		else
			set earliestEvent to item 1 of calEvents
			set latestEvent to item 1 of calEvents
			repeat with ev in calEvents
				if start date of ev < start date of earliestEvent then
					set earliestEvent to ev
				end if
				if start date of ev > start date of latestEvent then
					set latestEvent to ev
				end if
			end repeat
			set earlyDate to start date of earliestEvent
			set earlyName to summary of earliestEvent as string
			set lateDate to start date of latestEvent
			set lateName to summary of latestEvent as string
			-- Format dates as ISO 8601
			set isoEarly to my toISO8601Date(earlyDate)
			set isoLate to my toISO8601Date(lateDate)
			set csvText to csvText & my csvRow(acct, calName, isoEarly, earlyName, isoLate, lateName) & linefeed
		end if
	end repeat
end tell

-- Write to file
try
	do shell script "echo " & quoted form of csvText & " > " & quoted form of filePath
	display dialog "Calendar event extremes exported to: " & filePath buttons {"OK"} default button 1 with icon note
on error errMsg
	display dialog "Error writing file: " & errMsg buttons {"OK"} default button 1 with icon stop
end try

-- Helper: Format a date as ISO 8601 (yyyy-MM-ddTHH:mm:ss)
on toISO8601Date(d)
	if d is "" then return ""
	set y to year of d as integer
	set m to text -2 thru -1 of ("0" & ((month of d as integer) as string))
	set da to text -2 thru -1 of ("0" & ((day of d as integer) as string))
	set h to text -2 thru -1 of ("0" & ((hours of d as integer) as string))
	set mi to text -2 thru -1 of ("0" & ((minutes of d as integer) as string))
	set s to text -2 thru -1 of ("0" & ((seconds of d as integer) as string))
	return y & "-" & m & "-" & da & "T" & h & ":" & mi & ":" & s
end toISO8601Date

-- Helper: Properly escape CSV fields with quotes
on escapeQuotes(t)
	set t to t as string
	set {tid, AppleScript's text item delimiters} to {AppleScript's text item delimiters, "\""}
	set parts to text items of t
	set AppleScript's text item delimiters to "\"\""
	set t to parts as string
	set AppleScript's text item delimiters to tid
	return t
end escapeQuotes

-- Helper: Generate a single CSV row from fields
on csvRow(email, calName, eTime, eName, lTime, lName)
	return "\"" & my escapeQuotes(email) & "\",\"" & my escapeQuotes(calName) & "\",\"" & my escapeQuotes(eTime) & "\",\"" & my escapeQuotes(eName) & "\",\"" & my escapeQuotes(lTime) & "\",\"" & my escapeQuotes(lName) & "\""
end csvRow
