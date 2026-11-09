-- Excel for Mac opens a workbook, sets or clears AutoFilters on it with its own object model,
-- and saves it again: what Excel itself writes for a filter (tests/oracle.py, Session.operate
-- with FILTER_SCRIPT, drives it; run only under its machine-wide lock).
--
--   osascript tools/excel_filter.applescript <in.xlsx> <out.xlsx> <steps>
--
-- <steps> is one or more steps joined by ";", each "kind|sheet|range|field|values":
--   filter|Stock|A1:F8|1|Lamp~Desk   autofilter range "A1:F8" field 1, the values (joined by
--                                    "~"; "=" is the blank cells) with operator filter by value
--   filter|Stock|A1:F8|3|=           ... one value "=" alone: criteria1 "=" (the blank cells)
--   filter|Stock|A1:F8||             autofilter range "A1:F8": the buttons, no criteria
--   off|Stock                        autofilter mode false: the filter off
--   showall|Stock                    show all data: the criteria cleared, the buttons kept
-- A table's filter is set on the table's range.  Prints "done" when every step ran and the
-- workbook was saved as <out.xlsx>; "error <n>: ..." when Excel refused one (nothing saved).
-- Opening and saving are as excel_resave.applescript does them.
on run argv
	set inFile to POSIX file (item 1 of argv)
	set outHFS to (POSIX file (item 2 of argv)) as text
	set AppleScript's text item delimiters to "/"
	set inName to last text item of (item 1 of argv)
	set outName to last text item of (item 2 of argv)
	set AppleScript's text item delimiters to ";"
	set stepList to text items of (item 3 of argv)
	set AppleScript's text item delimiters to ""
	try
		with timeout of 60 seconds
			tell application "Microsoft Excel"
				open inFile
				set names to name of every workbook
				if names does not contain inName then
					set AppleScript's text item delimiters to ", "
					return "wrong workbook: " & (names as text)
				end if
				set wb to workbook inName
				repeat with stepText in stepList
					set AppleScript's text item delimiters to "|"
					set p to text items of (stepText as text)
					set AppleScript's text item delimiters to ""
					set k to item 1 of p
					set ws to worksheet (item 2 of p) of wb
					if k is "filter" then
						set r to range (item 3 of p) of ws
						set f to item 4 of p
						set AppleScript's text item delimiters to "~"
						set crit to text items of (item 5 of p)
						set AppleScript's text item delimiters to ""
						if f is "" then
							autofilter range r
						else if (count of crit) is 1 and item 1 of crit is "=" then
							autofilter range r field (f as integer) criteria1 "="
						else
							autofilter range r field (f as integer) criteria1 crit operator filter by value
						end if
					else if k is "off" then
						set autofilter mode of ws to false
					else if k is "showall" then
						show all data ws
					else
						return "error 0: unknown step " & k
					end if
				end repeat
				save workbook as wb filename outHFS file format workbook normal file format
				close workbook outName saving no
			end tell
		end timeout
	on error message number code
		try
			tell application "Microsoft Excel" to close workbook inName saving no
		end try
		return "error " & code & ": " & message
	end try
	return "done"
end run
