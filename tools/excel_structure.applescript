-- Excel for Mac opens a workbook, inserts, deletes or moves rows and columns on it, and saves
-- it again: what Excel itself makes of a structural edit (tests/oracle.py, Session.operate,
-- drives it; run only under its machine-wide lock).
--
--   osascript tools/excel_structure.applescript <in.xlsx> <out.xlsx> <steps>
--
-- <steps> is one or more steps joined by ";", each "kind|sheet|range|destination":
--   insert_rows|Stock|5:6|        insert into range "5:6", cells shifted down
--   insert_columns|Stock|C:D|     insert into range "C:D", cells shifted to the right
--   delete_rows|Stock|5:6|        delete range "5:6", cells shifted up
--   delete_columns|Stock|C:D|     delete range "C:D", cells shifted to the left
--   move|Stock|E:E|B:B            cut range "E:E", then insert into range "B:B": Excel's
--                                 Cut and "Insert Cut Cells" (the cut cells placed before B)
-- Prints "done" when every step ran and the workbook was saved as <out.xlsx>; "error <n>: ..."
-- when Excel refused a step (nothing is saved then).  Opening and saving are as
-- excel_resave.applescript does them (see its notes: the workbook saved by its name).
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
					set stepParts to text items of (stepText as text)
					set AppleScript's text item delimiters to ""
					set stepKind to item 1 of stepParts
					set ws to worksheet (item 2 of stepParts) of wb
					set stepRange to range (item 3 of stepParts) of ws
					if stepKind is "insert_rows" then
						insert into range stepRange shift shift down
					else if stepKind is "insert_columns" then
						insert into range stepRange shift shift to right
					else if stepKind is "delete_rows" then
						delete range stepRange shift shift up
					else if stepKind is "delete_columns" then
						delete range stepRange shift shift to left
					else if stepKind is "move" then
						cut range stepRange
						insert into range (range (item 4 of stepParts) of ws)
					else
						return "error 0: unknown step " & stepKind
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
