-- Excel for Mac opens a workbook, moves blocks of cells by Cut and "Insert Cut Cells" (a table's
-- columns among themselves, as dragging a table column does), and saves it again: what Excel
-- itself makes of reordering a table's columns (tests/oracle.py, Session.operate with
-- script=CELLS_SCRIPT, drives it; run only under its machine-wide lock).
--
--   osascript tools/excel_cells.applescript <in.xlsx> <out.xlsx> <steps>
--
-- <steps> is one or more steps joined by ";", each "kind|sheet|range|destination":
--   move_right|Shop|F3:F10|C3:C10  cut range "F3:F10", then insert into range "C3:C10",
--                                  the cells there shifted to the right
--   move|Stock|E:E|B:B             cut range "E:E", then insert into range "B:B"
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
					if stepKind is "move_right" then
						cut range stepRange
						insert into range (range (item 4 of stepParts) of ws) shift shift to right
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
