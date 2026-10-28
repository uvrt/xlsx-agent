-- Excel for Mac opens a workbook and saves it again: the oracle's one question, asked of the
-- real application (tests/oracle.py drives it; run only under its machine-wide lock).
--
--   osascript tools/excel_resave.applescript <in.xlsx> <out.xlsx>
--
-- Prints "done" when Excel opened the file and saved it as <out.xlsx>.  Both paths are in
-- the Office group container, which Excel may read and write without a grant.
--
-- What was measured on Excel for Mac 16.106 (2026-10-20, the main session's probe):
-- * the file is opened with a plain `open` of a POSIX file computed OUTSIDE the tell block;
--   `open workbook workbook file name` and a POSIX-file coercion inside the tell give -50;
-- * alerts stay ON: a file Excel would repair shows its prompt ("We found a problem with
--   some content... recover?", Yes/No; Dutch Ja/Nee) within about two seconds, and the
--   caller, polling System Events, clicks No and reports "repair".  This script then fails
--   on its own (the workbook never opened), so the caller never has to kill it mid-call --
--   a killed osascript can leave an AppleEvent queued that starts Excel again later;
-- * `save workbook as ... file format workbook normal file format` writes a .xlsx.
-- * (2026-10-28) the workbook saved is the one opened, by name, never "active workbook":
--   Excel can start with a blank workbook ("Map1" in a Dutch Excel, its sheet "Blad1")
--   that stays active after `open` -- the input is open but not active -- and saving the
--   active workbook saved the blank one (seen in the runs after a repair prompt was ended
--   by quitting Excel; why Excel starts so is not known).  When no open workbook carries
--   the input's file name, the script prints "wrong workbook:" and the names it has, and
--   saves nothing; saved as, the workbook carries the new name, so it is closed by that;
-- * an AppleEvent time-out (-1712: Excel did not answer -- with the screen locked, a
--   repair prompt nobody could see or answer) is printed as "error -1712: ..." rather
--   than failing, so the caller tells it from other errors in any language.
on run argv
	set inFile to POSIX file (item 1 of argv)
	set outHFS to (POSIX file (item 2 of argv)) as text
	set AppleScript's text item delimiters to "/"
	set inName to last text item of (item 1 of argv)
	set outName to last text item of (item 2 of argv)
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
				save workbook as workbook inName filename outHFS file format workbook normal file format
				close workbook outName saving no -- saved as, it carries the new name
			end tell
		end timeout
	on error message number code
		return "error " & code & ": " & message
	end try
	return "done"
end run
