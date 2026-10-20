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
on run argv
	set inFile to POSIX file (item 1 of argv)
	set outHFS to (POSIX file (item 2 of argv)) as text
	with timeout of 60 seconds
		tell application "Microsoft Excel"
			open inFile
			save workbook as active workbook filename outHFS file format workbook normal file format
			close active workbook saving no
		end tell
	end timeout
	return "done"
end run
