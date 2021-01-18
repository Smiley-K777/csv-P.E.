Option Explicit
    '****************LOCATION LOCATION********************
Dim ScriptDir : ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Dim FName : FName = inputbox ("Enter file name."&vbnewline&vbnewline&"Only CSV files can be accessed."&vbnewline&vbnewline&"Do not add (.csv) in file name.","SmileyK777","Testing File")
	if FName="" then Q
Dim outFile : outFile = ScriptDir & "\" & FName & ".csv"
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
'=======================================================================================
Dim NOP : NOP = 0
Dim Sposition : Sposition = 2
Dim Eposition : Eposition = 3
Dim objExcel : Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False'True	'Making Excel visible to user
Dim oCountWB : Set oCountWB = objExcel.Workbooks.Open(outFile)
Dim oCountWS : Set oCountWS = oCountWB.Worksheets(1)
Dim i : i=Eposition
do
Dim K : K=oCountWB.WorkSheets(1).Cells(Sposition,2).Value
Dim Y : Y=oCountWB.WorkSheets(1).Cells(Sposition,3).Value
Dim L : L=oCountWB.WorkSheets(1).Cells(Sposition,4).Value
Dim E : E=oCountWB.WorkSheets(1).Cells(Sposition,5).Value
Dim K1 : K1=oCountWB.WorkSheets(1).Cells(Eposition,2).Value
Dim Y1 : Y1=oCountWB.WorkSheets(1).Cells(Eposition,3).Value
Dim L1 : L1=oCountWB.WorkSheets(1).Cells(Eposition,4).Value
Dim E1 : E1=oCountWB.WorkSheets(1).Cells(Eposition,5).Value
Dim A : A =  K  & Y  & L  & E
Dim A1 : A1 = K1 & Y1 & L1 & E1
if A = A1 then
	'msgbox i &" "& Eposition & " its the same. " & A & " : " & A1
	Eposition=Eposition+1
	i=i+1
else
	Nop=Nop+1
	Dim PoutFile : PoutFile = ScriptDir & "\" & Nop & ". " & FName & ".csv"
	objFSO.CopyFile outFile, PoutFile
	Dim objExcelNFile : Set objExcelNFile = CreateObject("Excel.Application")
	objExcelNFile.Visible = False	'Making Excel visible to user
	Dim oCountWBNFile : Set oCountWBNFile = objExcelNFile.Workbooks.Open(PoutFile)
	Dim oCountWSNFile : Set oCountWSNFile = oCountWBNFile.Worksheets(1)
	Dim LastRow : LastRow = oCountWSNFile.UsedRange.Rows.Count
	Dim LastCol : LastCol = oCountWSNFile.UsedRange.Columns.Count
	if Nop=1 then
		'DELETE EVERYTHING AFTER END POSITION
		Dim P : P = Eposition + 1
		do
		If oCountWSNFile.Range("A" & P) = "" Then
			'msgbox "exit due to Eposition + 1 = nul"
			exit do
		else
			oCountWSNFile.Range("A" & Eposition & ":" & "A" & LastRow).EntireRow.Delete ' delete to LastRow
			'msgbox "Deleted till LastRow"
		end if
		loop
	else
		'DELETE EVERYTHING AFTER END POSITION
		P = Eposition + 1
		do
		If oCountWSNFile.Range("A" & P) = "" Then
			'msgbox "exit due to Eposition + 1 = nul"
			exit do
		else
			oCountWSNFile.Range("A" & Eposition & ":" & "A" & LastRow).EntireRow.Delete ' delete to LastRow
			'msgbox "Deleted till LastRow"
		end if
		loop
		'DELETE EVERYTHING BEFORE START POSITION
		Sposition = Sposition - 2
		oCountWSNFile.Range("A2:"&"A" & Sposition).EntireRow.Delete
	end if
	objExcelNFile.ActiveWorkbook.Save
	objExcelNFile.ActiveWorkbook.Close
	objExcelNFile.Application.Quit
	'MSGBOX"Position ["& Nop & "] has been Exported."

	LastRow = CInt(LastRow)
	if Eposition >= LastRow then
		Q
	else
		Sposition=Eposition+1
		Eposition=Sposition+1
	end if
end if
loop
Q
sub Q
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
WScript.Echo "SmileyK777 will now Close."
WScript.Quit
end sub
