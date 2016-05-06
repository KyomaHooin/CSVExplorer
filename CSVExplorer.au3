;
;  CSV to GDI graph parser by Richard Bruna (c) 2016
;
;
; TODO
;
; x line pixelate
; -saving location
; -Line description

;TUNE

#NoTrayIcon
#AutoIt3Wrapper_Icon=toolbox.ico

;INCLUDE

#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <Array.au3>
#include <File.au3>
#include <GraphGDIPlus.au3>
#include <ScreenCapture.au3>

;VAR

global $csv_file = '' ;CSV soubor pro import
global $csv_array = '' ; CSV array
global $vsechny_mistnosti = '';Pole vsech mistnosti
global $vybrane_mistnosti = ''; vybrane pole mistnosti
global $colors[5][2] = [[0xFFbb0000,0xFF3db11a],[0xffffd700,0xffffff00],[0xff00ced1,0xfffa8072],[0xffc6e2ff,0xffff00ff],[0xffffa500, 0xfff0f8ff]]
global $screen = 1; screen capture increment

;GUI

$gui = GUICreate("CSV Explorer", 621, 416, 296, 246)
$button_csv = GUICtrlCreateButton("CSV", 16, 16, 99, 25)
$list_mistnosti = GUICtrlCreateList("", 8, 56, 113, 266,$LBS_SORT + $LBS_EXTENDEDSEL)
$button_graf = GUICtrlCreateButton("GRAF", 16, 336, 99, 25)
$button_ulozit = GUICtrlCreateButton("ULOŽIT", 16, 376, 99, 25)
$text_error = GUICtrlCreateLabel('', 157, 392, 300, 17)

;CONTROL

;Allready running?
If UBound(ProcessList("CSVExplorer.exe"), $UBOUND_ROWS) > 2 Then
	MsgBox(48, "", "Program již byl spuštěn!")
	Exit
EndIf

;MAIN

GUISetState(@SW_SHOW)
GUICtrlSetColor($text_error, 0xFF0000)

;create empty gaph and refresh it
Global $graf = _GraphGDIPlus_Create($gui, 155, 15, 450, 365, 0xFF000000, 0xFFFFFFFF)
_GraphGDIPlus_Refresh($graf)

While 1

	$event = GUIGetMsg()

	If $event = $button_csv Then
		;clear list
		GUICtrlSetData($list_mistnosti, '')
		;get file
		$csv_file = FileOpenDialog("Vyber CSV soubor", @HomeDrive, "CSV soubor (*.csv)", $FD_FILEMUSTEXIST,'',$gui)
		If @error Then
			GUICtrlSetData($text_error, "Nebyl vybrán žádný soubor!")
		Else
			;clear errror text
			GUICtrlSetData($text_error, '')
			;parse CSV to array
			_FileReadToArray($csv_file, $csv_array, $FRTA_NOCOUNT + $FRTA_INTARRAYS, ';')
			;parse location
			parse_location()
		EndIf
	EndIf

	If $event = $button_graf Then
		;entry control
		If Not $csv_file Then
			GUICtrlSetData($text_error, "Nebyl vybrán žádný soubor!")
		ElseIf ubound(_GUICtrlListBox_GetSelItems($list_mistnosti)) < 2  Then
			GUICtrlSetData($text_error, "Nebyla vybrána místnost!")
		Elseif ubound(_GUICtrlListBox_GetSelItems($list_mistnosti)) > 5 then
			GUICtrlSetData($text_error, "Příliš mnoho místností!")
		else
			;clear errror text if any
			GUICtrlSetData($text_error, '')
			;get select
			$vybrane_mistnosti = _GUICtrlListBox_GetSelItemsText($list_mistnosti)
			;define global data  array by slect size
			global $vsechna_data[UBound($vybrane_mistnosti)]
			;get dte array
			$vsechna_data[0] = parse_date()
			;get the rest of the data arrays
			for $i=1 to ubound($vybrane_mistnosti) - 1
				;MsgBox(-1,"selection", $vybrane_mistnosti[$i])
				$vsechna_data[$i] = parse_data($vybrane_mistnosti[$i])
			next
			;graph init
			graph_init()
			;graph all
			graph_all()
		EndIf
	EndIf

	If $event = $button_ulozit Then;
		if ubound($vybrane_mistnosti) >= 2 then
			$filename = '\' & StringRegExpReplace(($csv_array[1])[1],'(.*)\.csv', '$1') & $screen & '.png'
			_ScreenCapture_CaptureWnd(@ScriptDir & $filename, $gui,158, 37, 608, 402)
			$screen += 1
		else
			GUICtrlSetData($text_error, "Nebyla vybrána místnost!")
		EndIf
	EndIf

	If $event = $GUI_EVENT_CLOSE Then
		_GraphGDIPlus_Delete($gui, $graf)
		Exit
	EndIf

	WEnd

;FUNC

;parse locations from CSV file
Func parse_location()
	;fourth line of file split by semicolon
	$f = FileOpen($csv_file)
	;create location array
	Global $vsechny_mistnosti = StringSplit(FileReadLine($csv_file, 4), ';', $STR_NOCOUNT)
	;from third value
	For $i = 2 To UBound($vsechny_mistnosti) - 1
		;find all location by temperature capital
		If StringLeft($vsechny_mistnosti[$i], 1) == 'T' Then
			;remove prefix and populate list
			GUICtrlSetData($list_mistnosti, StringTrimLeft($vsechny_mistnosti[$i], 2))
		EndIf
	Next
	FileClose($f)
EndFunc   ;==>parse_location


;parse data by location from CSV file
Func parse_date()
	local $d_data = ''
	;_ArrayDisplay($csv_array)
	For $i = 6 To UBound($csv_array, $UBOUND_ROWS) - 1
		$d_data &= StringRegExpReplace(($csv_array[$i])[0], '([0-9]+.)([0-9]+.).*','$1$2') & ';'
	Next
	$d_data = StringSplit($d_data, ';', $STR_NOCOUNT)
	_ArrayPush($d_data, 'datum', 1)
	Return $d_data
EndFunc

;parse data by location from CSV file
Func parse_data($selection)
	local $teplota = '', $vlhkost = '';
	;get back real CSV string value from selected room
	$teplota = 'T ' & $selection
	if StringIsUpper(StringMid($selection,1,1))  then
		$vlhkost = 'H ' & StringLeft($selection, 1) & '2' & StringTrimLeft($selection, 1) ;StringRegExpReplace($selection,"([A-Z]).*","\1[2]")
	else
		$vlhkost = 'H ' & $selection
	EndIf
	;get column indexes, start at second value
	$t_col_index = _ArraySearch($vsechny_mistnosti, $teplota, 2)
	$h_col_index = _ArraySearch($vsechny_mistnosti, $vlhkost, 2)
	;get data arrays
	local $t_data = '', $h_data = ''
	;populate arrays/strings
	For $i = 6 To UBound($csv_array, $UBOUND_ROWS) - 1
		;test if idex exist
		if $t_col_index > -1 then $t_data &= ($csv_array[$i])[$t_col_index] & ';'
		if $h_col_index > -1 then $h_data &= ($csv_array[$i])[$h_col_index] & ';'
	Next
	$t_data = StringSplit($t_data, ';', $STR_NOCOUNT)
	$h_data = StringSplit($h_data, ';', $STR_NOCOUNT)
	;insert header and remove trailer
	_ArrayPush($t_data, $teplota, 1)
	_ArrayPush($h_data, $vlhkost, 1)
	local $data[2] = [$t_data,$h_data]
	return $data
EndFunc   ;==>parse_data

;grah init

Func graph_init()
	_GraphGDIPlus_Delete($gui, $graf)
	;_GraphGDIPlus_Clear($graf)
	$graf = _GraphGDIPlus_Create($gui, 155, 15, 450, 365, 0xFF000000, 0xFFFFFFFF)
	_GraphGDIPlus_Set_RangeY($graf, 0, 100, 10, 1, 0)

	$date_sort = parse_date()
	;$date_sort = _ArrayUnique(parse_date(),0,0,0,$ARRAYUNIQUE_NOCOUNT)
	_GraphGDIPlus_Set_RangeX($graf, 0, UBound($vsechna_data[0]) - 1, $date_sort, 15, 1, 0)

	_GraphGDIPlus_Set_PenColor($graf, 0xFFc0c0c0)
	_GraphGDIPlus_Set_PenSize($graf, 1)

	For $y = 10 To 90 Step 10
		$prvni = True
		For $x = 0 To UBound($vsechna_data[0]) - 1 Step 50
			If $prvni = True Then _GraphGDIPlus_Plot_Start($graf, $x, $y)
			$prvni = False
			_GraphGDIPlus_Plot_dot($graf, $x, $y)
		Next
		_GraphGDIPlus_Refresh($graf)
	Next
EndFunc   ;==>graph_init

;create GDI graph
Func graph_all()
	for $i=1 To UBound($vsechna_data) - 1
		_GraphGDIPlus_Set_PenColor($graf, $colors[$i-1][0])
		_GraphGDIPlus_Set_PenSize($graf, 2)

		;TEPLOTA
		$prvni = True
		For $x = 0 To UBound(($vsechna_data[$i])[0]) - 2 Step 15
			$y = (($vsechna_data[$i])[0])[$x + 1]
			If $prvni = True Then _GraphGDIPlus_Plot_Start($graf, $x, $y)
			$prvni = False
			_GraphGDIPlus_Plot_line($graf, $x, $y)
		Next
		_GraphGDIPlus_Refresh($graf)

		;VLHKOST

		;pokud je co zobrazit
		if UBound(($vsechna_data[$i])[1]) > 1 then
			_GraphGDIPlus_Set_PenColor($graf, $colors[$i-1][1])
			_GraphGDIPlus_Set_PenSize($graf, 2)
			$prvni = True
			For $x = 0 To UBound(($vsechna_data[$i])[1]) - 2 Step 15
				$y = (($vsechna_data[$i])[1])[$x + 1]
				If $prvni = True Then _GraphGDIPlus_Plot_Start($graf, $x, $y)
				$prvni = False
				_GraphGDIPlus_Plot_Line($graf, $x, $y)
			Next
			_GraphGDIPlus_Refresh($graf)
		EndIf
	Next
EndFunc   ;==>graph_create
