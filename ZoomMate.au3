#AutoIt3Wrapper_UseX64=y
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <Date.au3>
#include <StringConstants.au3>
#include <TrayConstants.au3>
#include <GUIConstantsEx.au3>
#include "Includes\UIA_Functions-a.au3"
#include "Includes\CUIAutomation2.au3"

Opt("MustDeclareVars", 1)
Opt("GUIOnEventMode", 1)

; Windows messages for per-keystroke validation
Global Const $WM_COMMAND = 0x0111
Global Const $EN_CHANGE = 0x0300

; -------------------------
; Script configuration
; -------------------------
Global Const $CONFIG_FILE = @ScriptDir & "\zoom_config.ini"

; -------------------------
; Runtime variables
; -------------------------
Global $g_PrePostSettingsConfigured = False
Global $g_DuringMeetingSettingsConfigured = False

Global $g_UserSettings = ObjCreate("Scripting.Dictionary")

; i18n containers
Global $g_Languages = ObjCreate("Scripting.Dictionary") ; langCode -> dict of key->string
Global $g_LangCodeToName = ObjCreate("Scripting.Dictionary") ; langCode -> display name
Global $g_LangNameToCode = ObjCreate("Scripting.Dictionary") ; display name -> langCode
Global $g_CurrentLang = "en"

Func GetUserSetting($key)
	If $g_UserSettings.Exists($key) Then Return $g_UserSettings.Item($key)
	Return ""
EndFunc   ;==>GetUserSetting

; translation lookup with fallback to English, supports placeholders {0},{1},{2}
Func t($key, $p0 = Default, $p1 = Default, $p2 = Default)
	If $g_Languages.Exists($g_CurrentLang) Then
		Local $oDict = $g_Languages.Item($g_CurrentLang)
		If $oDict.Exists($key) Then
			Local $s = $oDict.Item($key)
			If $p0 <> Default Then $s = StringReplace($s, "{0}", $p0)
			If $p1 <> Default Then $s = StringReplace($s, "{1}", $p1)
			If $p2 <> Default Then $s = StringReplace($s, "{2}", $p2)
			Return $s
		EndIf
	EndIf
	If $g_Languages.Exists("en") Then
		Local $oEn = $g_Languages.Item("en")
		If $oEn.Exists($key) Then
			Local $s2 = $oEn.Item($key)
			If $p0 <> Default Then $s2 = StringReplace($s2, "{0}", $p0)
			If $p1 <> Default Then $s2 = StringReplace($s2, "{1}", $p1)
			If $p2 <> Default Then $s2 = StringReplace($s2, "{2}", $p2)
			Return $s2
		EndIf
	EndIf
	Return $key
EndFunc   ;==>t

Func _LoadTranslations()
	Local $sDir = @ScriptDir & "\i18n\*.ini"
	Local $hSearch = FileFindFirstFile($sDir)
	If $hSearch = -1 Then Return
	While 1
		Local $sFile = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		Local $lang = StringTrimRight($sFile, 4)
		Local $fullPath = @ScriptDir & "\i18n\" & $sFile
		Local $a = IniReadSection($fullPath, "translations")
		If @error Then ContinueLoop
		Local $dict = ObjCreate("Scripting.Dictionary")
		For $i = 1 To $a[0][0]
			$dict.Add($a[$i][0], $a[$i][1])
		Next
		If $g_Languages.Exists($lang) Then $g_Languages.Remove($lang)
		$g_Languages.Add($lang, $dict)

		; map display names
		Local $langName = ""
		If $dict.Exists("LANGNAME") Then $langName = $dict.Item("LANGNAME")
		If $langName = "" Then $langName = $lang
		If $g_LangCodeToName.Exists($lang) Then $g_LangCodeToName.Remove($lang)
		$g_LangCodeToName.Add($lang, $langName)
		If $g_LangNameToCode.Exists($langName) Then $g_LangNameToCode.Remove($langName)
		$g_LangNameToCode.Add($langName, $lang)
	WEnd
	FileClose($hSearch)
EndFunc   ;==>_LoadTranslations

Func _ListAvailableLanguageNames()
	Local $list = ""
	For $name In $g_LangNameToCode.Keys
		$list &= ($list = "" ? $name : "," & $name)
	Next
	Return $list
EndFunc   ;==>_ListAvailableLanguageNames

Func _GetLanguageDisplayName($code)
	If $g_LangCodeToName.Exists($code) Then Return $g_LangCodeToName.Item($code)
	Return $code
EndFunc   ;==>_GetLanguageDisplayName

Global $idSaveBtn
Global $idLanguagePicker

Global $g_FieldCtrls = ObjCreate("Scripting.Dictionary")
Global $g_DayLabelToNum = ObjCreate("Scripting.Dictionary")
Global $g_DayNumToLabel = ObjCreate("Scripting.Dictionary")
Global $g_ErrorAreaLabel = 0


Func _InitDayLabelMaps()
	; Build maps 1..7 using translations DAY_1..DAY_7
	Local $i
	For $i = 1 To 7
		Local $key = "DAY_" & $i
		Local $label = t($key)
		If Not $g_DayLabelToNum.Exists($label) Then $g_DayLabelToNum.Add($label, $i)
		If Not $g_DayNumToLabel.Exists(String($i)) Then $g_DayNumToLabel.Add(String($i), $label)
	Next
EndFunc   ;==>_InitDayLabelMaps

; -------------------------
; Initialize UIAutomation COM
; -------------------------
Global $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation)
If Not IsObj($oUIAutomation) Then
	Debug("Failed to create UIAutomation COM object.", "UIA")
	Exit
EndIf
Debug("UIAutomation COM created successfully.", "UIA")

Global $pDesktop
$oUIAutomation.GetRootElement($pDesktop)
Global $oDesktop = ObjCreateInterface($pDesktop, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
If Not IsObj($oDesktop) Then
	Debug(t("ERROR_GET_DESKTOP_ELEMENT_FAILED"), "ERROR")
	Exit
EndIf
Debug("Desktop element obtained.", "UIA")
Global $oZoomWindow = 0

; -------------------------
; Load or prompt meeting configuration (via GUI)
; -------------------------
Global $g_StatusMsg = "Idle"
Global $g_TrayIcon = @ScriptDir & "\zoommate.ico" ; Place your custom icon here

TraySetIcon($g_TrayIcon)
TraySetClick($TRAY_CLICK_SECONDARYDOWN)

Func UpdateTrayTooltip()
	TraySetToolTip("ZoomMate: " & $g_StatusMsg)
EndFunc   ;==>UpdateTrayTooltip

Global $g_ConfigGUI = 0

Func Debug($string, $type = "DEBUG", $noNotify = False)
	If ($string) Then
		ConsoleWrite("[" & $type & "] " & $string & @CRLF)
		If $type = "INFO" Or $type = "ERROR" Then
			$g_StatusMsg = $string
			If Not $noNotify Then
				TrayTip("ZoomMate", $string, 5, ($type = "INFO" ? 1 : 3))
			EndIf
		EndIf
		If $type = "ERROR" Then
			TraySetIcon($g_TrayIcon, 1) ; Error icon
		EndIf
	EndIf
EndFunc   ;==>Debug


Func ShowConfigGUI()
	If $g_ConfigGUI Then
		GUICtrlSetState($g_ConfigGUI, @SW_SHOW)
		Return
	EndIf

	_InitDayLabelMaps()
	$g_ConfigGUI = GUICreate(t("CONFIG_TITLE"), 320, 460)

	GUICtrlCreateLabel("Language:", 10, 160, 120, 20)
	$idLanguagePicker = GUICtrlCreateCombo("", 140, 160, 160, 20)
	GUICtrlSetData($idLanguagePicker, _ListAvailableLanguageNames(), _GetLanguageDisplayName(GetUserSetting("Language")))

	; Dynamically add text input fields and track their control IDs by key
	_AddTextInputField("MeetingID", t("LABEL_MEETING_ID"), 10, 10, 140, 10, 160)

	_AddDayDropdownField("MidweekDay", t("LABEL_MIDWEEK_DAY"), 10, 40, 200, 40, 100)
	_AddTextInputField("MidweekTime", t("LABEL_MIDWEEK_TIME"), 10, 70, 200, 70, 100)

	_AddDayDropdownField("WeekendDay", t("LABEL_WEEKEND_DAY"), 10, 100, 200, 100, 100)
	_AddTextInputField("WeekendTime", t("LABEL_WEEKEND_TIME"), 10, 130, 200, 130, 100)

	_AddTextInputField("HostToolsValue", t("LABEL_HOST_TOOLS"), 10, 190, 140, 190, 160)
	_AddTextInputField("ParticipantValue", t("LABEL_PARTICIPANT"), 10, 220, 140, 220, 160)
	_AddTextInputField("MuteAllValue", t("LABEL_MUTE_All"), 10, 250, 140, 250, 160)
	_AddTextInputField("YesValue", t("LABEL_YES"), 10, 280, 140, 280, 160)

	; Error area just above buttons (full width)
	$g_ErrorAreaLabel = GUICtrlCreateLabel("", 10, 310, 300, 20)
	GUICtrlSetColor($g_ErrorAreaLabel, 0xFF0000)

	$idSaveBtn = GUICtrlCreateButton(t("BTN_SAVE"), 60, 340, 80, 30)
	Local $idQuitBtn = GUICtrlCreateButton(t("BTN_QUIT"), 180, 340, 80, 30)

	; Save button is disabled by default
	GUICtrlSetState($idSaveBtn, $GUI_DISABLE)
	GUICtrlSetState($idQuitBtn, $GUI_ENABLE)

	GUICtrlSetOnEvent($idSaveBtn, "SaveConfigGUI")
	GUICtrlSetOnEvent($idQuitBtn, "QuitApp")

	CheckConfigFields() ; Initial check


	GUISetState(@SW_SHOW, $g_ConfigGUI)

	; Ensure per-keystroke validation via EN_CHANGE for Edit controls
	GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND_EditChange")
EndFunc   ;==>ShowConfigGUI

; Helper to add a label + input, and register the input control ID under a key
Func _AddTextInputField($key, $label, $xLabel, $yLabel, $xInput, $yInput, $wInput)
	GUICtrlCreateLabel($label, $xLabel, $yLabel, 180, 20)
	Local $idInput = GUICtrlCreateInput(GetUserSetting($key), $xInput, $yInput, $wInput, 20)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idInput)
	GUICtrlSetOnEvent($idInput, "CheckConfigFields")
EndFunc   ;==>_AddTextInputField

Func _AddDayDropdownField($key, $label, $xLabel, $yLabel, $xInput, $yInput, $wInput)
	GUICtrlCreateLabel($label, $xLabel, $yLabel, 180, 20)
	; Build day list from maps
	Local $dayList = ""
	Local $i
	; Show Monday (2) through Saturday (7) first
	For $i = 2 To 7
		Local $lbl = t("DAY_" & $i)
		$dayList &= ($dayList = "" ? $lbl : "|" & $lbl)
	Next
	; Then Sunday (1) last for UI niceness
	Local $lblSun = t("DAY_" & 1)
	$dayList &= ($dayList = "" ? $lblSun : "|" & $lblSun)

	Local $currentNum = String(GetUserSetting($key))
	Local $currentLabel = $currentNum
	If $g_DayNumToLabel.Exists($currentNum) Then $currentLabel = $g_DayNumToLabel.Item($currentNum)

	Local $idCombo = GUICtrlCreateCombo("", $xInput, $yInput, $wInput, 20)
	GUICtrlSetData($idCombo, $dayList, $currentLabel)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idCombo)
	GUICtrlSetOnEvent($idCombo, "CheckConfigFields")
EndFunc   ;==>_AddDayDropdownField

Func _IsValidMeetingID($s)
	$s = StringStripWS($s, 3)
	If $s = "" Then Return False
	If Not StringRegExp($s, "^\d{9,11}$") Then Return False
	Return True
EndFunc   ;==>_IsValidMeetingID

Func _IsValidTime($s)
	$s = StringStripWS($s, 3)
	If $s = "" Then Return False
	If Not StringRegExp($s, "^(\d{1,2}):(\d{2})$") Then Return False
	Local $a = StringSplit($s, ":")
	Local $h = Number($a[1])
	Local $m = Number($a[2])
	If $h < 0 Or $h > 23 Then Return False
	If $m < 0 Or $m > 59 Then Return False
	Return True
EndFunc   ;==>_IsValidTime

Func CheckConfigFields()
	Local $allFilled = True
	For $sKey In $g_FieldCtrls.Keys
		Local $ctrlId = $g_FieldCtrls.Item($sKey)
		Local $val = StringStripWS(GUICtrlRead($ctrlId), 3)
		If $val = "" Then
			$allFilled = False
			; mark required field with tooltip and light red background (no text color change)
			GUICtrlSetTip($ctrlId, t("ERROR_REQUIRED"))
			GUICtrlSetBkColor($ctrlId, 0xEEDDDD)
		EndIf
	Next

	; Additional format validation
	If $allFilled Then
		Local $ok = True
		; MeetingID must be 9-11 digits
		Local $idCtrl = $g_FieldCtrls.Item("MeetingID")
		If Not _IsValidMeetingID(GUICtrlRead($idCtrl)) Then $ok = False

		; Times must be valid HH:MM (24h)
		Local $midCtrl = $g_FieldCtrls.Item("MidweekTime")
		Local $wkdCtrl = $g_FieldCtrls.Item("WeekendTime")
		If Not _IsValidTime(GUICtrlRead($midCtrl)) Then $ok = False
		If Not _IsValidTime(GUICtrlRead($wkdCtrl)) Then $ok = False

		$allFilled = $ok
	EndIf

	; Aggregate error messages and color inputs red while invalid; also set tooltips
	Local $aMsgs = ($allFilled ? "" : t("ERROR_FIELDS_REQUIRED"))

	; Clear required tooltip/background for non-empty fields
	For $sKey In $g_FieldCtrls.Keys
		Local $ctrlId2 = $g_FieldCtrls.Item($sKey)
		Local $val2 = StringStripWS(GUICtrlRead($ctrlId2), 3)
		If $val2 <> "" Then
			GUICtrlSetTip($ctrlId2, "")
			GUICtrlSetBkColor($ctrlId2, 0xFFFFFF)
		EndIf
	Next
	Local $sIdVal = StringStripWS(GUICtrlRead($g_FieldCtrls.Item("MeetingID")), 3)
	If $sIdVal <> "" And Not _IsValidMeetingID($sIdVal) Then
		$aMsgs &= ($aMsgs = "" ? t("ERROR_MEETING_ID_FORMAT") : "  •  " & t("ERROR_MEETING_ID_FORMAT"))
		GUICtrlSetColor($g_FieldCtrls.Item("MeetingID"), 0xFF0000)
		GUICtrlSetTip($g_FieldCtrls.Item("MeetingID"), t("ERROR_MEETING_ID_FORMAT"))
	Else
		If $sIdVal <> "" Then
			GUICtrlSetColor($g_FieldCtrls.Item("MeetingID"), 0x000000)
			GUICtrlSetTip($g_FieldCtrls.Item("MeetingID"), "")
		EndIf
	EndIf

	Local $sMidVal = StringStripWS(GUICtrlRead($g_FieldCtrls.Item("MidweekTime")), 3)
	If $sMidVal <> "" And Not _IsValidTime($sMidVal) Then
		$aMsgs &= ($aMsgs = "" ? t("ERROR_TIME_FORMAT") : "  •  " & t("ERROR_TIME_FORMAT"))
		GUICtrlSetColor($g_FieldCtrls.Item("MidweekTime"), 0xFF0000)
		GUICtrlSetTip($g_FieldCtrls.Item("MidweekTime"), t("ERROR_TIME_FORMAT"))
	Else
		If $sMidVal <> "" Then
			GUICtrlSetColor($g_FieldCtrls.Item("MidweekTime"), 0x000000)
			GUICtrlSetTip($g_FieldCtrls.Item("MidweekTime"), "")
		EndIf
	EndIf

	Local $sWkdVal = StringStripWS(GUICtrlRead($g_FieldCtrls.Item("WeekendTime")), 3)
	If $sWkdVal <> "" And Not _IsValidTime($sWkdVal) Then
		$aMsgs &= ($aMsgs = "" ? t("ERROR_TIME_FORMAT") : "  •  " & t("ERROR_TIME_FORMAT"))
		GUICtrlSetColor($g_FieldCtrls.Item("WeekendTime"), 0xFF0000)
		GUICtrlSetTip($g_FieldCtrls.Item("WeekendTime"), t("ERROR_TIME_FORMAT"))
	Else
		If $sWkdVal <> "" Then
			GUICtrlSetColor($g_FieldCtrls.Item("WeekendTime"), 0x000000)
			GUICtrlSetTip($g_FieldCtrls.Item("WeekendTime"), "")
		EndIf
	EndIf

	If $g_ErrorAreaLabel <> 0 Then GUICtrlSetData($g_ErrorAreaLabel, $aMsgs)

	GUICtrlSetState($idSaveBtn, ($allFilled ? $GUI_ENABLE : $GUI_DISABLE))
EndFunc   ;==>CheckConfigFields

; WM_COMMAND handler to catch EN_CHANGE from Edit controls and validate live while typing
Func _WM_COMMAND_EditChange($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
	Local $iCtrlId = BitAND($wParam, 0xFFFF)
	Local $iNotify = BitShift($wParam, 16)
	If $iNotify = $EN_CHANGE Then
		; Trigger validation for any of our text inputs on each keystroke
		For $sKey In $g_FieldCtrls.Keys
			If $iCtrlId = $g_FieldCtrls.Item($sKey) Then
				CheckConfigFields()
				ExitLoop
			EndIf
		Next
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_COMMAND_EditChange

Func SaveConfigGUI()
	$g_UserSettings.RemoveAll()
	; Persist all known fields generically
	For $sKey In $g_FieldCtrls.Keys
		Local $ctrlId = $g_FieldCtrls.Item($sKey)
		Local $val = StringStripWS(GUICtrlRead($ctrlId), 3)
		; Convert day labels to numbers for specific keys
		If ($sKey = "MidweekDay" Or $sKey = "WeekendDay") Then
			If $g_DayLabelToNum.Exists($val) Then $val = String($g_DayLabelToNum.Item($val))
		EndIf
		$g_UserSettings.Add($sKey, $val)
		IniWrite($CONFIG_FILE, _GetIniSectionForKey($sKey), $sKey, GetUserSetting($sKey))
	Next

	; Language from dropdown
	Local $selDisplay = GUICtrlRead($idLanguagePicker)
	Local $selLang = "en"
	If $g_LangNameToCode.Exists($selDisplay) Then $selLang = $g_LangNameToCode.Item($selDisplay)
	$g_UserSettings.Add("Language", $selLang)
	IniWrite($CONFIG_FILE, "General", "Language", GetUserSetting("Language"))
	$g_CurrentLang = $selLang
	_InitDayLabelMaps() ; refresh day labels for new language

	CloseConfigGUI()
EndFunc   ;==>SaveConfigGUI

; Map a setting key to its INI section
Func _GetIniSectionForKey($sKey)
	Switch $sKey
		Case "MeetingID"
			Return "ZoomSettings"
		Case "MidweekDay", "MidweekTime", "WeekendDay", "WeekendTime"
			Return "Meetings"
		Case "Language"
			Return "General"
		Case Else
			Return "ZoomStrings"
	EndSwitch
EndFunc   ;==>_GetIniSectionForKey

Func CloseConfigGUI()
	GUIDelete($g_ConfigGUI)
	$g_ConfigGUI = 0
EndFunc   ;==>CloseConfigGUI

Func QuitApp()
	Exit
EndFunc   ;==>QuitApp

Func LoadMeetingConfig()
	$g_UserSettings.RemoveAll()
	$g_UserSettings.Add("MeetingID", IniRead($CONFIG_FILE, "ZoomSettings", "MeetingID", ""))
	$g_UserSettings.Add("MidweekDay", IniRead($CONFIG_FILE, "Meetings", "MidweekDay", ""))
	$g_UserSettings.Add("MidweekTime", IniRead($CONFIG_FILE, "Meetings", "MidweekTime", ""))
	$g_UserSettings.Add("WeekendDay", IniRead($CONFIG_FILE, "Meetings", "WeekendDay", ""))
	$g_UserSettings.Add("WeekendTime", IniRead($CONFIG_FILE, "Meetings", "WeekendTime", ""))
	$g_UserSettings.Add("HostToolsValue", IniRead($CONFIG_FILE, "ZoomStrings", "HostToolsValue", ""))
	$g_UserSettings.Add("ParticipantValue", IniRead($CONFIG_FILE, "ZoomStrings", "ParticipantValue", ""))
	$g_UserSettings.Add("MuteAllValue", IniRead($CONFIG_FILE, "ZoomStrings", "MuteAllValue", ""))
	$g_UserSettings.Add("YesValue", IniRead($CONFIG_FILE, "ZoomStrings", "YesValue", ""))

	Local $lang = IniRead($CONFIG_FILE, "General", "Language", "")
	If $lang = "" Then
		$lang = "en"
		IniWrite($CONFIG_FILE, "General", "Language", $lang)
	EndIf
	$g_UserSettings.Add("Language", $lang)
	$g_CurrentLang = $lang

	If GetUserSetting("MeetingID") = "" Or GetUserSetting("MidweekDay") = "" Or GetUserSetting("MidweekTime") = "" Or GetUserSetting("WeekendDay") = "" Or GetUserSetting("WeekendTime") = "" Or GetUserSetting("HostToolsValue") = "" Or GetUserSetting("ParticipantValue") = "" Or GetUserSetting("MuteAllValue") = "" Or GetUserSetting("YesValue") = "" Then
		ShowConfigGUI()
		While $g_ConfigGUI
			Sleep(10)
		WEnd
	EndIf
EndFunc   ;==>LoadMeetingConfig

; -------------------------
; Tray icon event handling
; -------------------------
Func TrayEvent()
	UpdateTrayTooltip()
	Local $trayMsg = TrayGetMsg()
	Switch $trayMsg
		Case $TRAY_EVENT_PRIMARYDOWN ; Tray icon left-click
			ShowConfigGUI()
			While $g_ConfigGUI
				Sleep(10)
			WEnd
	EndSwitch
EndFunc   ;==>TrayEvent

_LoadTranslations()
LoadMeetingConfig()
_InitDayLabelMaps()

Func ResponsiveSleep($s)
	Local $elapsed = 0
	Local $ms = $s * 1000
	While $elapsed < $ms
		TrayEvent()
		Sleep(50)
		$elapsed += 50
		; Optionally, handle other GUI events here if needed
	WEnd
EndFunc   ;==>ResponsiveSleep


; -------------------------
; Main loop
; -------------------------
Global $previousRunDay = -1
While True
	TrayEvent()

	Global $today = @WDAY
	If $today <> $previousRunDay Then
		Debug("New day detected. Resetting configuration flags.", "DEBUG")
		$previousRunDay = $today
		$g_PrePostSettingsConfigured = False
		$g_DuringMeetingSettingsConfigured = False
	EndIf

	Global $timeNow = _NowTime(4) ; HH:MM

	If $today = Number(GetUserSetting("MidweekDay")) Then
		CheckMeetingWindow(GetUserSetting("MidweekTime"))
		ResponsiveSleep(5)
	ElseIf $today = Number(GetUserSetting("WeekendDay")) Then
		CheckMeetingWindow(GetUserSetting("WeekendTime"))
		ResponsiveSleep(5)
	Else
		Debug(t("INFO_WAITING"), "INFO") ; Not a meeting day. Sleeping 1 hour.
		ResponsiveSleep(3600)
	EndIf
WEnd
; -------------------------
; Check current time against meeting window
; -------------------------
Func CheckMeetingWindow($meetingTime)
	If $meetingTime = "" Then Return

	Local $aParts = StringSplit($meetingTime, ":")
	Local $hour = Number($aParts[1])
	Local $min = Number($aParts[2])

	Local $nowMin = Number(@HOUR) * 60 + Number(@MIN)
	Local $meetingMin = $hour * 60 + $min

	If $nowMin >= ($meetingMin - 60) And $nowMin < ($meetingMin - 1) Then
		; Pre/Post settings (within 1 hour before meeting, but not 1 minute before)
		If Not $g_PrePostSettingsConfigured Then
			Local $zoomLaunched = _LaunchZoom()
			If Not $zoomLaunched Then
				Debug(t("ERROR_ZOOM_LAUNCH"), "ERROR")
			Else
				Debug(t("INFO_CONFIG_BEFORE_AFTER_START"), "INFO")
				_SetPreAndPostMeetingSettings()
				$g_PrePostSettingsConfigured = True
				Debug(t("INFO_CONFIG_BEFORE_AFTER_DONE"), "INFO")
			EndIf
		EndIf

	ElseIf $nowMin = ($meetingMin - 1) Then
		; 1 minute before meeting → During meeting settings
		If Not $g_DuringMeetingSettingsConfigured Then
			Local $zoomExists = IsObj(_GetZoomWindow())
			If Not $zoomExists Then
				Debug(t("ERROR_ZOOM_WINDOW_NOT_FOUND"), "ERROR")
			Else
				Debug(t("INFO_MEETING_STARTING_SOON_CONFIG"), "INFO")
				_SetDuringMeetingSettings()
				$g_DuringMeetingSettingsConfigured = True
				Debug(t("INFO_CONFIG_DURING_MEETING_DONE"), "INFO")
			EndIf
		EndIf

	ElseIf $nowMin >= $meetingMin Then
		; Meeting already started
		Local $minutesAgo = $nowMin - $meetingMin
		If $minutesAgo <= 120 Then
			Debug(t("INFO_MEETING_STARTED_AGO", $minutesAgo), "INFO", True)
		Else
			Debug(t("INFO_OUTSIDE_MEETING_WINDOW"), "INFO", True)
		EndIf

	Else
		; Too early - show countdown
		Local $minutesLeft = $meetingMin - $nowMin
		Debug(t("INFO_MEETING_STARTING_IN", $minutesLeft), "INFO", True)
	EndIf
EndFunc   ;==>CheckMeetingWindow

; -------------------------
; Launch Zoom
; -------------------------
Func _LaunchZoom()
	Debug(t("INFO_ZOOM_LAUNCHING"), "INFO")

	Local $meetingID = GetUserSetting("MeetingID")
	If $meetingID = "" Then
		Debug(t("ERROR_MEETING_ID_NOT_CONFIGURED"), "ERROR")
		Return SetError(1, 0, 0)
	EndIf

	Local $zoomURL = "zoommtg://zoom.us/join?confno=" & $meetingID
	ShellExecute($zoomURL)
	Debug(t("INFO_ZOOM_LAUNCHED") & ": " & $meetingID, "INFO")
	ResponsiveSleep(10)
	Return IsObj(_GetZoomWindow())
EndFunc   ;==>_LaunchZoom

Func FindElementByClassName($sClassName, $iScope = Default, $oParent = Default)
	Debug("Searching for element with class: '" & $sClassName & "'", "DEBUG")

	; Use desktop as default parent if not specified
	Local $oSearchParent = $oDesktop
	If $oParent <> Default Then $oSearchParent = $oParent

	; Create condition for class name
	Local $pClassCondition
	$oUIAutomation.CreatePropertyCondition($UIA_ClassNamePropertyId, $sClassName, $pClassCondition)

	; Find the element
	Local $pElement
	Local $scope = $TreeScope_Descendants
	If $iScope <> Default Then $scope = $iScope

	$oSearchParent.FindFirst($scope, $pClassCondition, $pElement)

	Local $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
	If Not IsObj($oElement) Then
		Debug("Element with class '" & $sClassName & "' not found.", "WARN")
		Return 0
	EndIf

	Debug("Element with class '" & $sClassName & "' found.", "DEBUG")
	Return $oElement
EndFunc   ;==>FindElementByClassName

; Now you can use it with the original scope if needed:
Func _GetZoomWindow()
	$oZoomWindow = FindElementByClassName("ConfMultiTabContentWndClass", $TreeScope_Children)
	If Not IsObj($oZoomWindow) Then Return SetError(1, 0, 0)
	Debug("Zoom window obtained.", "UIA")
	Return $oZoomWindow
EndFunc   ;==>_GetZoomWindow

Func FindElementByPartialName($sPartial, $aControlTypes = Default, $oParent = Default)
	Debug("Searching for element containing: '" & $sPartial & "'", "DEBUG")

	; Use desktop as default parent if not specified
	Local $oSearchParent = $oDesktop
	If $oParent <> Default Then
		$oSearchParent = $oParent
		Debug("Using custom parent element for search", "DEBUG")
	Else
		Debug("Using desktop as search parent", "DEBUG")
	EndIf

	; Default to button and menu item if not specified
	If $aControlTypes = Default Then
		Local $aDefaultTypes[2] = [$UIA_ButtonControlTypeId, $UIA_MenuItemControlTypeId]
		$aControlTypes = $aDefaultTypes
	EndIf

	; Search through each control type
	For $iType = 0 To UBound($aControlTypes) - 1
		Local $iControlType = $aControlTypes[$iType]
		Debug("Searching control type: " & $iControlType)

		; Create condition for this control type
		Local $pCondition
		$oUIAutomation.CreatePropertyCondition($UIA_ControlTypePropertyId, $iControlType, $pCondition)

		; Find all elements of this type under the specified parent
		Local $pElements
		$oSearchParent.FindAll($TreeScope_Descendants, $pCondition, $pElements)

		Local $oElements = ObjCreateInterface($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
		If IsObj($oElements) Then
			Local $iCount
			$oElements.Length($iCount)
			Debug("Found " & $iCount & " elements of this type.", "DEBUG")

			; Loop through each element
			For $i = 0 To $iCount - 1
				Local $pElement
				$oElements.GetElement($i, $pElement)

				Local $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
				If IsObj($oElement) Then
					; Get the name property
					Local $sName
					$oElement.GetCurrentPropertyValue($UIA_NamePropertyId, $sName)

					; Check if the partial string is contained in the element name
					If StringInStr($sName, $sPartial, $STR_NOCASESENSEBASIC) > 0 Then
						Debug("Element found with name: '" & $sName & "'", "DEBUG")
						Return $oElement
					EndIf
				EndIf
			Next
		EndIf
	Next

	Debug("No element found containing: '" & $sPartial & "'", "WARN")
	Return 0
EndFunc   ;==>FindElementByPartialName

Func _ClickElement($oElement, $ForceClick = False)
	ResponsiveSleep(0.5) ; brief pause to ensure UI is ready
	If Not IsObj($oElement) Then
		Debug(t("ERROR_INVALID_ELEMENT_OBJECT"), "WARN")
		Return False
	EndIf

	; Get element name for debugging
	Local $sElementName = ""
	$oElement.GetCurrentPropertyValue($UIA_NamePropertyId, $sElementName)
	Debug("Attempting to click element: '" & $sElementName & "'", "DEBUG")

	; Method 0: Try Mouse Click (only if forced)
	If $ForceClick Then
		Debug("Forcing click.", "DEBUG")
		UIA_MouseClick($oElement)
		Debug("Element clicked via Mouse Click.", "SUCCESS")
		Return True
	Else
		Debug("Skipping Method 0 (Mouse Click) for non-menu item.", "DEBUG")
	EndIf


	; Method 1: Try Invoke Pattern (works for most buttons)
	Local $pInvokePattern, $oInvokePattern
	$oElement.GetCurrentPattern($UIA_InvokePatternId, $pInvokePattern)
	If $pInvokePattern Then
		$oInvokePattern = ObjCreateInterface($pInvokePattern, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern)
		If IsObj($oInvokePattern) Then
			Debug("Using Invoke pattern for: '" & $sElementName & "'", "DEBUG")
			$oInvokePattern.Invoke()
			Debug("Element clicked via Invoke pattern.", "SUCCESS")
			Return True
		EndIf
	EndIf

	; Method 2: Try Legacy Accessible Pattern (works for menu items and some older controls)
	Local $pLegacyPattern, $oLegacyPattern
	$oElement.GetCurrentPattern($UIA_LegacyIAccessiblePatternId, $pLegacyPattern)
	If $pLegacyPattern Then
		$oLegacyPattern = ObjCreateInterface($pLegacyPattern, $sIID_IUIAutomationLegacyIAccessiblePattern, $dtagIUIAutomationLegacyIAccessiblePattern)
		If IsObj($oLegacyPattern) Then
			Debug("Using LegacyAccessible pattern for: '" & $sElementName & "'", "DEBUG")
			$oLegacyPattern.DoDefaultAction()
			Debug("Element clicked via LegacyAccessible pattern.", "SUCCESS")
			Return True
		EndIf
	EndIf

	; Method 3: Try Selection Item Pattern (for selectable menu items)
	Local $pSelectionItemPattern, $oSelectionItemPattern
	$oElement.GetCurrentPattern($UIA_SelectionItemPatternId, $pSelectionItemPattern)
	If $pSelectionItemPattern Then
		$oSelectionItemPattern = ObjCreateInterface($pSelectionItemPattern, $sIID_IUIAutomationSelectionItemPattern, $dtagIUIAutomationSelectionItemPattern)
		If IsObj($oSelectionItemPattern) Then
			Debug("Using SelectionItem pattern for: '" & $sElementName & "'", "DEBUG")
			$oSelectionItemPattern.Select()
			Debug("Element selected via SelectionItem pattern.", "SUCCESS")
			Return True
		EndIf
	EndIf

	; Method 4: Try Toggle Pattern (for toggle buttons/menu items)
	Local $pTogglePattern, $oTogglePattern
	$oElement.GetCurrentPattern($UIA_TogglePatternId, $pTogglePattern)
	If $pTogglePattern Then
		$oTogglePattern = ObjCreateInterface($pTogglePattern, $sIID_IUIAutomationTogglePattern, $dtagIUIAutomationTogglePattern)
		If IsObj($oTogglePattern) Then
			Debug("Using Toggle pattern for: '" & $sElementName & "'", "DEBUG")
			$oTogglePattern.Toggle()
			Debug("Element toggled via Toggle pattern.", "SUCCESS")
			Return True
		EndIf
	EndIf

	; Method 5: Fallback - Mouse click at element center
	Local $tRect
	$oElement.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $tRect)
	If $tRect Then
		; Parse the bounding rectangle (format: "left,top,width,height")
		Local $aRect = StringSplit($tRect, ",")
		If $aRect[0] >= 4 Then
			Local $iLeft = Number($aRect[1])
			Local $iTop = Number($aRect[2])
			Local $iWidth = Number($aRect[3])
			Local $iHeight = Number($aRect[4])

			Local $iCenterX = $iLeft + ($iWidth / 2)
			Local $iCenterY = $iTop + ($iHeight / 2)

			Debug("Using mouse click fallback at position: " & $iCenterX & "," & $iCenterY & " for: '" & $sElementName & "'", "DEBUG")

			; Ensure element is visible/enabled before clicking
			Local $bIsEnabled, $bIsOffscreen
			$oElement.GetCurrentPropertyValue($UIA_IsEnabledPropertyId, $bIsEnabled)
			$oElement.GetCurrentPropertyValue($UIA_IsOffscreenPropertyId, $bIsOffscreen)

			If $bIsEnabled And Not $bIsOffscreen Then
				MouseClick("left", $iCenterX, $iCenterY, 1, 0)
				Debug("Element clicked via mouse at center.", "SUCCESS")
				Return True
			Else
				Debug("Element not clickable - Enabled: " & $bIsEnabled & ", Offscreen: " & $bIsOffscreen, "WARN")
			EndIf
		EndIf
	EndIf

	; All methods failed
	Debug(t("ERROR_FAILED_CLICK_ELEMENT") & ": '" & $sElementName & "'", "ERROR")
	Return False
EndFunc   ;==>_ClickElement

Func _OpenHostTools()
	If Not IsObj($oZoomWindow) Then Return False
	Local $oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
	If Not IsObj($oHostMenu) Then
		Local $oButton = FindElementByPartialName(GetUserSetting("HostToolsValue"), Default, $oZoomWindow)
		If Not _ClickElement($oButton) Then
			Return False
		EndIf
	EndIf
	$oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
	Return $oHostMenu
EndFunc   ;==>_OpenHostTools

Func _CloseHostTools()
	; Click on the main window to hide the open menu
	_ClickElement($oZoomWindow, True)
	Debug("Host tools menu closed.", "UIA")
EndFunc   ;==>_CloseHostTools

Func _OpenParticipantsPanel()
	If Not IsObj($oZoomWindow) Then Return False
	Local $ListType[1] = [$UIA_ListControlTypeId]
	Local $oParticipantsPanel = FindElementByPartialName(GetUserSetting("ParticipantValue"), $ListType, $oZoomWindow)

	If Not IsObj($oParticipantsPanel) Then
		Local $oButton = FindElementByPartialName(GetUserSetting("ParticipantValue"), Default, $oZoomWindow)
		If Not _ClickElement($oButton) Then
			Debug("Failed to click Participants.", "ERROR")
			Return False
		EndIf
	EndIf
	$oParticipantsPanel = FindElementByPartialName(GetUserSetting("ParticipantValue"), $ListType, $oZoomWindow)
	If IsObj($oParticipantsPanel) Then
		Debug("Participants panel opened.", "UIA")
	Else
		Debug("Failed to open Participants panel.", "ERROR")
	EndIf
	Return $oParticipantsPanel
EndFunc   ;==>_OpenParticipantsPanel

Func GetSecuritySettingsState($aSettings)
	_OpenHostTools()

	; Create a dictionary object
	Local $oDict = ObjCreate("Scripting.Dictionary")

	For $i = 0 To UBound($aSettings) - 1
		Local $sSetting = $aSettings[$i]

		; Find menu item by partial name
		Local $oSetting = FindElementByPartialName($sSetting, Default, $oZoomWindow)
		Local $bEnabled = False
		If IsObj($oSetting) Then
			; Read the Name property
			Local $sLabel = ""
			$oSetting.GetCurrentPropertyValue($UIA_NamePropertyId, $sLabel)
			Local $sLabelLower = StringLower($sLabel)

			; Check for "unchecked" as whole word
			$bEnabled = Not StringRegExp($sLabelLower, "\bunchecked\b")
		Else
			Debug(t("ERROR_SETTING_NOT_FOUND") & ": '" & $sSetting & "'", "ERROR")
		EndIf

		Debug("Setting '" & $sSetting & "' is currently " & ($bEnabled ? "ENABLED" : "DISABLED"), "DEBUG")

		; Add to dictionary
		$oDict.Add($sSetting, $bEnabled)
	Next

	_CloseHostTools()

	Return $oDict
EndFunc   ;==>GetSecuritySettingsState

Func SetSecuritySetting($sSetting, $bDesired)
	Local $oHostMenu = _OpenHostTools()
	If Not IsObj($oHostMenu) Then Return False

	Local $oSetting = FindElementByPartialName($sSetting, Default, $oHostMenu)
	If Not IsObj($oSetting) Then Return

	Local $sLabel
	$oSetting.GetCurrentPropertyValue($UIA_NamePropertyId, $sLabel)

	Local $sLabelLower = StringLower($sLabel)
	Local $bEnabled = Not StringRegExp($sLabelLower, "\bunchecked\b")

	Debug("Setting '" & $sLabel & "' | Current: " & $bEnabled & " | Desired: " & $bDesired, "DEBUG")

	If $bEnabled <> $bDesired Then
		_ClickElement($oSetting, True)
		Debug("Toggled setting '" & $sSetting & "'", "SETTING CHANGE")
	Else
		_CloseHostTools()
	EndIf
EndFunc   ;==>SetSecuritySetting

Func ToggleFeed($feedType, $desiredState)
	Local $currentlyEnabled = False

	If $feedType = "Video" Then
		Local $stopMyVideoButton = FindElementByPartialName("Stop my video", Default, $oZoomWindow) ; Present if video currently enabled
		Local $startMyVideoButton = FindElementByPartialName("Start my video", Default, $oZoomWindow) ; Present if video currently disabled
		$currentlyEnabled = IsObj($stopMyVideoButton)

		If $desiredState <> $currentlyEnabled Then
			If IsObj($stopMyVideoButton) Then
				_ClickElement($stopMyVideoButton)
			ElseIf IsObj($startMyVideoButton) Then
				_ClickElement($startMyVideoButton)
			Else
				Debug("No video button found to toggle!", "WARN")
			EndIf
		EndIf

	ElseIf $feedType = "Audio" Then
		Local $muteHostButton = FindElementByPartialName("currently unmuted", Default, $oZoomWindow) ; Present if audio currently enabled
		Local $unmuteHostButton = FindElementByPartialName("Unmute my audio", Default, $oZoomWindow) ; Present if audio currently disabled
		$currentlyEnabled = IsObj($muteHostButton)

		If $desiredState <> $currentlyEnabled Then
			If IsObj($muteHostButton) Then
				_ClickElement($muteHostButton)
			ElseIf IsObj($unmuteHostButton) Then
				_ClickElement($unmuteHostButton)
			Else
				Debug("No audio button found to toggle!", "WARN")
			EndIf
		EndIf
	Else
		Debug(t("ERROR_UNKNOWN_FEED_TYPE") & ": '" & $feedType & "'", "WARN")
	EndIf
	ResponsiveSleep(1)
EndFunc   ;==>ToggleFeed


Func _SetPreAndPostMeetingSettings()
	SetSecuritySetting("Unmute", True)
	SetSecuritySetting("Share screen", False)
	ToggleFeed("Audio", False)
	ToggleFeed("Video", False)
EndFunc   ;==>_SetPreAndPostMeetingSettings

Func _SetDuringMeetingSettings()
	SetSecuritySetting("Unmute", False)
	SetSecuritySetting("Share screen", False)
	MuteAll()
	ToggleFeed("Audio", True)
	ToggleFeed("Video", True)
EndFunc   ;==>_SetDuringMeetingSettings

Func MuteAll()
	Debug("Attempting to 'Mute All' participants.")
	Local $oParticipantsPanel = _OpenParticipantsPanel()
	If Not IsObj($oParticipantsPanel) Then Return False
	Local $oButton = FindElementByPartialName(GetUserSetting("MuteAllValue"), Default, $oZoomWindow)
	If Not _ClickElement($oButton) Then
		Debug("'Mute all' button not found.", "WARN")
		Return False
	EndIf

	Return DialogClick("zChangeNameWndClass", GetUserSetting("YesValue"))
EndFunc   ;==>MuteAll

Func DialogClick($ClassName, $ButtonLabel)
	; Confirm dialog
	Local $oDialog = FindElementByClassName($ClassName)
	Local $oButton = FindElementByPartialName($ButtonLabel, Default, $oDialog)
	If _ClickElement($oButton) Then Return True
	Debug("Button '" & $ButtonLabel & "' not found.", "WARN")
	Return False

EndFunc   ;==>DialogClick

Func _GetZoomStatus()
	Local $aSettings[2] = ["Unmute themselves", "Share screen"]
	Local $oStates = GetSecuritySettingsState($aSettings)

	For $sKey In $oStates.Keys
		Debug($sKey & " = " & ($oStates.Item($sKey) ? "ENABLED" : "DISABLED"), "ZOOM STATUS")
	Next
EndFunc   ;==>_GetZoomStatus

