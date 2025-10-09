; ================================================================================================
; ZOOMMATE - Automated Zoom Meeting Management Tool
; ================================================================================================
; This script automatically manages Zoom meeting settings based on scheduled meeting times.
; It configures security settings before/after meetings and applies meeting-specific settings
; when meetings start.

; ================================================================================================
; COMPILER DIRECTIVES AND INCLUDES
; ================================================================================================
#AutoIt3Wrapper_UseX64=y
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <Date.au3>
#include <StringConstants.au3>
#include <TrayConstants.au3>
#include <GuiMenu.au3>
#include <GUIConstantsEx.au3>
#include <WindowsStylesConstants.au3>
#include <StaticConstants.au3>
#include <WinAPI.au3>
#include <EditConstants.au3>
#include "Includes\UIA_Functions-a.au3"
#include "Includes\CUIAutomation2.au3"

; ================================================================================================
; AUTOIT OPTIONS AND CONSTANTS
; ================================================================================================
; Set AutoIt options for better script behavior
Opt("MustDeclareVars", 1)        ; Force variable declarations
Opt("GUIOnEventMode", 1)         ; Enable GUI event mode
Opt("TrayMenuMode", 3)           ; Custom tray menu (no default, no auto-pause)

; Windows API constants for GUI message handling
Global Const $MF_BYCOMMAND = 0x00000000

; ================================================================================================
; CONFIGURATION AND GLOBAL VARIABLES
; ================================================================================================
; Configuration file path
Global Const $CONFIG_FILE = @ScriptDir & "\zoom_config.ini"

; Meeting automation state flags
Global $g_PrePostSettingsConfigured = False    ; Tracks if pre/post meeting settings are applied
Global $g_DuringMeetingSettingsConfigured = False ; Tracks if during-meeting settings are applied

; Notification control
Global $g_InitialNotificationWasShown = False ; Prevents repeated initial notifications

; User settings storage (Dictionary object for key-value pairs)
Global $g_UserSettings = ObjCreate("Scripting.Dictionary")

; Internationalization (i18n) containers
Global $g_Languages = ObjCreate("Scripting.Dictionary")        ; langCode -> translation dictionary
Global $g_LangCodeToName = ObjCreate("Scripting.Dictionary")   ; langCode -> display name
Global $g_LangNameToCode = ObjCreate("Scripting.Dictionary")   ; display name -> langCode
Global $g_CurrentLang = "en"                                   ; Current language setting

; GUI control references
Global $idSaveBtn                              ; Save button control ID
Global $idLanguagePicker                       ; Language dropdown control ID
Global $g_FieldCtrls = ObjCreate("Scripting.Dictionary")  ; Maps field names to control IDs
Global $g_ErrorAreaLabel = 0                   ; Error display label control ID
Global $g_ConfigGUI = 0                        ; Configuration GUI handle
Global $g_OverlayMessageGUI = 0                    ; Handle for the please-wait popup
Global $g_TooltipGUI = 0                       ; Handle for custom image tooltip
Global $g_InfoIconData = ObjCreate("Scripting.Dictionary")  ; Maps info icon IDs to image paths
Global $g_ElementNamesGUI = 0                      ; Handle for element names display GUI
Global $g_ElementNamesEdit = 0                     ; Handle for element names edit control

; Day mapping containers for internationalization
Global $g_DayLabelToNum = ObjCreate("Scripting.Dictionary")    ; Day name -> number (1-7)
Global $g_DayNumToLabel = ObjCreate("Scripting.Dictionary")    ; Day number -> name

; Status and tray icon variables
Global $g_StatusMsg = "Idle"                   ; Current status message
Global $g_TrayIcon = @ScriptDir & "\zoommate.ico" ; Tray icon path

; UIAutomation COM objects (for Zoom window interaction)
Global $oUIAutomation                          ; Main UIAutomation interface
Global $pDesktop                               ; Desktop element pointer
Global $oDesktop                               ; Desktop element object
Global $oZoomWindow = 0                        ; Zoom window element

; Meeting timing control
Global $previousRunDay = -1                    ; Tracks day changes for state reset

; Keyboard shortcut to trigger post-meeting settings
Global $g_KeyboardShortcut = ""               ; Current keyboard shortcut (e.g., "^!z")
Global $g_HotkeyRegistered = False             ; Tracks if hotkey is currently registered

; ================================================================================================
; USER SETTINGS MANAGEMENT
; ================================================================================================

; Retrieves a user setting value by key
; @param $key - The setting key to retrieve
; @return String - The setting value or empty string if not found
Func GetUserSetting($key)
	If $g_UserSettings.Exists($key) Then Return $g_UserSettings.Item($key)
	Return ""
EndFunc   ;==>GetUserSetting

; ================================================================================================
; INTERNATIONALIZATION (I18N) FUNCTIONS
; ================================================================================================

; Translation lookup function with placeholder support
; @param $key - Translation key to look up
; @param $p0-$p2 - Optional placeholder values for {0}, {1}, {2} substitution
; @return String - Translated text with placeholders replaced
Func t($key, $p0 = Default, $p1 = Default, $p2 = Default)
	; Try current language first
	If $g_Languages.Exists($g_CurrentLang) Then
		Local $oDict = $g_Languages.Item($g_CurrentLang)
		If $oDict.Exists($key) Then
			Local $s = $oDict.Item($key)
			; Replace placeholders if provided
			If $p0 <> Default Then $s = StringReplace($s, "{0}", $p0)
			If $p1 <> Default Then $s = StringReplace($s, "{1}", $p1)
			If $p2 <> Default Then $s = StringReplace($s, "{2}", $p2)
			Return $s
		EndIf
	EndIf

	; Fallback to English if current language doesn't have the key
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

	; Ultimate fallback: return the key itself
	Return $key
EndFunc   ;==>t

; Loads all translation files from the i18n directory
; Scans for *.ini files and builds language dictionaries
Func _LoadTranslations()
	Local $sDir = @ScriptDir & "\i18n\*.ini"
	Local $hSearch = FileFindFirstFile($sDir)
	If $hSearch = -1 Then Return ; No translation files found

	While 1
		Local $sFile = FileFindNextFile($hSearch)
		If @error Then ExitLoop

		; Extract language code from filename (remove .ini extension)
		Local $lang = StringTrimRight($sFile, 4)
		Local $fullPath = @ScriptDir & "\i18n\" & $sFile

		; Read all translations from the [translations] section
		Local $a = IniReadSection($fullPath, "translations")
		If @error Then ContinueLoop

		; Build translation dictionary for this language
		Local $dict = ObjCreate("Scripting.Dictionary")
		For $i = 1 To $a[0][0]
			$dict.Add($a[$i][0], $a[$i][1])
		Next

		; Store language dictionary
		If $g_Languages.Exists($lang) Then $g_Languages.Remove($lang)
		$g_Languages.Add($lang, $dict)

		; Build language name mappings for GUI dropdown
		Local $langName = ""
		If $dict.Exists("LANGNAME") Then $langName = $dict.Item("LANGNAME")
		If $langName = "" Then $langName = $lang ; Use code as fallback

		If $g_LangCodeToName.Exists($lang) Then $g_LangCodeToName.Remove($lang)
		$g_LangCodeToName.Add($lang, $langName)
		If $g_LangNameToCode.Exists($langName) Then $g_LangNameToCode.Remove($langName)
		$g_LangNameToCode.Add($langName, $lang)
	WEnd
	FileClose($hSearch)
EndFunc   ;==>_LoadTranslations

; Builds a comma-separated list of available language display names
; @return String - Comma-separated list of language names
Func _ListAvailableLanguageNames()
	Local $list = ""
	For $name In $g_LangNameToCode.Keys
		$list &= ($list = "" ? $name : "," & $name)
	Next
	Return $list
EndFunc   ;==>_ListAvailableLanguageNames

; Gets the display name for a language code
; @param $code - Language code (e.g., "en", "es")
; @return String - Display name or the code itself if not found
Func _GetLanguageDisplayName($code)
	If $g_LangCodeToName.Exists($code) Then Return $g_LangCodeToName.Item($code)
	Return $code
EndFunc   ;==>_GetLanguageDisplayName

; Initializes day name to number mappings using translations
; Maps localized day names (DAY_1 through DAY_7) to numbers 1-7
Func _InitDayLabelMaps()
	Local $i
	For $i = 1 To 7
		Local $key = "DAY_" & $i
		Local $label = t($key)
		If Not $g_DayLabelToNum.Exists($label) Then $g_DayLabelToNum.Add($label, $i)
		If Not $g_DayNumToLabel.Exists(String($i)) Then $g_DayNumToLabel.Add(String($i), $label)
	Next
EndFunc   ;==>_InitDayLabelMaps

; ================================================================================================
; DEBUG AND STATUS FUNCTIONS
; ================================================================================================

; Debug logging and status update function
; @param $string - Message to log/display
; @param $type - Message type (DEBUG, INFO, ERROR, etc.)
; @param $noNotify - If True, suppress tray notifications
Func Debug($string, $type = "DEBUG", $noNotify = False)
	If ($string) Then
		; Always log to console
		ConsoleWrite("[" & $type & "] " & $string & @CRLF)

		; Update status and show tray notification for important messages
		If $type = "INFO" Or $type = "ERROR" Then
			$g_StatusMsg = $string
			If Not $noNotify Then
				TrayTip("ZoomMate", $string, 5, ($type = "INFO" ? 1 : 3))
			EndIf
		EndIf

		; Change tray icon for errors
		If $type = "ERROR" Then
			TraySetIcon($g_TrayIcon, 1) ; Error icon
		EndIf
	EndIf
EndFunc   ;==>Debug

; Updates the tray icon tooltip with current status
Func UpdateTrayTooltip()
	TraySetToolTip("ZoomMate: " & $g_StatusMsg)
EndFunc   ;==>UpdateTrayTooltip

; ================================================================================================
; CONFIGURATION LOADING AND SAVING
; ================================================================================================

; Loads meeting configuration from INI file
; If any required settings are missing, opens the configuration GUI
Func LoadMeetingConfig()
	; Clear existing settings
	$g_UserSettings.RemoveAll()

	; Load all required settings from INI file
	$g_UserSettings.Add("MeetingID", IniRead($CONFIG_FILE, "ZoomSettings", "MeetingID", ""))
	$g_UserSettings.Add("MidweekDay", IniRead($CONFIG_FILE, "Meetings", "MidweekDay", ""))
	$g_UserSettings.Add("MidweekTime", IniRead($CONFIG_FILE, "Meetings", "MidweekTime", ""))
	$g_UserSettings.Add("WeekendDay", IniRead($CONFIG_FILE, "Meetings", "WeekendDay", ""))
	$g_UserSettings.Add("WeekendTime", IniRead($CONFIG_FILE, "Meetings", "WeekendTime", ""))
	$g_UserSettings.Add("HostToolsValue", IniRead($CONFIG_FILE, "ZoomStrings", "HostToolsValue", ""))
	$g_UserSettings.Add("ParticipantValue", IniRead($CONFIG_FILE, "ZoomStrings", "ParticipantValue", ""))
	$g_UserSettings.Add("MuteAllValue", IniRead($CONFIG_FILE, "ZoomStrings", "MuteAllValue", ""))
	$g_UserSettings.Add("MoreMeetingControlsValue", IniRead($CONFIG_FILE, "ZoomStrings", "MoreMeetingControlsValue", ""))
	$g_UserSettings.Add("YesValue", IniRead($CONFIG_FILE, "ZoomStrings", "YesValue", ""))
	$g_UserSettings.Add("UncheckedValue", IniRead($CONFIG_FILE, "ZoomStrings", "UncheckedValue", ""))
	$g_UserSettings.Add("CurrentlyUnmutedValue", IniRead($CONFIG_FILE, "ZoomStrings", "CurrentlyUnmutedValue", ""))
	$g_UserSettings.Add("UnmuteAudioValue", IniRead($CONFIG_FILE, "ZoomStrings", "UnmuteAudioValue", ""))
	$g_UserSettings.Add("StopVideoValue", IniRead($CONFIG_FILE, "ZoomStrings", "StopVideoValue", ""))
	$g_UserSettings.Add("StartVideoValue", IniRead($CONFIG_FILE, "ZoomStrings", "StartVideoValue", ""))
	$g_UserSettings.Add("KeyboardShortcut", IniRead($CONFIG_FILE, "General", "KeyboardShortcut", ""))

	; Window snapping preference (Disabled|Left|Right)
	$g_UserSettings.Add("SnapZoomSide", IniRead($CONFIG_FILE, "General", "SnapZoomSide", "Disabled"))

	; Load language setting
	Local $lang = IniRead($CONFIG_FILE, "General", "Language", "")
	If $lang = "" Then
		$lang = "en"
		IniWrite($CONFIG_FILE, "General", "Language", $lang)
	EndIf
	$g_UserSettings.Add("Language", $lang)
	$g_CurrentLang = $lang

	; Load keyboard shortcut setting
	$g_KeyboardShortcut = GetUserSetting("KeyboardShortcut")
	If $g_KeyboardShortcut <> "" Then
		_UpdateKeyboardShortcut()
	EndIf

	; Check if all required settings are configured
	If GetUserSetting("MeetingID") = "" Or GetUserSetting("MidweekDay") = "" Or GetUserSetting("MidweekTime") = "" Or GetUserSetting("WeekendDay") = "" Or GetUserSetting("WeekendTime") = "" Or GetUserSetting("HostToolsValue") = "" Or GetUserSetting("ParticipantValue") = "" Or GetUserSetting("MuteAllValue") = "" Or GetUserSetting("YesValue") = "" Or GetUserSetting("UncheckedValue") = "" Or GetUserSetting("CurrentlyUnmutedValue") = "" Or GetUserSetting("UnmuteAudioValue") = "" Or GetUserSetting("StopVideoValue") = "" Or GetUserSetting("StartVideoValue") = "" Then
		; Open configuration GUI if any settings are missing
		ShowConfigGUI()
		While $g_ConfigGUI
			Sleep(100)
		WEnd
	Else
		Debug(t("INFO_CONFIG_LOADED"), "INFO")
		; print meeting schedule for verification
		Debug("Midweek Meeting: " & t("DAY_" & GetUserSetting("MidweekDay")) & " at " & GetUserSetting("MidweekTime"), "INFO", True)
		Debug("Weekend Meeting: " & t("DAY_" & GetUserSetting("WeekendDay")) & " at " & GetUserSetting("WeekendTime"), "INFO", True)
	EndIf
EndFunc   ;==>LoadMeetingConfig

; Maps a setting key to its appropriate INI file section
; @param $sKey - Setting key name
; @return String - INI section name
Func _GetIniSectionForKey($sKey)
	Switch $sKey
		Case "MeetingID"
			Return "ZoomSettings"
		Case "MidweekDay", "MidweekTime", "WeekendDay", "WeekendTime"
			Return "Meetings"
		Case "Language", "SnapZoomSide", "KeyboardShortcut"
			Return "General"
		Case Else
			Return "ZoomStrings"
	EndSwitch
EndFunc   ;==>_GetIniSectionForKey

; ================================================================================================
; INPUT VALIDATION FUNCTIONS
; ================================================================================================

; Validates meeting ID format (9-11 digits)
; @param $s - String to validate
; @return Boolean - True if valid meeting ID format
Func _IsValidMeetingID($s)
	$s = StringStripWS($s, 3)  ; Remove leading/trailing whitespace
	If $s = "" Then Return False
	If Not StringRegExp($s, "^\d{9,11}$") Then Return False
	Return True
EndFunc   ;==>_IsValidMeetingID

; Validates time format (HH:MM in 24-hour format)
; @param $s - String to validate
; @return Boolean - True if valid time format
Func _IsValidTime($s)
	$s = StringStripWS($s, 3)  ; Remove leading/trailing whitespace
	If $s = "" Then Return False
	If Not StringRegExp($s, "^(\d{1,2}):(\d{2})$") Then Return False

	; Validate hour and minute ranges
	Local $a = StringSplit($s, ":")
	Local $h = Number($a[1])
	Local $m = Number($a[2])
	If $h < 0 Or $h > 23 Then Return False
	If $m < 0 Or $m > 59 Then Return False
	Return True
EndFunc   ;==>_IsValidTime

; Validates keyboard shortcut format
; @param $s - String to validate
; @return Boolean - True if valid keyboard shortcut format
Func _IsValidKeyboardShortcut($s)
	$s = StringStripWS($s, 3)  ; Remove leading/trailing whitespace
	If $s = "" Then Return True  ; Empty shortcut is valid (means no hotkey)

	; Basic validation for AutoIt hotkey format: modifiers + key
	; Valid modifiers: ^ (Ctrl), ! (Alt), + (Shift), # (Win)
	; Valid keys: a-z, A-Z, 0-9, F1-F12, etc.
	If Not StringRegExp($s, "^[\^\!\+\#]*[a-zA-Z0-9]$") Then Return False

	; Must have at least one modifier key
	If Not StringRegExp($s, "[\^\!\+\#]") Then Return False

	Return True
EndFunc   ;==>_IsValidKeyboardShortcut

; ================================================================================================
; CONFIGURATION GUI FUNCTIONS
; ================================================================================================

; Shows the configuration GUI for user to input settings
Func ShowConfigGUI()
	; If GUI already exists, just show it
	If $g_ConfigGUI Then
		GUICtrlSetState($g_ConfigGUI, @SW_SHOW)
		Return
	EndIf

	; Initialize day mappings for current language
	_InitDayLabelMaps()

	; Create main configuration window
	$g_ConfigGUI = GUICreate(t("CONFIG_TITLE"), 320, 600)
	GUISetOnEvent($GUI_EVENT_CLOSE, "SaveConfigGUI", $g_ConfigGUI)

	; Language selection dropdown
	GUICtrlCreateLabel("Language:", 10, 160, 120, 20)
	$idLanguagePicker = GUICtrlCreateCombo("", 140, 160, 160, 20)
	GUICtrlSetData($idLanguagePicker, _ListAvailableLanguageNames(), _GetLanguageDisplayName(GetUserSetting("Language")))

	; Snap Zoom to side (Disabled|Left|Right)
	GUICtrlCreateLabel(t("LABEL_SNAP_ZOOM_TO"), 10, 190, 120, 20)
	Local $idSnapZoom = GUICtrlCreateCombo("", 140, 190, 160, 20)
	Local $snapVal = GetUserSetting("SnapZoomSide")
	Local $snapDisplay = t("SNAP_DISABLED")
	If $snapVal = "Left" Then
		$snapDisplay = t("SNAP_LEFT")
	ElseIf $snapVal = "Right" Then
		$snapDisplay = t("SNAP_RIGHT")
	EndIf
	GUICtrlSetData($idSnapZoom, t("SNAP_DISABLED") & "|" & t("SNAP_LEFT") & "|" & t("SNAP_RIGHT"), $snapDisplay)
	If Not $g_FieldCtrls.Exists("SnapZoomSide") Then $g_FieldCtrls.Add("SnapZoomSide", $idSnapZoom)
	GUICtrlSetOnEvent($idSnapZoom, "CheckConfigFields")

	; Meeting configuration fields
	_AddTextInputField("MeetingID", t("LABEL_MEETING_ID"), 10, 10, 170, 10, 160)

	; Midweek meeting settings
	_AddDayDropdownField("MidweekDay", t("LABEL_MIDWEEK_DAY"), 10, 40, 200, 40, 100)
	_AddTextInputField("MidweekTime", t("LABEL_MIDWEEK_TIME"), 10, 70, 200, 70, 100)

	; Weekend meeting settings
	_AddDayDropdownField("WeekendDay", t("LABEL_WEEKEND_DAY"), 10, 100, 200, 100, 100)
	_AddTextInputField("WeekendTime", t("LABEL_WEEKEND_TIME"), 10, 130, 200, 130, 100)

	_AddTextInputFieldWithTooltip("KeyboardShortcut", t("LABEL_KEYBOARD_SHORTCUT"), 10, 220, 140, 220, 160, "LABEL_KEYBOARD_SHORTCUT_EXPLAIN", '')

	; Zoom UI element text values (for internationalization support)
	_AddTextInputFieldWithTooltip("HostToolsValue", t("LABEL_HOST_TOOLS"), 10, 250, 140, 250, 160, "LABEL_HOST_TOOLS_EXPLAIN", "host_tools.jpg")
	_AddTextInputFieldWithTooltip("ParticipantValue", t("LABEL_PARTICIPANT"), 10, 280, 140, 280, 160, "LABEL_PARTICIPANT_EXPLAIN", "participant.jpg")
	_AddTextInputFieldWithTooltip("MuteAllValue", t("LABEL_MUTE_ALL"), 10, 310, 140, 310, 160, "LABEL_MUTE_ALL_EXPLAIN", "mute_all.jpg")
	_AddTextInputFieldWithTooltip("YesValue", t("LABEL_YES"), 10, 340, 140, 340, 160, "LABEL_YES_EXPLAIN", "yes.jpg")
	_AddTextInputFieldWithTooltip("UncheckedValue", t("LABEL_UNCHECKED_VALUE"), 10, 370, 140, 370, 160, "LABEL_UNCHECKED_VALUE_EXPLAIN", "unchecked.jpg")
	_AddTextInputFieldWithTooltip("CurrentlyUnmutedValue", t("LABEL_CURRENTLY_UNMUTED_VALUE"), 10, 400, 140, 400, 160, "LABEL_CURRENTLY_UNMUTED_VALUE_EXPLAIN", "currently_unmuted.jpg")
	_AddTextInputFieldWithTooltip("UnmuteAudioValue", t("LABEL_UNMUTE_AUDIO_VALUE"), 10, 430, 140, 430, 160, "LABEL_UNMUTE_AUDIO_VALUE_EXPLAIN", "unmute_audio.jpg")
	_AddTextInputFieldWithTooltip("StopVideoValue", t("LABEL_STOP_VIDEO_VALUE"), 10, 460, 140, 460, 160, "LABEL_STOP_VIDEO_VALUE_EXPLAIN", "stop_video.jpg")
	_AddTextInputFieldWithTooltip("StartVideoValue", t("LABEL_START_VIDEO_VALUE"), 10, 490, 140, 490, 160, "LABEL_START_VIDEO_VALUE_EXPLAIN", "start_video.jpg")

	; Error display area
	$g_ErrorAreaLabel = GUICtrlCreateLabel("", 10, 520, 300, 20)
	GUICtrlSetColor($g_ErrorAreaLabel, 0xFF0000) ; Red text for errors

	; Action buttons
	$idSaveBtn = GUICtrlCreateButton(t("BTN_SAVE"), 10, 550, 80, 30)
	Local $idGetElementsBtn = GUICtrlCreateButton("Get Element Names", 100, 550, 120, 30)
	Local $idQuitBtn = GUICtrlCreateButton(t("BTN_QUIT"), 230, 550, 80, 30)

	; Set initial button states
	GUICtrlSetState($idSaveBtn, $GUI_DISABLE)  ; Disabled until all fields valid
	GUICtrlSetState($idQuitBtn, $GUI_ENABLE)

	; Set button event handlers
	GUICtrlSetOnEvent($idSaveBtn, "SaveConfigGUI")
	GUICtrlSetOnEvent($idGetElementsBtn, "GetElementNames")
	GUICtrlSetOnEvent($idQuitBtn, "QuitApp")

	; Perform initial validation check
	CheckConfigFields()

	; Show the GUI and register message handler for real-time validation
	GUISetState(@SW_SHOW, $g_ConfigGUI)
	; Shows a "Please Wait" message dialog during long operations
	GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND_EditChange")
EndFunc   ;==>ShowConfigGUI

; Shows a message dialog during long operations with i18n support
; @param $messageType - Type of message to show ('PLEASE_WAIT' or 'POST_MEETING_HIT_KEY')
Func ShowOverlayMessage($messageType = 'PLEASE_WAIT')
	; If already showing, just update it
	If $g_OverlayMessageGUI <> 0 Then
		; Update existing GUI with new message
		Local $title = t($messageType & '_TITLE')
		Local $text = t($messageType & '_TEXT')
		WinSetTitle($g_OverlayMessageGUI, '', $title)
		Local $idLbl = _GetOverlayMessageLabelControl()
		If $idLbl <> 0 Then
			; Use GUICtrl functions since we have the control ID
			GUICtrlSetData($idLbl, $text)
			; Update font styling to use default Windows system font
			GUICtrlSetFont($idLbl, 14, 700, Default, "Segoe UI")
		EndIf
		GUISetState(@SW_SHOW, $g_OverlayMessageGUI)
		WinSetOnTop(HWnd($g_OverlayMessageGUI), '', $WINDOWS_ONTOP)
		Return
	EndIf

	Local $iW = 350
	Local $iH = 140
	Local $iX = (@DesktopWidth - $iW) / 2
	Local $iY = (@DesktopHeight - $iH) / 2

	; Create borderless, always-on-top popup on primary monitor
	$g_OverlayMessageGUI = GUICreate(t($messageType & '_TITLE'), $iW, $iH, $iX, $iY, $WS_POPUP, $WS_EX_TOPMOST)
	GUISetBkColor(0x0000FF, $g_OverlayMessageGUI) ; Blue background

	; Create a label that supports both centering and word wrapping
	Local $idLbl = GUICtrlCreateLabel(t($messageType & '_TEXT'), 10, 10, $iW - 20, $iH - 20, $SS_CENTER)
	GUICtrlSetColor($idLbl, 0xFFFFFF)
	GUICtrlSetFont($idLbl, 14, 700, Default, "Segoe UI") ; Use default Windows system font

	GUISetState(@SW_SHOW, $g_OverlayMessageGUI)
	WinSetOnTop(HWnd($g_OverlayMessageGUI), '', $WINDOWS_ONTOP)
EndFunc   ;==>ShowOverlayMessage

; Hides and destroys the "Please Wait" message dialog
Func HideOverlayMessage()
	If $g_OverlayMessageGUI <> 0 Then
		GUIDelete($g_OverlayMessageGUI)
		$g_OverlayMessageGUI = 0
	EndIf
EndFunc   ;==>HideOverlayMessage

Func _GetOverlayMessageLabelControl()
	If $g_OverlayMessageGUI = 0 Then Return 0

	; Get all controls in the GUI using WinAPI
	Local $aControls = _WinAPI_EnumChildWindows($g_OverlayMessageGUI)
	If @error Or Not IsArray($aControls) Then Return 0

	; Find the label control (usually the first and only control)
	For $i = 1 To $aControls[0][0]
		Local $hCtrl = $aControls[$i][0]
		Local $sClass = _WinAPI_GetClassName($hCtrl)
		If $sClass = "Static" Then
			; Try to get the control ID from the handle
			Local $ctrlID = _WinAPI_GetDlgCtrlID($hCtrl)
			If $ctrlID > 0 Then Return $ctrlID
		EndIf
	Next

	Return 0
EndFunc   ;==>_GetOverlayMessageLabelControl

; Helper function to add text input field with label
; @param $key - Settings key name
; @param $label - Display label text
; @param $xLabel,$yLabel - Label position
; @param $xInput,$yInput - Input position
; @param $wInput - Input width
Func _AddTextInputField($key, $label, $xLabel, $yLabel, $xInput, $yInput, $wInput)
	GUICtrlCreateLabel($label, $xLabel, $yLabel, 180, 20)
	Local $idInput = GUICtrlCreateInput(GetUserSetting($key), $xInput, $yInput, $wInput, 20)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idInput)
	GUICtrlSetOnEvent($idInput, "CheckConfigFields")
EndFunc   ;==>_AddTextInputField

; Helper function to add text input field with label, info icon, and tooltip
; @param $key - Settings key name
; @param $label - Display label text
; @param $xLabel,$yLabel - Label position
; @param $xInput,$yInput - Input position
; @param $wInput - Input width
; @param $explainKey - Translation key for explanation text
; @param $imageName - Image filename for tooltip
Func _AddTextInputFieldWithTooltip($key, $label, $xLabel, $yLabel, $xInput, $yInput, $wInput, $explainKey, $imageName)
	GUICtrlCreateLabel($label, $xLabel, $yLabel, 110, 20)

	; Create info icon label (using Unicode info symbol)
	Local $idInfoIcon = GUICtrlCreateLabel("ℹ️", $xLabel + 115, $yLabel, 15, 20)
	GUICtrlSetColor($idInfoIcon, 0x0066CC) ; Blue color
	GUICtrlSetFont($idInfoIcon, 10, 700) ; Bold
	GUICtrlSetCursor($idInfoIcon, 0) ; Hand cursor

	; Build tooltip text with explanation and image reference
	Local $tooltipText = t($explainKey)
	Local $imagePath = @ScriptDir & "\images\" & $imageName

	; Check if image exists, if not use placeholder path
	If Not FileExists($imagePath) Then
		$imagePath = @ScriptDir & "\images\placeholder.jpg"
	EndIf

	; Set standard tooltip with explanation text only
	GUICtrlSetTip($idInfoIcon, $tooltipText)

	; Store image path for custom tooltip display on click
	If Not $g_InfoIconData.Exists($idInfoIcon) Then
		$g_InfoIconData.Add($idInfoIcon, $imagePath)
	EndIf

	; Set click event to show image tooltip
	GUICtrlSetOnEvent($idInfoIcon, "_ShowImageTooltip")

	Local $idInput = GUICtrlCreateInput(GetUserSetting($key), $xInput, $yInput, $wInput, 20)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idInput)
	GUICtrlSetOnEvent($idInput, "CheckConfigFields")
EndFunc   ;==>_AddTextInputFieldWithTooltip

; Helper function to add day selection dropdown with label
; @param $key - Settings key name
; @param $label - Display label text
; @param $xLabel,$yLabel - Label position
; @param $xInput,$yInput - Input position
; @param $wInput - Input width
Func _AddDayDropdownField($key, $label, $xLabel, $yLabel, $xInput, $yInput, $wInput)
	GUICtrlCreateLabel($label, $xLabel, $yLabel, 180, 20)

	; Build day list (Monday-Saturday, then Sunday for better UI flow)
	Local $dayList = ""
	Local $i
	For $i = 2 To 7  ; Monday through Saturday
		Local $lbl = t("DAY_" & $i)
		$dayList &= ($dayList = "" ? $lbl : "|" & $lbl)
	Next
	Local $lblSun = t("DAY_" & 1)  ; Sunday last
	$dayList &= ($dayList = "" ? $lblSun : "|" & $lblSun)

	; Set current selection based on saved setting
	Local $currentNum = String(GetUserSetting($key))
	Local $currentLabel = $currentNum
	If $g_DayNumToLabel.Exists($currentNum) Then $currentLabel = $g_DayNumToLabel.Item($currentNum)

	Local $idCombo = GUICtrlCreateCombo("", $xInput, $yInput, $wInput, 20)
	GUICtrlSetData($idCombo, $dayList, $currentLabel)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idCombo)
	GUICtrlSetOnEvent($idCombo, "CheckConfigFields")
EndFunc   ;==>_AddDayDropdownField

; Validates all configuration fields and updates UI accordingly
; Enables/disables save button, shows validation errors, and updates field styling
Func CheckConfigFields()
	Local $allFilled = True

	; Check if all fields have values
	For $sKey In $g_FieldCtrls.Keys
		Local $ctrlID = $g_FieldCtrls.Item($sKey)
		Local $val = StringStripWS(GUICtrlRead($ctrlID), 3)
		If $val = "" Then
			$allFilled = False
			; Mark required field with tooltip and light red background
			GUICtrlSetTip($ctrlID, t("ERROR_REQUIRED"))
			GUICtrlSetBkColor($ctrlID, 0xEEDDDD)
		EndIf
	Next

	; Additional format validation if all fields have values
	If $allFilled Then
		Local $ok = True

		; Validate meeting ID format
		Local $idCtrl = $g_FieldCtrls.Item("MeetingID")
		If Not _IsValidMeetingID(GUICtrlRead($idCtrl)) Then $ok = False

		; Validate time formats
		Local $midCtrl = $g_FieldCtrls.Item("MidweekTime")
		Local $wkdCtrl = $g_FieldCtrls.Item("WeekendTime")
		If Not _IsValidTime(GUICtrlRead($midCtrl)) Then $ok = False
		If Not _IsValidTime(GUICtrlRead($wkdCtrl)) Then $ok = False

		; Validate keyboard shortcut format (if not empty)
		Local $kbCtrl = $g_FieldCtrls.Item("KeyboardShortcut")
		Local $kbVal = StringStripWS(GUICtrlRead($kbCtrl), 3)
		If $kbVal <> "" And Not _IsValidKeyboardShortcut($kbVal) Then $ok = False

		$allFilled = $ok
	EndIf

	; Build error message and update field styling
	Local $aMsgs = ($allFilled ? "" : t("ERROR_FIELDS_REQUIRED"))

	; Clear required indicators for non-empty fields
	For $sKey In $g_FieldCtrls.Keys
		Local $ctrlId2 = $g_FieldCtrls.Item($sKey)
		Local $val2 = StringStripWS(GUICtrlRead($ctrlId2), 3)
		If $val2 <> "" Then
			GUICtrlSetTip($ctrlId2, "")
			GUICtrlSetBkColor($ctrlId2, 0xFFFFFF)
		EndIf
	Next

	; Specific validation for meeting ID
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

	; Specific validation for midweek time
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

	; Specific validation for weekend time
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

	; Specific validation for keyboard shortcut
	Local $sKbVal = StringStripWS(GUICtrlRead($g_FieldCtrls.Item("KeyboardShortcut")), 3)
	If $sKbVal <> "" And Not _IsValidKeyboardShortcut($sKbVal) Then
		$aMsgs &= ($aMsgs = "" ? t("ERROR_KEYBOARD_SHORTCUT_FORMAT") : "  •  " & t("ERROR_KEYBOARD_SHORTCUT_FORMAT"))
		GUICtrlSetColor($g_FieldCtrls.Item("KeyboardShortcut"), 0xFF0000)
		GUICtrlSetTip($g_FieldCtrls.Item("KeyboardShortcut"), t("ERROR_KEYBOARD_SHORTCUT_FORMAT"))
	Else
		If $sKbVal <> "" Then
			GUICtrlSetColor($g_FieldCtrls.Item("KeyboardShortcut"), 0x000000)
			GUICtrlSetTip($g_FieldCtrls.Item("KeyboardShortcut"), "")
		EndIf
	EndIf

	; Update error display area
	If $g_ErrorAreaLabel <> 0 Then GUICtrlSetData($g_ErrorAreaLabel, $aMsgs)

	; Enable/disable save button based on validation
	GUICtrlSetState($idSaveBtn, ($allFilled ? $GUI_ENABLE : $GUI_DISABLE))

	; Enable/disable window close button based on validation
	_EnableCloseButton($g_ConfigGUI, $allFilled)
EndFunc   ;==>CheckConfigFields

; Windows message handler for real-time validation during typing
; @param $hWnd - Window handle
; @param $iMsg - Message ID
; @param $wParam - Additional message parameter
; @param $lParam - Additional message parameter
; @return Integer - Message handling result
Func _WM_COMMAND_EditChange($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
	Local $iCtrlId = BitAND($wParam, 0xFFFF)      ; Extract control ID
	Local $iNotify = BitShift($wParam, 16)        ; Extract notification code

	; React to text change notifications from edit controls
	If $iNotify = $EN_CHANGE Then
		For $sKey In $g_FieldCtrls.Keys
			If $iCtrlId = $g_FieldCtrls.Item($sKey) Then
				CheckConfigFields()  ; Trigger validation
				ExitLoop
			EndIf
		Next
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_COMMAND_EditChange

; Enables or disables the window close button (X)
; @param $hWnd - Window handle
; @param $bEnable - True to enable, False to disable
; @return Integer - Success/error code
Func _EnableCloseButton($hWnd, $bEnable = True)
	Local $hMenu = DllCall("user32.dll", "hwnd", "GetSystemMenu", "hwnd", $hWnd, "int", 0)[0]
	If @error Or $hMenu = 0 Then Return SetError(1, 0, 0)

	Local $iFlag = ($bEnable ? $MF_ENABLED : $MF_GRAYED)
	DllCall("user32.dll", "int", "EnableMenuItem", "hwnd", $hMenu, "uint", $SC_CLOSE, "uint", BitOR($MF_BYCOMMAND, $iFlag))
	Return 1
EndFunc   ;==>_EnableCloseButton

; Saves configuration settings to INI file and closes GUI
Func SaveConfigGUI()
	$g_UserSettings.RemoveAll()

	; Save all field values to settings and INI file
	For $sKey In $g_FieldCtrls.Keys
		Local $ctrlID = $g_FieldCtrls.Item($sKey)
		Local $val = StringStripWS(GUICtrlRead($ctrlID), 3)

		; Convert day labels back to numbers for storage
		If ($sKey = "MidweekDay" Or $sKey = "WeekendDay") Then
			If $g_DayLabelToNum.Exists($val) Then $val = String($g_DayLabelToNum.Item($val))
		EndIf

		; Convert translated snap selection back to internal values
		If ($sKey = "SnapZoomSide") Then
			If $val = t("SNAP_LEFT") Then
				$val = "Left"
			ElseIf $val = t("SNAP_RIGHT") Then
				$val = "Right"
			Else
				$val = "Disabled"
			EndIf
		EndIf

		$g_UserSettings.Add($sKey, $val)
		IniWrite($CONFIG_FILE, _GetIniSectionForKey($sKey), $sKey, GetUserSetting($sKey))
	Next

	; Save language setting
	Local $selDisplay = GUICtrlRead($idLanguagePicker)
	Local $selLang = "en"
	If $g_LangNameToCode.Exists($selDisplay) Then $selLang = $g_LangNameToCode.Item($selDisplay)
	$g_UserSettings.Add("Language", $selLang)
	IniWrite($CONFIG_FILE, "General", "Language", GetUserSetting("Language"))
	$g_CurrentLang = $selLang

	; Save keyboard shortcut setting
	IniWrite($CONFIG_FILE, "General", "KeyboardShortcut", GetUserSetting("KeyboardShortcut"))
	$g_KeyboardShortcut = GetUserSetting("KeyboardShortcut")

	; Register/unregister hotkey based on new setting
	_UpdateKeyboardShortcut()

	_InitDayLabelMaps() ; Refresh day labels for new language
	$g_PrePostSettingsConfigured = False ; Reset pre/post meeting settings flag
	$g_DuringMeetingSettingsConfigured = False ; Reset during-meeting settings flag
	$g_InitialNotificationWasShown = False ; Reset initial notification flag
	LoadMeetingConfig() ; Reload config to apply any meeting schedule changes

	CloseConfigGUI()
EndFunc   ;==>SaveConfigGUI

; Closes and destroys the configuration GUI
Func CloseConfigGUI()
	GUIDelete($g_ConfigGUI)
	$g_ConfigGUI = 0
	_CloseImageTooltip() ; Close any open image tooltip
EndFunc   ;==>CloseConfigGUI

; Shows a custom tooltip window with an image when info icon is clicked
Func _ShowImageTooltip()
	Local $idClicked = @GUI_CtrlId

	; Check if this info icon has an associated image
	If Not $g_InfoIconData.Exists($idClicked) Then Return

	Local $imagePath = $g_InfoIconData.Item($idClicked)

	; Close any existing tooltip
	_CloseImageTooltip()

	; Create tooltip window with image
	Local $iW = 220
	Local $iH = 170
	Local $mousePos = MouseGetPos()
	Local $iX = $mousePos[0] + 10
	Local $iY = $mousePos[1] + 10

	; Adjust position if tooltip would go off screen
	If $iX + $iW > @DesktopWidth Then $iX = @DesktopWidth - $iW - 10
	If $iY + $iH > @DesktopHeight Then $iY = @DesktopHeight - $iH - 10

	$g_TooltipGUI = GUICreate("", $iW, $iH, $iX, $iY, $WS_POPUP, $WS_EX_TOPMOST)
	GUISetBkColor(0xFFFFE1, $g_TooltipGUI) ; Light yellow background

	Debug($imagePath)
	Debug(FileExists($imagePath))

	; Add image if it exists
	If FileExists($imagePath) Then
		GUICtrlCreatePic($imagePath, 10, 10, 200, 150)
	Else
		GUICtrlCreateLabel("Image not found:", 10, 10, 200, 20)
		GUICtrlCreateLabel($imagePath, 10, 35, 200, 100)
	EndIf

	GUISetState(@SW_SHOW, $g_TooltipGUI)

	; Auto-close after 5 seconds or when clicking anywhere
	AdlibRegister("_CloseImageTooltip", 3000)
EndFunc   ;==>_ShowImageTooltip

; Closes the custom image tooltip window
Func _CloseImageTooltip()
	If $g_TooltipGUI <> 0 Then
		GUIDelete($g_TooltipGUI)
		$g_TooltipGUI = 0
		AdlibUnRegister("_CloseImageTooltip")
	EndIf
EndFunc   ;==>_CloseImageTooltip

; Exits the application
Func QuitApp()
	Exit
EndFunc   ;==>QuitApp

; ================================================================================================
; TRAY ICON EVENT HANDLING
; ================================================================================================

; Sets up tray icon and handles tray events
TraySetIcon($g_TrayIcon)

; Handles all tray icon events (left-click, right-click, double-click)
Func TrayEvent()
	UpdateTrayTooltip()
	Local $trayMsg = TrayGetMsg()
	Switch $trayMsg
		Case $TRAY_EVENT_PRIMARYDOWN ; Left-click
			ShowConfigGUI()
			While $g_ConfigGUI
				Sleep(100)
			WEnd
		Case $TRAY_EVENT_SECONDARYDOWN ; Right-click
			ShowConfigGUI()
			While $g_ConfigGUI
				Sleep(100)
			WEnd
		Case $TRAY_EVENT_PRIMARYDOUBLE ; Double-click
			ShowConfigGUI()
			While $g_ConfigGUI
				Sleep(100)
			WEnd
	EndSwitch
EndFunc   ;==>TrayEvent

; Sleep function that continues to handle tray events during wait
; @param $s - Seconds to sleep
Func ResponsiveSleep($s)
	Local $elapsed = 0
	Local $ms = $s * 1000
	While $elapsed < $ms
		TrayEvent()           ; Continue handling tray events
		Sleep(50)            ; Small incremental sleep
		$elapsed += 50
	WEnd
EndFunc   ;==>ResponsiveSleep

; ================================================================================================
; UIAUTOMATION COM INITIALIZATION
; ================================================================================================

; Initialize UIAutomation COM object for interacting with Zoom UI
$oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation)
If Not IsObj($oUIAutomation) Then
	Debug("Failed to create UIAutomation COM object.", "UIA")
	Exit
EndIf
Debug("UIAutomation COM created successfully.", "UIA")

; Get desktop element as root for UI searches
$oUIAutomation.GetRootElement($pDesktop)
$oDesktop = ObjCreateInterface($pDesktop, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
If Not IsObj($oDesktop) Then
	Debug(t("ERROR_GET_DESKTOP_ELEMENT_FAILED"), "ERROR")
	Exit
EndIf
Debug("Desktop element obtained.", "UIA")

; ================================================================================================
; ZOOM WINDOW AND UI ELEMENT DISCOVERY
; ================================================================================================

; Finds UI element by class name within a specified scope
; @param $sClassName - Class name to search for
; @param $iScope - Search scope (default: descendants)
; @param $oParent - Parent element to search within (default: desktop)
; @return Object - Found element or 0 if not found
Func FindElementByClassName($sClassName, $iScope = Default, $oParent = Default)
	Debug("Searching for element with class: '" & $sClassName & "'", "DEBUG")

	; Use desktop as default parent if not specified
	Local $oSearchParent = $oDesktop
	If $oParent <> Default Then $oSearchParent = $oParent

	; Create condition for class name search
	Local $pClassCondition
	$oUIAutomation.CreatePropertyCondition($UIA_ClassNamePropertyId, $sClassName, $pClassCondition)

	; Perform the search
	Local $pElement
	Local $scope = $TreeScope_Descendants
	If $iScope <> Default Then $scope = $iScope
	$oSearchParent.FindFirst($scope, $pClassCondition, $pElement)

	; Create element interface
	Local $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
	If Not IsObj($oElement) Then
		Debug("Element with class '" & $sClassName & "' not found.", "WARN")
		Return 0
	EndIf

	Debug("Element with class '" & $sClassName & "' found.", "DEBUG")
	Return $oElement
EndFunc   ;==>FindElementByClassName

; Gets reference to the main Zoom meeting window
; @return Object - Zoom window element or error
Func _GetZoomWindow()
	$oZoomWindow = FindElementByClassName("ConfMultiTabContentWndClass", $TreeScope_Children)
	If Not IsObj($oZoomWindow) Then Return SetError(1, 0, 0)
	Debug("Zoom window obtained.", "UIA")
EndFunc   ;==>_GetZoomWindow

; Focuses the main Zoom meeting window
; @return Boolean - True if successful, False otherwise
Func FocusZoomWindow()
	_GetZoomWindow()
	If Not IsObj($oZoomWindow) Then
		Debug("Unable to obtain Zoom window.", "DEBUG")
		Return False
	EndIf

	; Get the native HWND property from the UIA element
	Local $hWnd
	$oZoomWindow.GetCurrentPropertyValue($UIA_NativeWindowHandlePropertyId, $hWnd)

	If $hWnd And $hWnd <> 0 Then
		; Convert to HWND pointer
		$hWnd = Ptr($hWnd)
		WinActivate($hWnd)
		If WinWaitActive($hWnd, "", 3) Then
			Debug("Zoom window activated and focused.", "SUCCESS")
			Return True
		Else
			Debug("Zoom window did not become active within timeout.", "WARN")
		EndIf
	Else
		Debug("Zoom window has no valid HWND property.", "ERROR")
	EndIf

	Return False
EndFunc   ;==>FocusZoomWindow


; Snaps the Zoom window to a side of the primary monitor using Windows snap shortcuts
; @return Boolean - True if moved, False otherwise
Func _SnapZoomWindowToSide()
	Local $snapSide = GetUserSetting("SnapZoomSide")
	If StringLower($snapSide) = "disabled" Then
		Debug("Zoom window snapping disabled; skipping.", "DEBUG")
		Return False
	EndIf

	_GetZoomWindow()
	If Not IsObj($oZoomWindow) Then
		Debug("_SnapZoomWindowToSide: Zoom window not found", "WARN")
		Return False
	EndIf

	; Activate the Zoom window first using the proper focus function
	FocusZoomWindow()

	; Use Windows snap shortcuts instead of manual positioning
	If StringLower($snapSide) = "left" Then
		Send("#{LEFT}") ; Windows + Left arrow
		Debug("Sent Windows+Left to snap Zoom window to left half", "DEBUG")
	ElseIf StringLower($snapSide) = "right" Then
		Send("#{RIGHT}") ; Windows + Right arrow
		Debug("Sent Windows+Right to snap Zoom window to right half", "DEBUG")
	Else
		Debug("_SnapZoomWindowToSide: Invalid snap side: " & $snapSide, "ERROR")
		Return False
	EndIf

	; Wait 0.5 seconds for snap animation to complete
	Sleep(500)

	; Send Escape to dismiss any remaining Windows snap UI
	Send("{ESC}")
	Debug("Sent Escape to dismiss Windows snap UI", "DEBUG")

	Debug("Zoom window snapped to " & $snapSide & " half of primary monitor using Windows shortcuts.", "DEBUG")
	Return True
EndFunc   ;==>_SnapZoomWindowToSide


; Finds UI element by partial name match across multiple control types
; @param $sPartial - Partial text to search for in element names
; @param $aControlTypes - Array of control types to search (default: button and menu item)
; @param $oParent - Parent element to search within (default: desktop)
; @return Object - Found element or 0 if not found
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

	; Search through each specified control type
	For $iType = 0 To UBound($aControlTypes) - 1
		Local $iControlType = $aControlTypes[$iType]
		Debug("Searching control type: " & $iControlType)

		; Create condition for this control type
		Local $pCondition
		$oUIAutomation.CreatePropertyCondition($UIA_ControlTypePropertyId, $iControlType, $pCondition)

		; Find all elements of this type
		Local $pElements
		$oSearchParent.FindAll($TreeScope_Descendants, $pCondition, $pElements)

		Local $oElements = ObjCreateInterface($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
		If IsObj($oElements) Then
			Local $iCount
			$oElements.Length($iCount)
			Debug("Found " & $iCount & " elements of this type.", "DEBUG")

			; Check each element for partial name match
			For $i = 0 To $iCount - 1
				Local $pElement
				$oElements.GetElement($i, $pElement)

				Local $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
				If IsObj($oElement) Then
					; Get element name and check for partial match
					Local $sName
					$oElement.GetCurrentPropertyValue($UIA_NamePropertyId, $sName)
					Debug("Element found with name: '" & $sName & "'", "DEBUG")

					If StringInStr($sName, $sPartial, $STR_NOCASESENSEBASIC) > 0 Then
						Debug("Matching element found with name: '" & $sName & "'", "DEBUG")
						Return $oElement
					EndIf
				EndIf
			Next
		EndIf
	Next

	Debug("No element found containing: '" & $sPartial & "'", "WARN")
	Return 0
EndFunc   ;==>FindElementByPartialName


; Searches for UI elements by control type and prints their hierarchy for debugging
; @param $iControlType - Control type ID to search for
; @param $oParent - Parent element to search within (default: desktop)
; @return Object - First element found or 0 if none found
Func FindElementByControlType($iControlType, $oParent = Default)
	Debug("Searching for ALL elements of control type: " & $iControlType, "DEBUG")

	Local $oSearchParent = $oDesktop
	If $oParent <> Default Then
		$oSearchParent = $oParent
		Debug("Using custom parent element for search", "DEBUG")
	Else
		Debug("Using desktop as search parent", "DEBUG")
	EndIf

	Local $pCondition
	$oUIAutomation.CreatePropertyCondition($UIA_ControlTypePropertyId, $iControlType, $pCondition)

	Local $pElements
	; 🔑 Use Descendants instead of Subtree
	$oSearchParent.FindAll($TreeScope_Descendants, $pCondition, $pElements)

	Local $oElements = ObjCreateInterface($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
	If Not IsObj($oElements) Then
		Debug("Failed to create element array object", "ERROR")
		Return 0
	EndIf

	Local $count
	$oElements.Length($count)

	If $count = 0 Then
		Debug("No elements found for control type: " & $iControlType, "WARN")
		Return 0
	EndIf

	Debug("Found " & $count & " element(s) of control type " & $iControlType, "DEBUG")

	; If only one element, return it directly
	If $count = 1 Then
		Local $pElem
		$oElements.GetElement(0, $pElem)
		Local $oElem = ObjCreateInterface($pElem, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)

		Local $sName
		$oElem.GetCurrentPropertyValue($UIA_NamePropertyId, $sName)
		Debug("Single element name: " & $sName, "DEBUG")

		Return $oElem
	EndIf

	; Otherwise, print hierarchy recursively
	For $i = 0 To $count - 1
		Local $pElement
		$oElements.GetElement($i, $pElement)

		Local $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
		If IsObj($oElement) Then
			; Get element name and log it
			$sName = GetElementName($oElement)
			Debug("Element #" & $i & " name: '" & $sName & "'", "INFO")

			; Optional: recursive tree print
			_PrintElementTree($oElement, 1)
		EndIf
	Next

	Return 0
EndFunc   ;==>FindElementByControlType



; Recursively prints element name + children
Func _PrintElementTree($oElem, $level = 0)
	If Not IsObj($oElem) Then Return

	; Get element name
	Local $sName = ""
	$oElem.GetCurrentPropertyValue($UIA_NamePropertyId, $sName)

	; Print with indentation
	Local $indent = StringFormat("%" & ($level * 2) & "s", "")
	Debug($indent & "- " & $sName, "TREE")

	; Get children (no condition)
	Local $pChildren
	$oElem.FindAll($TreeScope_Children, 0, $pChildren)

	Local $oChildren = ObjCreateInterface($pChildren, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
	If Not IsObj($oChildren) Then Return

	Local $childCount
	$oChildren.Length($childCount)

	For $i = 0 To $childCount - 1
		Local $pChild
		$oChildren.GetElement($i, $pChild)

		Local $oChild = ObjCreateInterface($pChild, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
		If IsObj($oChild) Then
			_PrintElementTree($oChild, $level + 1)
		EndIf
	Next
EndFunc   ;==>_PrintElementTree


; ================================================================================================
; UI ELEMENT INTERACTION FUNCTIONS
; ================================================================================================

; Clicks a UI element using multiple methods for maximum compatibility
; @param $oElement - Element to click
; @param $ForceClick - If True, forces mouse click method
; @param $BoundingRectangle - If True, forces click by bounding rectangle
; @return Boolean - True if successful, False otherwise
Func _ClickElement($oElement, $ForceClick = False, $BoundingRectangle = False)
	ResponsiveSleep(0.5) ; Brief pause to ensure UI is ready
	If Not IsObj($oElement) Then
		Debug(t("ERROR_INVALID_ELEMENT_OBJECT"), "WARN")
		Return False
	EndIf

	; Get element name for debugging
	Local $sElementName = GetElementName($oElement)
	Debug("Attempting to click element: '" & $sElementName & "'", "DEBUG")

	; Method 0: Force mouse click (only when requested)
	If $ForceClick Then
		; Method 0.5: Force mouse click by bounding rectangle (only when requested)
		If $BoundingRectangle Then
			ClickByBoundingRectangle($oElement)
			Debug("Element clicked via bounding rectangle.", "SUCCESS")
		Else
			UIA_MouseClick($oElement)
			Debug("Element clicked via UIA_MouseClick.", "SUCCESS")
		EndIf
		Return True
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

	; Method 2: Try Legacy Accessible Pattern (works for menu items and older controls)
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
	ClickByBoundingRectangle($oElement)

	; All click methods failed
	Debug(t("ERROR_FAILED_CLICK_ELEMENT") & ": '" & $sElementName & "'", "ERROR")
	Return False
EndFunc   ;==>_ClickElement

Func ClickByBoundingRectangle($oElement)
	; Method 0.1: Click by bounding rectangle center
	Local $sElementName = GetElementName($oElement)
	Local $tRect
	$oElement.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $tRect)
	UIA_GetArrayPropertyValueAsString($tRect)
	Debug("Element bounding rectangle: " & $tRect, "DEBUG")
	If Not $tRect Then
		Debug("No bounding rectangle for element: '" & $sElementName & "'", "ERROR")
		Return False
	EndIf

	Local $aRect = StringSplit($tRect, ",")
	If $aRect[0] < 4 Then
		Debug("Invalid rectangle format for: '" & $sElementName & "'", "ERROR")
		Return False
	EndIf

	Local $iLeft = Number($aRect[1])
	Local $iTop = Number($aRect[2])
	Local $iWidth = Number($aRect[3])
	Local $iHeight = Number($aRect[4])

	Local $iCenterX = $iLeft + ($iWidth / 2)
	Local $iCenterY = $iTop + ($iHeight / 2)

	Debug("Using mouse click fallback at position: " & $iCenterX & "," & $iCenterY & " for: '" & $sElementName & "'", "DEBUG")

	; Ensure element is clickable before attempting
	Local $bIsEnabled, $bIsOffscreen
	$oElement.GetCurrentPropertyValue($UIA_IsEnabledPropertyId, $bIsEnabled)
	$oElement.GetCurrentPropertyValue($UIA_IsOffscreenPropertyId, $bIsOffscreen)

	If $bIsEnabled And Not $bIsOffscreen Then
		MouseClick("primary", $iCenterX, $iCenterY, 1, 0)
		Debug("Element clicked via mouse at center.", "SUCCESS")
		Return True
	Else
		Debug("Element not clickable - Enabled: " & $bIsEnabled & ", Offscreen: " & $bIsOffscreen, "WARN")
		Return False
	EndIf
EndFunc   ;==>ClickByBoundingRectangle


; Gets the name property of a UI element
; @param $oElement - The UI element object
; @return String - Element name or empty string if not found
Func GetElementName($oElement)
	Local $sName = ""
	If IsObj($oElement) Then
		$oElement.GetCurrentPropertyValue($UIA_NamePropertyId, $sName)
	EndIf
	Return $sName
EndFunc   ;==>GetElementName


; Hovers over a UI element by moving the mouse to its center
; @param $oElement - The UIA element object
; @param $iHoverTime - Time in milliseconds to hold the hover (default: 1000ms)
; @param $SlightOffset - If True, adds slight random offset to avoid exact center
; @return Boolean - True if successful, False otherwise
Func _HoverElement($oElement, $iHoverTime = 1000, $SlightOffset = False)
	ResponsiveSleep(0.3) ; Small buffer before hover
	If Not IsObj($oElement) Then
		Debug("Invalid element passed to _HoverElement", "WARN")
		Return False
	EndIf

	; Get element name for debugging
	Local $sElementName = GetElementName($oElement)
	Debug("Attempting to hover element: '" & $sElementName & "'", "DEBUG")

	; Get bounding rectangle
	Local $tRect
	$oElement.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $tRect)
	UIA_GetArrayPropertyValueAsString($tRect)
	Debug("Element bounding rectangle: " & $tRect, "DEBUG")
	If Not $tRect Then
		Debug("No bounding rectangle for element: '" & $sElementName & "'", "ERROR")
		Return False
	EndIf

	Local $aRect = StringSplit($tRect, ",")
	If $aRect[0] < 4 Then
		Debug("Invalid rectangle format for: '" & $sElementName & "'", "ERROR")
		Return False
	EndIf

	Local $iLeft = Number($aRect[1])
	Local $iTop = Number($aRect[2])
	Local $iWidth = Number($aRect[3])
	Local $iHeight = Number($aRect[4])

	Local $iCenterX = $iLeft + ($iWidth / 2)
	Local $iCenterY = $iTop + ($iHeight / 2)

	If $SlightOffset Then
		; Add slight random offset to avoid exact center (may help with some UI elements)
		Local $iOffsetX = Random(2, 5, 1)
		Local $iOffsetY = Random(-5, -2, 1)
		$iCenterX += $iOffsetX
		$iCenterY += $iOffsetY
		Debug("Applying slight offset to hover position: " & $iOffsetX & "," & $iOffsetY, "DEBUG")
	EndIf

	Debug("Hovering at: " & $iCenterX & "," & $iCenterY & " for " & $iHoverTime & "ms", "DEBUG")

	; Move mouse to center of element and hold position
	MouseMove($iCenterX, $iCenterY, 0)
	Sleep($iHoverTime)

	Debug("Hover completed on element: '" & $sElementName & "'", "SUCCESS")
	Return True
EndFunc   ;==>_HoverElement

; Moves mouse to the start of an element and optionally clicks it
; @param $oElement - The UIA element object
; @param $Click - If True, performs a click after moving (default: False)
; @return Boolean - True if successful, False otherwise
Func _MoveMouseToStartOfElement($oElement, $Click = False)
	ResponsiveSleep(0.3) ; Small buffer before move
	If Not IsObj($oElement) Then
		Debug("Invalid element passed to _MoveMouseToStartOfElement", "WARN")
		Return False
	EndIf

	; Get element name for debugging
	Local $sElementName = GetElementName($oElement)
	Debug("Attempting to move mouse to start of element: '" & $sElementName & "'", "DEBUG")

	; Get bounding rectangle
	Local $tRect
	$oElement.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $tRect)
	UIA_GetArrayPropertyValueAsString($tRect)
	Debug("Element bounding rectangle: " & $tRect, "DEBUG")
	If Not $tRect Then
		Debug("No bounding rectangle for element: '" & $sElementName & "'", "ERROR")
		Return False
	EndIf

	Local $aRect = StringSplit($tRect, ",")
	If $aRect[0] < 4 Then
		Debug("Invalid rectangle format for: '" & $sElementName & "'", "ERROR")
		Return False
	EndIf

	Local $iLeft = Number($aRect[1])
	Local $iTop = Number($aRect[2])
;~ Local $iWidth = Number($aRect[3])
	Local $iHeight = Number($aRect[4])

	; Move mouse to start (left edge, vertically centered)
	Local $iStartX = $iLeft + Random(5, 30, 1) ; Random offset from left edge
	Local $iStartY = $iTop + ($iHeight / 2) + Random(-5, 5, 1) ; Random offset from center

	Debug("Moving mouse to start position: " & $iStartX & "," & $iStartY, "DEBUG")

	; Move mouse to start of element
	MouseMove($iStartX, $iStartY, 0)
	Sleep(200) ; Brief pause after move
	Debug("Mouse moved to start of element: '" & $sElementName & "'", "SUCCESS")

	If $Click Then
		MouseClick("primary", $iStartX, $iStartY, 1, 0)
		Debug("Element clicked at start position.", "SUCCESS")
	EndIf

	Return True
EndFunc   ;==>_MoveMouseToStartOfElement


; ================================================================================================
; ZOOM-SPECIFIC UI INTERACTION FUNCTIONS
; ================================================================================================

; Opens the Host Tools menu in Zoom
; @return Object - Host menu object or False if failed
Func _OpenHostTools()
	If Not IsObj($oZoomWindow) Then Return False

	; Check if host menu is already open
	Local $oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
	If Not IsObj($oHostMenu) Then

		; Controls might be hidden, show them by moving the mouse
		Debug("Controls might be hidden. Moving mouse to show controls.", "DEBUG")
		Debug("Getting Zoom window.", "DEBUG")
		_GetZoomWindow()
		Debug("Moving mouse to start of element: 'Zoom window'", "DEBUG")
		 _MoveMouseToStartOfElement($oZoomWindow, True)
		 Debug("Clicking Host Tools button.", "DEBUG")

		; Menu not open, find and click the Host Tools button
		Local $oHostToolsButton = FindElementByPartialName(GetUserSetting("HostToolsValue"), Default, $oZoomWindow)

		; Scenario 1: Try to find Host Tools button directly
		If IsObj($oHostToolsButton) Then
			Debug("Host Tools button found directly; clicking.", "DEBUG")
			If Not _ClickElement($oHostToolsButton) Then
				Debug("Failed to click Host Tools.", "ERROR")
				Return False
			EndIf
		Else
			; Scenario 2: Try to find More menu, then Host Tools
			Debug("Host Tools button not found, looking for 'More' button.", "DEBUG")
			Local $oMoreMenu = GetMoreMenu()
			If IsObj($oMoreMenu) Then
				; Now look for the Host Tools button in the More menu
				Local $oHostToolsMenuItem = FindElementByPartialName(GetUserSetting("HostToolsValue"), Default, $oMoreMenu)
				If IsObj($oHostToolsMenuItem) Then
					Debug("Found Host Tools menu item. Hovering it to open submenu.", "DEBUG")
					If _HoverElement($oHostToolsMenuItem, 500) Then
						; Return the now-open host menu
						$oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
						Return $oHostMenu
					Else
						Debug("Failed to hover Host Tools menu item in More menu.", "ERROR")
						Return False
					EndIf
				Else
					Debug("Failed to find Host Tools menu item in More menu.", "ERROR")
					Return False
				EndIf
			Else
				Debug("Failed to find More menu to access Host Tools.", "ERROR")
				Return False
			EndIf
		EndIf
	EndIf

	; Return the now-open host menu
	$oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
	Return $oHostMenu
EndFunc   ;==>_OpenHostTools

; Closes the Host Tools menu by clicking on the main window
Func _CloseHostTools()
	_MoveMouseToStartOfElement($oZoomWindow, True) ; Click at start of window to ensure menu closes
	Debug("Host tools menu closed.", "UIA")
EndFunc   ;==>_CloseHostTools

; Opens the "More" menu in Zoom if available
; @return Object - More menu object or False if failed
Func GetMoreMenu()
	If Not IsObj($oZoomWindow) Then Return False

	Local $oMoreMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)

	If Not IsObj($oMoreMenu) Then
		; Menu not open, find and click the More button
		Debug("More menu not open, attempting to open.", "DEBUG")
		Debug("Controls might be hidden. Moving mouse to show controls.", "DEBUG")
		_MoveMouseToStartOfElement($oZoomWindow)

		Local $oMoreButton = FindElementByPartialName(GetUserSetting("MoreMeetingControlsValue"), Default, $oZoomWindow)
		If Not IsObj($oMoreButton) Then
			Debug("Failed to find More button.", "ERROR")
			Return False
		EndIf
		Debug("Clicking More button to open menu.", "DEBUG")
		If Not _ClickElement($oMoreButton, True) Then
			Debug("Failed to click More button.", "ERROR")
			Return False
		EndIf
		; Wait briefly for the More menu to open
		ResponsiveSleep(0.5)
	EndIf

	; Return the now-open More menu
	$oMoreMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
	If IsObj($oMoreMenu) Then
		Debug("More menu opened.", "UIA")
	Else
		Debug("Failed to open More menu.", "ERROR")
	EndIf
	Return $oMoreMenu
EndFunc   ;==>GetMoreMenu

; Opens the Participants panel in Zoom
; @return Object - Participants panel object or False if failed
Func _OpenParticipantsPanel()
	If Not IsObj($oZoomWindow) Then Return False

	; Controls might be hidden, show them by moving the mouse
	Debug("Controls might be hidden. Moving mouse to show controls.", "DEBUG")
	_MoveMouseToStartOfElement($oZoomWindow)

	Local $ListType[1] = [$UIA_ListControlTypeId]
	Local $oParticipantsPanel = FindElementByPartialName(GetUserSetting("ParticipantValue"), $ListType, $oZoomWindow)

	If Not IsObj($oParticipantsPanel) Then
		; Panel not open, find and click the Participants button
		Debug("Participants panel not open, attempting to open.", "UIA")
		Local $oMainParticipantsButton = FindElementByPartialName(GetUserSetting("ParticipantValue"), Default, $oZoomWindow)

		; Scenario 1: Try to find Participants button directly
		If IsObj($oMainParticipantsButton) Then
			Debug("Participants button found directly; clicking.", "DEBUG")
			If Not _ClickElement($oMainParticipantsButton) Then
				Debug("Failed to click Participants.", "ERROR")
				Return False
			EndIf
		Else
			; Scenario 2: Try to find More menu, then Participants
			Debug("Participants button not found, looking for 'More' button.", "DEBUG")
			Local $oMoreMenu = GetMoreMenu()
			If IsObj($oMoreMenu) Then
				; Now look for the Participants button in the More menu
				Local $oParticipantsMenuItem = FindElementByPartialName(GetUserSetting("ParticipantValue"), Default, $oMoreMenu)
				If IsObj($oParticipantsMenuItem) Then
					Debug("Found Participants menu item. Hovering it to open submenu.", "DEBUG")
					If _HoverElement($oParticipantsMenuItem, 1200) Then         ; 1.2s hover to ensure submenu appears
						; Now look for the Participants button again in the submenu
						Debug("Looking for Participants button again in submenu.", "DEBUG")
						Local $oParticipantsSubMenuItem = FindElementByPartialName(GetUserSetting("ParticipantValue"), Default, $oZoomWindow)
						If IsObj($oParticipantsSubMenuItem) Then
							Debug("Final Participants button found. Clicking it.", "DEBUG")
							_HoverElement($oParticipantsSubMenuItem, 500)
							_MoveMouseToStartOfElement($oParticipantsSubMenuItem, True)
							Debug("Participants button clicked.", "SUCCESS")
							ResponsiveSleep(0.5) ; Move mouse to start of element and click to avoid hover issues
						Else
							Debug("Failed to find Participants button in submenu.", "ERROR")
							Return False
						EndIf
					Else
						Debug("Failed to hover Participants menu item in More menu.", "ERROR")
						Return False
					EndIf
				Else
					Debug("Failed to find Participants menu item in More menu.", "ERROR")
					Return False
				EndIf
			Else
				Debug("Failed to find More menu to access Participants.", "ERROR")
				Return False
			EndIf
		EndIf
	EndIf

	; Return the now-open participants panel
	$oParticipantsPanel = FindElementByPartialName(GetUserSetting("ParticipantValue"), $ListType, $oZoomWindow)
	If IsObj($oParticipantsPanel) Then
		Debug("Participants panel opened.", "UIA")
		_SnapZoomWindowToSide()
	Else
		Debug("Failed to open Participants panel.", "ERROR")
	EndIf
	Return $oParticipantsPanel
EndFunc   ;==>_OpenParticipantsPanel

; ================================================================================================
; KEYBOARD SHORTCUT MANAGEMENT
; ================================================================================================

; Registers or unregisters the keyboard shortcut for manual trigger
; Called when keyboard shortcut setting is changed
Func _UpdateKeyboardShortcut()
	; Unregister any currently registered hotkey
	If $g_HotkeyRegistered Then
		HotKeySet($g_KeyboardShortcut)
		$g_HotkeyRegistered = False
		Debug("Previous keyboard shortcut unregistered: " & $g_KeyboardShortcut, "DEBUG")
	EndIf

	; Update the global keyboard shortcut variable
	$g_KeyboardShortcut = GetUserSetting("KeyboardShortcut")

	; Register new hotkey if not empty
	If $g_KeyboardShortcut <> "" Then
		; Validate the shortcut format (basic validation)
		If StringRegExp($g_KeyboardShortcut, "^[\^\!\+\#]+[a-zA-Z0-9]$") Then
			HotKeySet($g_KeyboardShortcut, "_ManualTrigger")
			$g_HotkeyRegistered = True
			Debug("New keyboard shortcut registered: " & $g_KeyboardShortcut, "DEBUG")
		Else
			Debug("Invalid keyboard shortcut format: " & $g_KeyboardShortcut, "ERROR")
			$g_KeyboardShortcut = ""
			IniWrite($CONFIG_FILE, "General", "KeyboardShortcut", "")
		EndIf
	Else
		Debug("Keyboard shortcut cleared", "DEBUG")
	EndIf
EndFunc   ;==>_UpdateKeyboardShortcut

; Manual trigger function activated by keyboard shortcut
; Allows user to manually apply post-meeting settings
Func _ManualTrigger()    ; Show message and wait for user input before applying settings
	Debug("Manual trigger: Showing post-meeting message", "DEBUG")

	; Show the post-meeting message and wait for Enter key
	ShowOverlayMessage('POST_MEETING_HIT_KEY')

	; Wait for user to press Enter or ESC
	Local $userInput = _WaitForEnterOrEscape()

	; Hide the message
	HideOverlayMessage()

	; Apply settings if user pressed Enter
	If $userInput = "ENTER" Then
		Debug("User pressed Enter: Applying post-meeting settings", "DEBUG")
		_SetPreAndPostMeetingSettings()
	Else
		Debug("User pressed Escape or closed dialog: Cancelling", "DEBUG")
	EndIf
EndFunc   ;==>_ManualTrigger

; Waits for user to press Enter or Escape key
; @return String - "ENTER" if Enter was pressed, "ESCAPE" if Escape was pressed or dialog closed
Func _WaitForEnterOrEscape()
	Local $hUser32 = DllOpen("user32.dll")

	While True
		; Check for Enter key
		If _IsKeyPressed($hUser32, 0x0D) Then ; VK_RETURN
			DllClose($hUser32)
			Return "ENTER"
		EndIf

		; Check for Escape key
		If _IsKeyPressed($hUser32, 0x1B) Then ; VK_ESCAPE
			DllClose($hUser32)
			Return "ESCAPE"
		EndIf

		; Check if GUI was closed (handle becomes invalid)
		If $g_OverlayMessageGUI = 0 Or Not WinExists(HWnd($g_OverlayMessageGUI)) Then
			DllClose($hUser32)
			Return "ESCAPE"
		EndIf

		ResponsiveSleep(0.1) ; Small delay to avoid high CPU usage
	WEnd
EndFunc   ;==>_WaitForEnterOrEscape

; Helper function to check if a key is currently pressed
; @param $hDLL - Handle to user32.dll
; @param $iKeyCode - Virtual key code to check
; @return Boolean - True if key is pressed
Func _IsKeyPressed($hDLL, $iKeyCode)
	Local $aRet = DllCall($hDLL, "short", "GetAsyncKeyState", "int", $iKeyCode)
	If @error Or Not IsArray($aRet) Then Return False
	Return BitAND($aRet[0], 0x8000) <> 0
EndFunc   ;==>_IsKeyPressed

; ================================================================================================
; ZOOM SETTINGS MANAGEMENT FUNCTIONS
; ================================================================================================

; Gets the current state of security settings (enabled/disabled)
; @param $aSettings - Array of setting names to check
; @return Object - Dictionary containing setting states
Func GetSecuritySettingsState($aSettings)
	Local $oHostMenu = _OpenHostTools()
	If Not IsObj($oHostMenu) Then Return False

	; Create dictionary to store setting states
	Local $oDict = ObjCreate("Scripting.Dictionary")

	For $i = 0 To UBound($aSettings) - 1
		Local $sSetting = $aSettings[$i]

		; Find the setting menu item
		Local $oSetting = FindElementByPartialName($sSetting, Default, $oHostMenu)
		Local $bEnabled = False
		If IsObj($oSetting) Then
			; Read the setting's label to determine if it's checked
			Local $sLabel = ""
			$oSetting.GetCurrentPropertyValue($UIA_NamePropertyId, $sLabel)
			Local $sLabelLower = StringLower($sLabel)

			; Setting is enabled if NOT marked as "unchecked"
			$bEnabled = Not StringRegExp($sLabelLower, "\b" & StringLower(GetUserSetting("UncheckedValue")) & "\b")
		Else
			Debug(t("ERROR_SETTING_NOT_FOUND") & ": '" & $sSetting & "'", "ERROR")
		EndIf

		Debug("Setting '" & $sSetting & "' is currently " & ($bEnabled ? "ENABLED" : "DISABLED"), "DEBUG")
		$oDict.Add($sSetting, $bEnabled)
	Next

	_CloseHostTools()
	Return $oDict
EndFunc   ;==>GetSecuritySettingsState

; Sets a security setting to the desired state (enabled/disabled)
; @param $sSetting - Setting name to modify
; @param $bDesired - Desired state (True=enabled, False=disabled)
Func SetSecuritySetting($sSetting, $bDesired)
	Local $oHostMenu = _OpenHostTools()
	If Not IsObj($oHostMenu) Then Return False

	Local $oSetting = FindElementByPartialName($sSetting, Default, $oHostMenu)
	If Not IsObj($oSetting) Then Return

	; Check current state
	Local $sLabel
	$oSetting.GetCurrentPropertyValue($UIA_NamePropertyId, $sLabel)
	Local $sLabelLower = StringLower($sLabel)
	Local $bEnabled = Not StringRegExp($sLabelLower, "\b" & StringLower(GetUserSetting("UncheckedValue")) & "\b")

	Debug("Setting '" & $sLabel & "' | Current: " & $bEnabled & " | Desired: " & $bDesired, "DEBUG")

	; Only click if state needs to change
	If $bEnabled <> $bDesired Then
		_HoverElement($oSetting, 50)
		_MoveMouseToStartOfElement($oSetting, True) ; Click at start of element to ensure change
		Debug("Toggled setting '" & $sSetting & "'", "SETTING CHANGE")
	Else
		_CloseHostTools()
	EndIf
EndFunc   ;==>SetSecuritySetting

; Toggles host's audio or video feed on/off
; @param $feedType - "Video" or "Audio"
; @param $desiredState - True to enable, False to disable
Func ToggleFeed($feedType, $desiredState)
	Local $currentlyEnabled = False

	; Controls might be hidden, show them by moving the mouse
	Debug("Controls might be hidden. Moving mouse to show controls.", "DEBUG")
	_MoveMouseToStartOfElement($oZoomWindow)

	If $feedType = "Video" Then
		; Check for video control buttons to determine current state
		Local $stopMyVideoButton = FindElementByPartialName(GetUserSetting("StopVideoValue"), Default, $oZoomWindow)
		Local $startMyVideoButton = FindElementByPartialName(GetUserSetting("StartVideoValue"), Default, $oZoomWindow)
		$currentlyEnabled = IsObj($stopMyVideoButton)

		; Toggle if needed
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
		; Check for audio control buttons to determine current state
		Local $muteHostButton = FindElementByPartialName(GetUserSetting("CurrentlyUnmutedValue"), Default, $oZoomWindow)
		Local $unmuteHostButton = FindElementByPartialName(GetUserSetting("UnmuteAudioValue"), Default, $oZoomWindow)
		$currentlyEnabled = IsObj($muteHostButton)

		; Toggle if needed
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

; Mutes all meeting participants
; @return Boolean - True if successful, False otherwise
Func MuteAll()
	Debug("Attempting to 'Mute All' participants.")

	; Open participants panel
	Local $oParticipantsPanel = _OpenParticipantsPanel()
	If Not IsObj($oParticipantsPanel) Then Return False

	; Find and click "Mute All" button
	Local $oButton = FindElementByPartialName(GetUserSetting("MuteAllValue"), Default, $oZoomWindow)
	If Not _ClickElement($oButton) Then
		Debug("'Mute all' button not found.", "WARN")
		Return False
	EndIf

	; Confirm the action in dialog
	Return DialogClick("zChangeNameWndClass", GetUserSetting("YesValue"))
EndFunc   ;==>MuteAll

; Clicks a button in a dialog window by class name and button text
; @param $ClassName - Dialog window class name
; @param $ButtonLabel - Button text to click
; @return Boolean - True if successful, False otherwise
Func DialogClick($ClassName, $ButtonLabel)
	Local $oDialog = FindElementByClassName($ClassName)
	Local $oButton = FindElementByPartialName($ButtonLabel, Default, $oDialog)
	If _ClickElement($oButton) Then Return True
	Debug("Button '" & $ButtonLabel & "' not found.", "WARN")
	Return False
EndFunc   ;==>DialogClick

; Displays current status of Zoom security settings (for debugging)
; Func _GetZoomStatus()
; 	Local $aSettings[2] = ["Unmute themselves", "Share screen"]
; 	Local $oStates = GetSecuritySettingsState($aSettings)

; 	For $sKey In $oStates.Keys
; 		Debug($sKey & " = " & ($oStates.Item($sKey) ? "ENABLED" : "DISABLED"), "ZOOM STATUS")
; 	Next
; EndFunc   ;==>_GetZoomStatus

; ================================================================================================
; MEETING AUTOMATION FUNCTIONS
; ================================================================================================

; Launches Zoom with the configured meeting ID
; @return Boolean - True if successful, False otherwise
Func _LaunchZoom()
	Debug(t("INFO_ZOOM_LAUNCHING"), "INFO")

	Local $meetingID = GetUserSetting("MeetingID")
	If $meetingID = "" Then
		Debug(t("ERROR_MEETING_ID_NOT_CONFIGURED"), "ERROR")
		Return SetError(1, 0, 0)
	EndIf

	; Use Zoom URL protocol to launch meeting directly
	Local $zoomURL = "zoommtg://zoom.us/join?confno=" & $meetingID
	ShellExecute($zoomURL)
	Debug(t("INFO_ZOOM_LAUNCHED") & ": " & $meetingID, "INFO")
	ResponsiveSleep(10)  ; Wait for Zoom to launch

	_SnapZoomWindowToSide()

	_GetZoomWindow()

	Return IsObj($oZoomWindow)
EndFunc   ;==>_LaunchZoom

; Configures settings before and after meetings
; - Enables unmute permission for participants
; - Disables screen sharing permission
; - Turns off host audio and video
Func _SetPreAndPostMeetingSettings()
	Debug(t("INFO_CONFIG_BEFORE_AFTER_START"), "INFO")
	ShowOverlayMessage('PLEASE_WAIT')
	FocusZoomWindow()               ; Ensure Zoom window is focused
	SetSecuritySetting("Unmute", True)          ; Allow participants to unmute
	SetSecuritySetting("Share screen", False)   ; Prevent screen sharing
	ToggleFeed("Audio", False)                  ; Turn off host audio
	ToggleFeed("Video", False)                  ; Turn off host video
	; TODO: Unmute All function
	HideOverlayMessage()
	Debug(t("INFO_CONFIG_BEFORE_AFTER_DONE"), "INFO")
EndFunc   ;==>_SetPreAndPostMeetingSettings

; Configures settings during active meetings
; - Disables unmute permission (host controls audio)
; - Disables screen sharing permission
; - Mutes all participants
; - Turns on host audio and video
Func _SetDuringMeetingSettings()
	Debug(t("INFO_MEETING_STARTING_SOON_CONFIG"), "INFO")
	ShowOverlayMessage('PLEASE_WAIT')
	FocusZoomWindow()                          ; Ensure Zoom window is focused
	SetSecuritySetting("Unmute", False)         ; Prevent participant self-unmute
	SetSecuritySetting("Share screen", False)   ; Prevent screen sharing
	MuteAll()                                   ; Mute all participants
	ToggleFeed("Audio", True)                   ; Turn on host audio
	ToggleFeed("Video", True)                   ; Turn on host video
	HideOverlayMessage()
	Debug(t("INFO_CONFIG_DURING_MEETING_DONE"), "INFO")
EndFunc   ;==>_SetDuringMeetingSettings

; ================================================================================================
; MEETING TIMING AND MAIN LOGIC
; ================================================================================================

; Checks current time against meeting schedule and applies appropriate settings
; @param $meetingTime - Scheduled meeting time in HH:MM format
Func CheckMeetingWindow($meetingTime)
	If $meetingTime = "" Then Return

	Local $secondsToWait = 5     ; Default interval between checks

	; Parse meeting time
	Local $aParts = StringSplit($meetingTime, ":")
	Local $hour = Number($aParts[1])
	Local $min = Number($aParts[2])

	; Convert current time and meeting time to minutes for easier comparison
	Local $nowMin = Number(@HOUR) * 60 + Number(@MIN)
	Local $meetingMin = $hour * 60 + $min

	If $nowMin >= ($meetingMin - 60) And $nowMin < ($meetingMin - 1) Then
		; Pre-meeting window (1 hour before to 1 minute before)
		If Not $g_PrePostSettingsConfigured Then
			Local $zoomLaunched = _LaunchZoom()
			If Not $zoomLaunched Then
				Debug(t("ERROR_ZOOM_LAUNCH"), "ERROR")
			Else
				_SetPreAndPostMeetingSettings()
				$g_PrePostSettingsConfigured = True
			EndIf
		EndIf

	ElseIf $nowMin = ($meetingMin - 1) Then
		; Meeting start window (1 minute before meeting)
		If Not $g_DuringMeetingSettingsConfigured Then
			_GetZoomWindow()
			Local $zoomExists = IsObj($oZoomWindow)
			If Not $zoomExists Then
				Debug(t("ERROR_ZOOM_WINDOW_NOT_FOUND"), "ERROR")
			Else
				_SetDuringMeetingSettings()
				$g_DuringMeetingSettingsConfigured = True
			EndIf
		EndIf

	ElseIf $nowMin >= $meetingMin Then
		; Meeting already started
		Local $minutesAgo = $nowMin - $meetingMin
		If $minutesAgo <= 120 Then
			Debug(t("INFO_MEETING_STARTED_AGO", $minutesAgo), "INFO", $g_InitialNotificationWasShown)
		Else
			Debug(t("INFO_OUTSIDE_MEETING_WINDOW"), "INFO", $g_InitialNotificationWasShown)
		EndIf
		$secondsToWait = 30 ; Check every 30 seconds if meeting already started
	Else
		; Too early - show countdown to meeting
		Local $minutesLeft = $meetingMin - $nowMin
		Debug(t("INFO_MEETING_STARTING_IN", $minutesLeft), "INFO", $g_InitialNotificationWasShown)
	EndIf

	ResponsiveSleep($secondsToWait)
	$g_InitialNotificationWasShown = True
EndFunc   ;==>CheckMeetingWindow

; Load translations and configuration
_LoadTranslations()
LoadMeetingConfig()
_InitDayLabelMaps()

; Debugging functions here
; _LaunchZoom()
; _GetZoomWindow()
; FocusZoomWindow()
; _SetDuringMeetingSettings()

; Main application loop
While True
	; Handle tray icon events
	TrayEvent()

	; Check if day has changed to reset automation flags
	Global $today = @WDAY
	If $today <> $previousRunDay Then
		Debug("New day detected. Resetting configuration flags.", "DEBUG")
		$previousRunDay = $today
		$g_PrePostSettingsConfigured = False
		$g_DuringMeetingSettingsConfigured = False
	EndIf

	Global $timeNow = _NowTime(4) ; Get current time in HH:MM format

	; Check if today is a scheduled meeting day
	If $today = Number(GetUserSetting("MidweekDay")) Then
		CheckMeetingWindow(GetUserSetting("MidweekTime"))
	ElseIf $today = Number(GetUserSetting("WeekendDay")) Then
		CheckMeetingWindow(GetUserSetting("WeekendTime"))
	Else
		; Not a meeting day - wait 1 minute before checking again
		Debug(t("INFO_NO_MEETING_SCHEDULED"), "INFO", $g_InitialNotificationWasShown)
		$g_InitialNotificationWasShown = True
		ResponsiveSleep(60)
	EndIf
WEnd

; ================================================================================================
; ELEMENT NAME COLLECTION FUNCTIONS
; ================================================================================================

; Collects all element names (UIA_NamePropertyId) from specified windows using default control types
; @param $oZoomWindow - Zoom window element
; @param $oHostMenu - Host menu element (optional)
; @return Array - Array of unique element names, trimmed and sorted
Func GetElementNamesFromWindows($oZoomWindow, $oHostMenu = 0)
	Local $aNames = []

	If Not IsObj($oZoomWindow) Then
		Debug("GetElementNamesFromWindows: Invalid Zoom window object", "DEBUG")
		Return $aNames
	EndIf

	; Define control types to search (same as FindElementByPartialName default)
	Local $aControlTypes[2] = [$UIA_ButtonControlTypeId, $UIA_MenuItemControlTypeId]

	; Collect names from Zoom window
	_CollectElementNames($oZoomWindow, $aControlTypes, $aNames)

	Debug("Collected these values: " & $aNames & " from Zoom window", "DEBUG")

	; Collect names from Host menu if provided
	If IsObj($oHostMenu) Then
		_CollectElementNames($oHostMenu, $aControlTypes, $aNames)
	EndIf

	; Remove duplicates and sort
	$aNames = _ArrayUnique($aNames)
	_ArraySort($aNames)

	Debug("Collected " & UBound($aNames) & " unique element names", "DEBUG")
	Return $aNames
EndFunc   ;==>GetElementNamesFromWindows

; Helper function to collect element names from a parent element
; @param $oParent - Parent element to search within
; @param $aControlTypes - Array of control types to search
; @param ByRef $aNames - Array to store collected names
Func _CollectElementNames($oParent, $aControlTypes, ByRef $aNames)
	If Not IsObj($oParent) Then Return

	; Search each control type
	For $iType = 0 To UBound($aControlTypes) - 1
		Local $iControlType = $aControlTypes[$iType]

		; Create condition for this control type
		Local $pCondition
		$oUIAutomation.CreatePropertyCondition($UIA_ControlTypePropertyId, $iControlType, $pCondition)

		; Find all elements of this type
		Local $pElements
		$oParent.FindAll($TreeScope_Descendants, $pCondition, $pElements)

		Local $oElements = ObjCreateInterface($pElements, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
		If IsObj($oElements) Then
			Local $iCount
			$oElements.Length($iCount)

			; Extract name from each element
			For $i = 0 To $iCount - 1
				Local $pElement
				$oElements.GetElement($i, $pElement)

				Local $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
				If IsObj($oElement) Then
					Local $sName
					$oElement.GetCurrentPropertyValue($UIA_NamePropertyId, $sName)

					; Trim whitespace and add if not empty
					$sName = StringStripWS($sName, 3)
					If $sName <> "" Then
						_ArrayAdd($aNames, $sName)
					EndIf
				EndIf
			Next
		EndIf
	Next
EndFunc   ;==>_CollectElementNames

; ================================================================================================
; ELEMENT NAMES DISPLAY GUI
; ================================================================================================

; Global variables for element names display GUI
Global $g_ElementNamesGUI = 0
Global $g_ElementNamesEdit = 0

; Shows a closeable and moveable textarea with collected element names
; @param $aNames - Array of element names to display
Func ShowElementNamesGUI($aNames)
	; Close any existing GUI
	CloseElementNamesGUI()

	Local $iW = 500
	Local $iH = 400
	Local $iX = (@DesktopWidth - $iW) / 2
	Local $iY = (@DesktopHeight - $iH) / 2

	; Create moveable and closeable GUI
	$g_ElementNamesGUI = GUICreate("Zoom Element Names", $iW, $iH, $iX, $iY, $WS_CAPTION + $WS_SYSMENU + $WS_MINIMIZEBOX, $WS_EX_TOPMOST)
	GUISetOnEvent($GUI_EVENT_CLOSE, "CloseElementNamesGUI")

	; Create edit control for displaying element names
	Local $idEdit = GUICtrlCreateEdit(_ArrayToString($aNames, @CRLF), 10, 10, $iW - 20, $iH - 50, $ES_READONLY + $WS_VSCROLL + $WS_HSCROLL)
	GUICtrlSetFont($idEdit, 9, 400, 0, "Courier New") ; Monospace font for better readability
	$g_ElementNamesEdit = $idEdit

	; Create close button
	Local $idCloseBtn = GUICtrlCreateButton("Close", $iW - 80, $iH - 35, 70, 25)
	GUICtrlSetOnEvent($idCloseBtn, "CloseElementNamesGUI")

	; Make GUI moveable by dragging the title bar
	GUISetOnEvent($GUI_EVENT_PRIMARYDOWN, "_StartDragElementNamesGUI")

	GUISetState(@SW_SHOW, $g_ElementNamesGUI)

	Debug("Element names GUI displayed with " & UBound($aNames) & " names", "DEBUG")
EndFunc   ;==>ShowElementNamesGUI

; Closes the element names display GUI
Func CloseElementNamesGUI()
	If $g_ElementNamesGUI <> 0 Then
		GUIDelete($g_ElementNamesGUI)
		$g_ElementNamesGUI = 0
		$g_ElementNamesEdit = 0
		Debug("Element names GUI closed", "DEBUG")
	EndIf
EndFunc   ;==>CloseElementNamesGUI

; Handles dragging of the element names GUI
Func _StartDragElementNamesGUI()
	If $g_ElementNamesGUI = 0 Then Return

	; Get mouse position relative to GUI
	Local $mousePos = MouseGetPos()
	Local $guiPos = WinGetPos($g_ElementNamesGUI)

	Local $offsetX = $mousePos[0] - $guiPos[0]
	Local $offsetY = $mousePos[1] - $guiPos[1]

	; Drag while mouse is down
	While _IsMouseDown()
		$mousePos = MouseGetPos()
		WinMove($g_ElementNamesGUI, "", $mousePos[0] - $offsetX, $mousePos[1] - $offsetY)
		Sleep(10)
	WEnd
EndFunc   ;==>_StartDragElementNamesGUI

; Helper function to check if mouse button is down
Func _IsMouseDown()
	Local $aMousePos = MouseGetPos()
	Local $iState = DllCall("user32.dll", "int", "GetAsyncKeyState", "int", 0x01) ; VK_LBUTTON
	Return BitAND($iState[0], 0x8000) <> 0
EndFunc   ;==>_IsMouseDown

; ================================================================================================
; MAIN BUTTON HANDLER
; ================================================================================================

; Button handler for "Get Element Names" button
Func GetElementNames()
	Debug("Get Element Names button clicked", "DEBUG")

	; Check if Zoom meeting is in progress
	If Not FocusZoomWindow() Then
		; Meeting not in progress - show message
		MsgBox($MB_OK + $MB_ICONINFORMATION, "ZoomMate", "No active Zoom meeting found. Please ensure Zoom is open and a meeting is in progress.")
		Debug("No active Zoom meeting found", "WARN")
		Return
	EndIf

	; Meeting is in progress - collect element names
	Debug("Active Zoom meeting found, collecting element names...", "DEBUG")

	; Get Zoom window object
	_GetZoomWindow()

	If Not IsObj($oZoomWindow) Then
		MsgBox($MB_OK + $MB_ICONERROR, "ZoomMate", "Failed to get Zoom window object.")
		Debug("Failed to get Zoom window object", "DEBUG")
		Return
	EndIf

	; Open Host Tools menu to collect names from it too
	Local $oHostMenu = _OpenHostTools()
	If Not IsObj($oHostMenu) Then
		Debug("Failed to open Host Tools menu, collecting names from Zoom window only", "WARN")
	EndIf

	; Collect element names from both windows
	Local $aNames = GetElementNamesFromWindows($oZoomWindow, $oHostMenu)

	If UBound($aNames) = 0 Then
		MsgBox($MB_OK + $MB_ICONWARNING, "ZoomMate", "No element names found. This might indicate an issue with the UIAutomation interface.")
		Debug("No element names collected", "WARN")
		Return
	EndIf

	; Display the collected names in the GUI
	ShowElementNamesGUI($aNames)

	; Close Host Tools menu if it was opened
	If IsObj($oHostMenu) Then
		_CloseHostTools()
	EndIf

	Debug("Element names collection completed", "SUCCESS")
EndFunc   ;==>GetElementNames
