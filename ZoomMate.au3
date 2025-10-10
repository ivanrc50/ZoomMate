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
#include <GDIPlus.au3>
#include <WinAPI.au3>
#include <EditConstants.au3>
#include <File.au3>
#include "Includes\UIA_Functions-a.au3"
#include "Includes\CUIAutomation2.au3"

; ================================================================================================
; FILE INSTALLATION - Extract embedded images to script directory
; ================================================================================================
FileInstall("images\host_tools.jpg", @ScriptDir & "\images\host_tools.jpg", 1)
FileInstall("images\more_meeting_controls.jpg", @ScriptDir & "\images\more_meeting_controls.jpg", 1)
FileInstall("images\participant.jpg", @ScriptDir & "\images\participant.jpg", 1)
FileInstall("images\mute_all.jpg", @ScriptDir & "\images\mute_all.jpg", 1)
FileInstall("images\yes.jpg", @ScriptDir & "\images\yes.jpg", 1)
FileInstall("images\security_unmute.jpg", @ScriptDir & "\images\security_unmute.jpg", 1)
FileInstall("images\security_share_screen.jpg", @ScriptDir & "\images\security_share_screen.jpg", 1)
FileInstall("images\placeholder.jpg", @ScriptDir & "\images\placeholder.jpg", 1)
#include "Includes\i18n.au3"

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
; TIMING AND PERFORMANCE CONSTANTS
; ================================================================================================
Global Const $HOVER_DEFAULT_MS = 1000
Global Const $CLICK_DELAY_MS = 500
Global Const $WINDOW_SNAP_DELAY_MS = 500
Global Const $PRE_MEETING_MINUTES = 60
Global Const $MEETING_START_WARNING_MINUTES = 1
Global Const $SNAP_TOLERANCE_PX = 50
Global Const $CLICK_TIMEOUT_MS = 5000
Global Const $ELEMENT_SEARCH_RETRY_COUNT = 3
Global Const $ELEMENT_SEARCH_RETRY_DELAY_MS = 500
Global Const $UI_AUTOMATION_CACHE_TIME_MS = 2000

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
Global $g_ElementNamesSelectionGUI = 0             ; Handle for element names selection GUI
Global $g_ElementNamesSelectionList = 0            ; Handle for element names selection list
Global $g_ElementNamesSelectionResult = ""         ; Selected element name result
Global $g_ElementNamesSelectionCallback = ""       ; Callback function for selection
Global $g_ActiveFieldForLookup = 0                 ; Currently active field for lookup operations
Global $g_FieldLabels = ObjCreate("Scripting.Dictionary")  ; Maps field names to label control IDs

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
; UNICODE STRING HANDLING FUNCTIONS (Simplified)
; ================================================================================================

; Converts a Unicode string to UTF-8 bytes for INI file storage
; @param $sText - Unicode string to convert
; @return String - UTF-8 encoded string
Func _StringToUTF8($sText)
	Return BinaryToString(StringToBinary($sText, 4), 1)
EndFunc   ;==>_StringToUTF8

; Converts UTF-8 bytes from INI file back to Unicode string
; @param $sUTF8 - UTF-8 encoded string from INI file
; @return String - Unicode string
Func _UTF8ToString($sUTF8)
	Return BinaryToString(StringToBinary($sUTF8, 1), 4)
EndFunc   ;==>_UTF8ToString

; Translation lookup function with placeholder support
; @param $key - Translation key to look up
; @param $p0-$p2 - Optional placeholder values for {0}, {1}, {2} substitution
; @return String - Translated text with placeholders replaced
Func t($key, $p0 = Default, $p1 = Default, $p2 = Default)
	; Get the configured language from settings (fallback to English if not set)
	Local $currentLang = GetUserSetting("Language")
	If $currentLang = "" Then $currentLang = "en"

	; Get translations for the current language
	Local $translations = _GetLanguageTranslations($currentLang)

	If $translations.Exists($key) Then
		Local $s = $translations.Item($key)
		; Replace placeholders if provided
		If $p0 <> Default Then $s = StringReplace($s, "{0}", $p0, 0, $STR_CASESENSE)
		If $p1 <> Default Then $s = StringReplace($s, "{1}", $p1, 0, $STR_CASESENSE)
		If $p2 <> Default Then $s = StringReplace($s, "{2}", $p2, 0, $STR_CASESENSE)
		Return $s
	EndIf

	; Ultimate fallback: return the key itself
	Return $key
EndFunc   ;==>t



; Initializes day name to number mappings using translations
; Maps localized day names (DAY_1 through DAY_7) to numbers 1-7
Func _InitDayLabelMaps()
	; Clear existing mappings before reinitializing
	$g_DayLabelToNum.RemoveAll()
	$g_DayNumToLabel.RemoveAll()

	Local $i
	For $i = 1 To 7
		Local $key = "DAY_" & $i
		Local $label = t($key)
		If Not $g_DayLabelToNum.Exists($label) Then $g_DayLabelToNum.Add($label, $i)
		If Not $g_DayNumToLabel.Exists(String($i)) Then $g_DayNumToLabel.Add(String($i), $label)
		Debug("  " & $label & " -> " & $i, "VERBOSE")
	Next
	Debug("Day mappings initialized", "VERBOSE")
EndFunc   ;==>_InitDayLabelMaps

; ================================================================================================
; DEBUG AND STATUS FUNCTIONS
; ================================================================================================

; Debug logging and status update function with enhanced formatting and user notification
; @param $string - Message to log/display
; @param $type - Message type (DEBUG, INFO, ERROR, SUCCESS, VERBOSE, etc.)
; @param $noNotify - If True, suppress overlay notifications
; @param $isVerbose - If True, this is verbose debug logging (only shown in console)
; @param $functionName - Name of the calling function (optional, auto-detected if not provided)
Func Debug($string, $type = "VERBOSE", $noNotify = False, $isVerbose = False)

	If ($string) Then
		; Format timestamp for better log readability
		Local $timestamp = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC

		; Console logging with enhanced formatting
		If $isVerbose Then
			; Verbose debug logging - only to console, no user notification
			ConsoleWrite("[" & $timestamp & "] " & " VERBOSE: " & $string & @CRLF)
		Else
			; Standard logging - to console with type identification
			ConsoleWrite("[" & $timestamp & "] " & $type & ": " & $string & @CRLF)

			; Update status for important messages (non-verbose)
			If $type = "INFO" Or $type = "ERROR" Then
				$g_StatusMsg = $string

				; Show overlay notification for important messages (unless suppressed)
				If Not $noNotify Then
					Local $isError = ($type = "ERROR")
					ShowOverlayMessage($string, $isError, Not $isError)
				EndIf
			EndIf

			; Change tray icon for errors
			If $type = "ERROR" Then
				TraySetIcon($g_TrayIcon, 1) ; Error icon
			EndIf
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
	$g_UserSettings.Add("MeetingID", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomSettings", "MeetingID", "")))
	$g_UserSettings.Add("MidweekDay", _UTF8ToString(IniRead($CONFIG_FILE, "Meetings", "MidweekDay", "")))
	$g_UserSettings.Add("MidweekTime", _UTF8ToString(IniRead($CONFIG_FILE, "Meetings", "MidweekTime", "")))
	$g_UserSettings.Add("WeekendDay", _UTF8ToString(IniRead($CONFIG_FILE, "Meetings", "WeekendDay", "")))
	$g_UserSettings.Add("WeekendTime", _UTF8ToString(IniRead($CONFIG_FILE, "Meetings", "WeekendTime", "")))
	$g_UserSettings.Add("HostToolsValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "HostToolsValue", "")))
	$g_UserSettings.Add("ParticipantValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "ParticipantValue", "")))
	$g_UserSettings.Add("MuteAllValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "MuteAllValue", "")))
	$g_UserSettings.Add("MoreMeetingControlsValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "MoreMeetingControlsValue", "")))
	$g_UserSettings.Add("YesValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "YesValue", "")))
	$g_UserSettings.Add("UncheckedValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "UncheckedValue", "")))
	$g_UserSettings.Add("CurrentlyUnmutedValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "CurrentlyUnmutedValue", "")))
	$g_UserSettings.Add("UnmuteAudioValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "UnmuteAudioValue", "")))
	$g_UserSettings.Add("StopVideoValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "StopVideoValue", "")))
	$g_UserSettings.Add("StartVideoValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "StartVideoValue", "")))
	$g_UserSettings.Add("ZoomSecurityUnmuteValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "ZoomSecurityUnmuteValue", "")))
	$g_UserSettings.Add("ZoomSecurityShareScreenValue", _UTF8ToString(IniRead($CONFIG_FILE, "ZoomStrings", "ZoomSecurityShareScreenValue", "")))
	$g_UserSettings.Add("KeyboardShortcut", _UTF8ToString(IniRead($CONFIG_FILE, "General", "KeyboardShortcut", "")))

	; Window snapping preference (Disabled|Left|Right)
	$g_UserSettings.Add("SnapZoomSide", _UTF8ToString(IniRead($CONFIG_FILE, "General", "SnapZoomSide", "Disabled")))

	; Load language setting
	Local $lang = _UTF8ToString(IniRead($CONFIG_FILE, "General", "Language", ""))
	If $lang = "" Then
		$lang = "en"
		IniWrite($CONFIG_FILE, "General", "Language", _StringToUTF8($lang))
	EndIf
	$g_UserSettings.Add("Language", $lang)
	$g_CurrentLang = $lang

	; Load keyboard shortcut setting
	$g_KeyboardShortcut = GetUserSetting("KeyboardShortcut")
	If $g_KeyboardShortcut <> "" Then
		_UpdateKeyboardShortcut()
	EndIf

	; Check if all required settings are configured
	If GetUserSetting("MeetingID") = "" Or GetUserSetting("MidweekDay") = "" Or GetUserSetting("MidweekTime") = "" Or GetUserSetting("WeekendDay") = "" Or GetUserSetting("WeekendTime") = "" Or GetUserSetting("HostToolsValue") = "" Or GetUserSetting("ParticipantValue") = "" Or GetUserSetting("MuteAllValue") = "" Or GetUserSetting("YesValue") = "" Or GetUserSetting("MoreMeetingControlsValue") = "" Or GetUserSetting("UncheckedValue") = "" Or GetUserSetting("CurrentlyUnmutedValue") = "" Or GetUserSetting("UnmuteAudioValue") = "" Or GetUserSetting("StopVideoValue") = "" Or GetUserSetting("StartVideoValue") = "" Or GetUserSetting("ZoomSecurityUnmuteValue") = "" Or GetUserSetting("ZoomSecurityShareScreenValue") = "" Then
		; Open configuration GUI if any settings are missing
		ShowConfigGUI()
		While $g_ConfigGUI
			Sleep(100)
		WEnd
	Else
		Debug(t("INFO_CONFIG_LOADED"), "INFO")
		Debug("Midweek Meeting: " & t("DAY_" & GetUserSetting("MidweekDay")) & " at " & GetUserSetting("MidweekTime"), "VERBOSE", True)
		Debug("Weekend Meeting: " & t("DAY_" & GetUserSetting("WeekendDay")) & " at " & GetUserSetting("WeekendTime"), "VERBOSE", True)
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

	; Create main configuration window with initial estimated height
	Local $initialWidth = 640
	Local $initialHeight = 630
	$g_ConfigGUI = GUICreate(t("CONFIG_TITLE"), $initialWidth, $initialHeight)
	GUISetOnEvent($GUI_EVENT_CLOSE, "SaveConfigGUI", $g_ConfigGUI)

	Local $currentY = 10

	; ================================================================================================
	; SECTION 1: MEETING INFORMATION
	; ================================================================================================
	Local $idSection1 = _AddSectionHeader(t("SECTION_MEETING_INFO"), 10, $currentY)
	If Not $g_FieldLabels.Exists("Section1") Then $g_FieldLabels.Add("Section1", $idSection1)
	$currentY += 25

	; Meeting configuration fields
	_AddTextInputField("MeetingID", t("LABEL_MEETING_ID"), 10, $currentY, 300, $currentY, 200)
	$currentY += 30

	; Midweek meeting settings
	_AddDayDropdownField("MidweekDay", t("LABEL_MIDWEEK_DAY"), 10, $currentY, 300, $currentY, 200)
	$currentY += 30
	_AddTextInputField("MidweekTime", t("LABEL_MIDWEEK_TIME"), 10, $currentY, 300, $currentY, 200)
	$currentY += 40

	; Weekend meeting settings
	_AddDayDropdownField("WeekendDay", t("LABEL_WEEKEND_DAY"), 10, $currentY, 300, $currentY, 200)
	$currentY += 30
	_AddTextInputField("WeekendTime", t("LABEL_WEEKEND_TIME"), 10, $currentY, 300, $currentY, 200)
	$currentY += 40

	; Add separator line
	GUICtrlCreateLabel("", 10, $currentY, 620, 2)
	GUICtrlSetBkColor(-1, 0xCCCCCC)
	$currentY += 15

	; ================================================================================================
	; SECTION 2: ZOOM INTERFACE LABELS
	; ================================================================================================
	Local $idSection2 = _AddSectionHeader(t("SECTION_ZOOM_LABELS"), 10, $currentY)
	If Not $g_FieldLabels.Exists("Section2") Then $g_FieldLabels.Add("Section2", $idSection2)
	$currentY += 25

	; Zoom UI element text values (for internationalization support)
	_AddTextInputFieldWithTooltipAndLookup("HostToolsValue", t("LABEL_HOST_TOOLS"), 10, $currentY, 300, $currentY, 200, "LABEL_HOST_TOOLS_EXPLAIN", "host_tools.jpg")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("MoreMeetingControlsValue", t("LABEL_MORE_MEETING_CONTROLS"), 10, $currentY, 300, $currentY, 200, "LABEL_MORE_MEETING_CONTROLS_EXPLAIN", "more_meeting_controls.jpg")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("ParticipantValue", t("LABEL_PARTICIPANT"), 10, $currentY, 300, $currentY, 200, "LABEL_PARTICIPANT_EXPLAIN", "participant.jpg")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("MuteAllValue", t("LABEL_MUTE_ALL"), 10, $currentY, 300, $currentY, 200, "LABEL_MUTE_ALL_EXPLAIN", "mute_all.jpg")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("YesValue", t("LABEL_YES"), 10, $currentY, 300, $currentY, 200, "LABEL_YES_EXPLAIN", "yes.jpg")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("UncheckedValue", t("LABEL_UNCHECKED_VALUE"), 10, $currentY, 300, $currentY, 200, "LABEL_UNCHECKED_VALUE_EXPLAIN", "")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("CurrentlyUnmutedValue", t("LABEL_CURRENTLY_UNMUTED_VALUE"), 10, $currentY, 300, $currentY, 200, "LABEL_CURRENTLY_UNMUTED_VALUE_EXPLAIN", "")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("UnmuteAudioValue", t("LABEL_UNMUTE_AUDIO_VALUE"), 10, $currentY, 300, $currentY, 200, "LABEL_UNMUTE_AUDIO_VALUE_EXPLAIN", "")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("StopVideoValue", t("LABEL_STOP_VIDEO_VALUE"), 10, $currentY, 300, $currentY, 200, "LABEL_STOP_VIDEO_VALUE_EXPLAIN", "")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("StartVideoValue", t("LABEL_START_VIDEO_VALUE"), 10, $currentY, 300, $currentY, 200, "LABEL_START_VIDEO_VALUE_EXPLAIN", "")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("ZoomSecurityUnmuteValue", t("LABEL_ZOOM_SECURITY_UNMUTE"), 10, $currentY, 300, $currentY, 200, "LABEL_ZOOM_SECURITY_UNMUTE_EXPLAIN", "security_unmute.jpg")
	$currentY += 30
	_AddTextInputFieldWithTooltipAndLookup("ZoomSecurityShareScreenValue", t("LABEL_ZOOM_SECURITY_SHARE_SCREEN"), 10, $currentY, 300, $currentY, 200, "LABEL_ZOOM_SECURITY_SHARE_SCREEN_EXPLAIN", "security_share_screen.jpg")
	$currentY += 40

	; Add separator line
	GUICtrlCreateLabel("", 10, $currentY, 620, 2)
	GUICtrlSetBkColor(-1, 0xCCCCCC)
	$currentY += 15

	; ================================================================================================
	; SECTION 3: GENERAL SETTINGS
	; ================================================================================================
	Local $idSection3 = _AddSectionHeader(t("SECTION_GENERAL_SETTINGS"), 10, $currentY)
	If Not $g_FieldLabels.Exists("Section3") Then $g_FieldLabels.Add("Section3", $idSection3)
	$currentY += 25

	; Language selection dropdown
	Local $idLanguageLabel = GUICtrlCreateLabel(t("LABEL_LANGUAGE"), 10, $currentY, 200, 20)
	$idLanguagePicker = GUICtrlCreateCombo("", 300, $currentY, 200, 20)
	If Not $g_FieldLabels.Exists("Language") Then $g_FieldLabels.Add("Language", $idLanguageLabel)
	; Ensure translations are initialized before getting language list
	If $translations.Count = 0 Then _InitializeTranslations()
	Local $langList = _ListAvailableLanguageNames()
	Local $currentLang = GetUserSetting("Language")
	If $currentLang = "" Then $currentLang = "en"
	Local $currentDisplay = _GetLanguageDisplayName($currentLang)
	GUICtrlSetData($idLanguagePicker, $langList, $currentDisplay)
	GUICtrlSetOnEvent($idLanguagePicker, "_OnLanguageChanged") ; Handle language change
	$currentY += 30

	; Snap Zoom to side (Disabled|Left|Right)
	Local $idSnapLabel = GUICtrlCreateLabel(t("LABEL_SNAP_ZOOM_TO"), 10, $currentY, 200, 20)
	Global $idSnapZoom = GUICtrlCreateCombo("", 300, $currentY, 200, 20)
	If Not $g_FieldLabels.Exists("SnapZoomSide") Then $g_FieldLabels.Add("SnapZoomSide", $idSnapLabel)
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
	$currentY += 40

	_AddTextInputFieldWithTooltip("KeyboardShortcut", t("LABEL_KEYBOARD_SHORTCUT"), 10, $currentY, 300, $currentY, 200, "LABEL_KEYBOARD_SHORTCUT_EXPLAIN", '')
	$currentY += 40

	; Error display area (wider to match new GUI width)
	$g_ErrorAreaLabel = GUICtrlCreateLabel("", 10, $currentY, 620, 20)
	GUICtrlSetColor($g_ErrorAreaLabel, 0xFF0000) ; Red text for errors
	$currentY += 30

	; Action buttons (adjusted for wider GUI)
	Global $idSaveBtn = GUICtrlCreateButton(t("BTN_SAVE"), 10, $currentY, 100, 30)
	Global $idQuitBtn = GUICtrlCreateButton(t("BTN_QUIT"), 120, $currentY, 100, 30)
	$currentY += 30


	; Set initial button states
	GUICtrlSetState($idSaveBtn, $GUI_DISABLE)  ; Disabled until all fields valid
	GUICtrlSetState($idQuitBtn, $GUI_ENABLE)

	; Set button event handlers
	GUICtrlSetOnEvent($idSaveBtn, "SaveConfigGUI")
	GUICtrlSetOnEvent($idQuitBtn, "QuitApp")

	; ================================================================================================
	; DYNAMIC HEIGHT CALCULATION AND GUI RESIZING
	; ================================================================================================

	; Calculate required height based on content
	Local $buttonHeight = 30
	Local $buttonMargin = 20
	Local $requiredHeight = $currentY + $buttonHeight + $buttonMargin

	; Resize the GUI to fit the content exactly
	Local $aPos = WinGetPos($g_ConfigGUI)
	If $aPos[3] <> $requiredHeight Then
		WinMove($g_ConfigGUI, "", $aPos[0], $aPos[1], $initialWidth, $requiredHeight)
	EndIf

	; Perform initial validation check
	CheckConfigFields()

	; Show the GUI and register message handler for real-time validation
	GUISetState(@SW_SHOW, $g_ConfigGUI)
	; Shows a "Please Wait" message dialog during long operations
	GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND_EditChange")
EndFunc   ;==>ShowConfigGUI

; Helper function to return the maximum of two values
; @param $a - First value
; @param $b - Second value
; @return The maximum of the two values
Func _Max($a, $b)
	if $a = "" Then Return $b
	if $b = "" Then Return $a
	if $a = "" And $b = "" Then Return ""
	if Not IsNumber($a) Then Return $b
	if Not IsNumber($b) Then Return $a
	Return ($a > $b) ? $a : $b
EndFunc

; Immediately saves a specific field value to settings and INI file
; @param $key - Settings key name
; @param $value - Value to save
Func SaveFieldImmediately($key, $value)
	; Save to in-memory settings
	$g_UserSettings.Add($key, $value)

	; Save to INI file
	IniWrite($CONFIG_FILE, _GetIniSectionForKey($key), $key, _StringToUTF8($value))

	Debug("Immediately saved field " & $key & " with value: " & $value, "VERBOSE")
EndFunc   ;==>SaveFieldImmediately

; Handler for immediate saving of specific fields when they change
Func _OnImmediateSaveFieldChange()
	Local $idChanged = @GUI_CtrlId

	; Check if this is one of the fields that should save immediately
	If $idChanged = $g_FieldCtrls.Item("HostToolsValue") Or $idChanged = $g_FieldCtrls.Item("MoreMeetingControlsValue") Then
		Local $fieldKey = ""
		For $sKey In $g_FieldCtrls.Keys
			If $g_FieldCtrls.Item($sKey) = $idChanged Then
				$fieldKey = $sKey
				ExitLoop
			EndIf
		Next

		If $fieldKey <> "" Then
			Local $value = StringStripWS(GUICtrlRead($idChanged), 3)
			If $value <> "" Then
				SaveFieldImmediately($fieldKey, $value)
				Debug("Immediately saved " & $fieldKey & " with value: " & $value, "VERBOSE")
			EndIf
		EndIf
	EndIf
	
EndFunc   ;==>_OnImmediateSaveFieldChange

; Handler for when language selection changes - refreshes all GUI labels immediately
Func _OnLanguageChanged()
	Local $selDisplay = GUICtrlRead($idLanguagePicker)
	Local $selLang = _GetLanguageCodeFromDisplayName($selDisplay)
	If $selLang = "" Then $selLang = "en"  ; Fallback to English if not found

	; Update current language setting
	$g_UserSettings.Item("Language") = $selLang
	IniWrite($CONFIG_FILE, "General", "Language", _StringToUTF8($selLang))
	$g_CurrentLang = $selLang

	; Refresh day labels for new language
	_InitDayLabelMaps()

	; Refresh all GUI labels and field values
	_RefreshGUILabels()
	
EndFunc   ;==>_OnLanguageChanged

; Refreshes all GUI labels and field values when language changes
Func _RefreshGUILabels()
	If $g_ConfigGUI = 0 Then Return

	; Update window title
	WinSetTitle($g_ConfigGUI, "", t("CONFIG_TITLE"))

	; Update section headers (we need to find them by position or recreate)
	; For now, we'll focus on updating the key controls we can identify

	; Update language picker data with new language list
	Local $langList = _ListAvailableLanguageNames()
	Local $currentDisplay = _GetLanguageDisplayName($g_CurrentLang)
	GUICtrlSetData($idLanguagePicker, $langList, $currentDisplay)

	; Update snap zoom combo box data
	Local $snapVal = GetUserSetting("SnapZoomSide")
	Local $snapDisplay = t("SNAP_DISABLED")
	If $snapVal = "Left" Then
		$snapDisplay = t("SNAP_LEFT")
	ElseIf $snapVal = "Right" Then
		$snapDisplay = t("SNAP_RIGHT")
	EndIf
	GUICtrlSetData($idSnapZoom, t("SNAP_DISABLED") & "|" & t("SNAP_LEFT") & "|" & t("SNAP_RIGHT"), $snapDisplay)

	; Update button labels
	GUICtrlSetData($idSaveBtn, t("BTN_SAVE"))
	GUICtrlSetData($idQuitBtn, t("BTN_QUIT"))

	; Update day dropdowns with new language labels
	_RefreshDayDropdowns()

	; Update field labels for text inputs (this is more complex, would need to track label IDs)
	; For now, we'll update the main visible ones

	; Update field labels using the stored label control IDs
	For $sKey In $g_FieldLabels.Keys
		Local $labelCtrl = $g_FieldLabels.Item($sKey)
		If $labelCtrl <> 0 Then
			; For now, we'll update specific known labels
			Switch $sKey
				Case "MeetingID"
					GUICtrlSetData($labelCtrl, t("LABEL_MEETING_ID"))
				Case "MidweekDay"
					GUICtrlSetData($labelCtrl, t("LABEL_MIDWEEK_DAY"))
				Case "MidweekTime"
					GUICtrlSetData($labelCtrl, t("LABEL_MIDWEEK_TIME"))
				Case "WeekendDay"
					GUICtrlSetData($labelCtrl, t("LABEL_WEEKEND_DAY"))
				Case "WeekendTime"
					GUICtrlSetData($labelCtrl, t("LABEL_WEEKEND_TIME"))
				Case "HostToolsValue"
					GUICtrlSetData($labelCtrl, t("LABEL_HOST_TOOLS"))
				Case "MoreMeetingControlsValue"
					GUICtrlSetData($labelCtrl, t("LABEL_MORE_MEETING_CONTROLS"))
				Case "ParticipantValue"
					GUICtrlSetData($labelCtrl, t("LABEL_PARTICIPANT"))
				Case "MuteAllValue"
					GUICtrlSetData($labelCtrl, t("LABEL_MUTE_ALL"))
				Case "YesValue"
					GUICtrlSetData($labelCtrl, t("LABEL_YES"))
				Case "UncheckedValue"
					GUICtrlSetData($labelCtrl, t("LABEL_UNCHECKED_VALUE"))
				Case "CurrentlyUnmutedValue"
					GUICtrlSetData($labelCtrl, t("LABEL_CURRENTLY_UNMUTED_VALUE"))
				Case "UnmuteAudioValue"
					GUICtrlSetData($labelCtrl, t("LABEL_UNMUTE_AUDIO_VALUE"))
				Case "StopVideoValue"
					GUICtrlSetData($labelCtrl, t("LABEL_STOP_VIDEO_VALUE"))
				Case "StartVideoValue"
					GUICtrlSetData($labelCtrl, t("LABEL_START_VIDEO_VALUE"))
				Case "ZoomSecurityUnmuteValue"
					GUICtrlSetData($labelCtrl, t("LABEL_ZOOM_SECURITY_UNMUTE"))
				Case "ZoomSecurityShareScreenValue"
					GUICtrlSetData($labelCtrl, t("LABEL_ZOOM_SECURITY_SHARE_SCREEN"))
				Case "KeyboardShortcut"
					GUICtrlSetData($labelCtrl, t("LABEL_KEYBOARD_SHORTCUT"))
				Case "Language"
					GUICtrlSetData($labelCtrl, t("LABEL_LANGUAGE"))
				Case "SnapZoomSide"
					GUICtrlSetData($labelCtrl, t("LABEL_SNAP_ZOOM_TO"))
				Case "Section1"
					GUICtrlSetData($labelCtrl, t("SECTION_MEETING_INFO"))
				Case "Section2"
					GUICtrlSetData($labelCtrl, t("SECTION_ZOOM_LABELS"))
				Case "Section3"
					GUICtrlSetData($labelCtrl, t("SECTION_GENERAL_SETTINGS"))
			EndSwitch
		EndIf
	Next

	; Clear and rebuild error area
	If $g_ErrorAreaLabel <> 0 Then
		GUICtrlSetData($g_ErrorAreaLabel, "")
	EndIf

	; Re-run validation to update any error messages and button states
	CheckConfigFields()
	
EndFunc   ;==>_RefreshGUILabels

; Refreshes day dropdown controls with new language labels
Func _RefreshDayDropdowns()
	; Update MidweekDay dropdown
	Local $midweekCtrl = $g_FieldCtrls.Item("MidweekDay")
	If $midweekCtrl <> 0 Then
		Local $currentMidweekNum = String(GetUserSetting("MidweekDay"))
		Local $currentMidweekLabel = $currentMidweekNum
		If $g_DayNumToLabel.Exists($currentMidweekNum) Then
			$currentMidweekLabel = $g_DayNumToLabel.Item($currentMidweekNum)
		EndIf

		Local $dayList = ""
		For $i = 2 To 7  ; Monday through Saturday
			Local $lbl = t("DAY_" & $i)
			$dayList &= ($dayList = "" ? $lbl : "|" & $lbl)
		Next
		Local $lblSun = t("DAY_" & 1)  ; Sunday last
		$dayList &= ($dayList = "" ? $lblSun : "|" & $lblSun)

		GUICtrlSetData($midweekCtrl, $dayList, $currentMidweekLabel)
	EndIf

	; Update WeekendDay dropdown
	Local $weekendCtrl = $g_FieldCtrls.Item("WeekendDay")
	If $weekendCtrl <> 0 Then
		Local $currentWeekendNum = String(GetUserSetting("WeekendDay"))
		Local $currentWeekendLabel = $currentWeekendNum
		If $g_DayNumToLabel.Exists($currentWeekendNum) Then
			$currentWeekendLabel = $g_DayNumToLabel.Item($currentWeekendNum)
		EndIf

		Local $dayList = ""
		For $i = 2 To 7  ; Monday through Saturday
			Local $lbl = t("DAY_" & $i)
			$dayList &= ($dayList = "" ? $lbl : "|" & $lbl)
		Next
		Local $lblSun = t("DAY_" & 1)  ; Sunday last
		$dayList &= ($dayList = "" ? $lblSun : "|" & $lblSun)

		GUICtrlSetData($weekendCtrl, $dayList, $currentWeekendLabel)
	EndIf
EndFunc   ;==>_RefreshDayDropdowns


; Helper function to add a section header with styling
; @param $text - Header text to display
; @param $x - X position
; @param $y - Y position
Func _AddSectionHeader($text, $x, $y)
	Local $idLabel = GUICtrlCreateLabel($text, $x, $y, 620, 20)
	GUICtrlSetFont($idLabel, 10, 700, Default, "Segoe UI") ; Bold, larger font
	GUICtrlSetColor($idLabel, 0x0066CC) ; Blue color
	GUICtrlSetBkColor($idLabel, 0xE8F4FD) ; Light blue background
	Return $idLabel
EndFunc   ;==>_AddSectionHeader
; Shows a message dialog during long operations with i18n support and enhanced styling for errors vs info messages
; @param $messageType - Type of message to show ('PLEASE_WAIT', 'POST_MEETING_HIT_KEY', or custom error/info messages)
; @param $isError - Boolean indicating if this is an error message (red background, requires click to dismiss)
; @param $autoDismiss - Boolean indicating if the message should auto-dismiss (default true for info, false for errors)
Func ShowOverlayMessage($messageType = 'PLEASE_WAIT', $isError = False, $autoDismiss = True)
	; If already showing, just update it
	If $g_OverlayMessageGUI <> 0 Then
		; Update existing GUI with new message
		Local $text = t($messageType)
		WinSetTitle($g_OverlayMessageGUI, '', '')
		Local $idLblExisting = _GetOverlayMessageLabelControl()
		If $idLblExisting <> 0 Then
			; Use GUICtrl functions since we have the control ID
			GUICtrlSetData($idLblExisting, $text)
			; Update font styling to use default Windows system font
			GUICtrlSetFont($idLblExisting, 14, 700, Default, "Segoe UI")
			; Update background color based on error state
			If $isError Then
				GUICtrlSetBkColor($idLblExisting, 0xFFE6E6) ; Light red for errors
				GUICtrlSetColor($idLblExisting, 0xCC0000) ; Dark red text for errors
			Else
				GUICtrlSetBkColor($idLblExisting, 0xE6F3FF) ; Light blue for info
				GUICtrlSetColor($idLblExisting, 0x0066CC) ; Blue text for info
			EndIf
		EndIf
		GUISetState(@SW_SHOW, $g_OverlayMessageGUI)
		WinSetOnTop(HWnd($g_OverlayMessageGUI), '', $WINDOWS_ONTOP)
		; Set up auto-dismiss timer for non-error messages
		If Not $isError And $autoDismiss Then
			AdlibRegister("HideOverlayMessage", 5000) ; Auto-dismiss after 5 seconds for info messages
		EndIf
		Return
	EndIf

	Local $iW = 350
	Local $iH = 140
	Local $iX = (@DesktopWidth - $iW) / 2
	Local $iY = (@DesktopHeight - $iH) / 2

	; Create borderless, always-on-top popup on primary monitor
	$g_OverlayMessageGUI = GUICreate(t($messageType), $iW, $iH, $iX, $iY, $WS_POPUP, $WS_EX_TOPMOST)

	; Set background color based on error state
	Local $bgColor = ($isError ? 0xFFE6E6 : 0xE6F3FF) ; Light red for errors, light blue for info
	GUISetBkColor($bgColor, $g_OverlayMessageGUI)

	; Create a label that supports both centering and word wrapping
	Local $idLbl = GUICtrlCreateLabel(t($messageType), 10, 10, $iW - 20, $iH - 20, $SS_CENTER)
	; Set text color based on error state
	Local $textColor = ($isError ? 0xCC0000 : 0x0066CC) ; Dark red for errors, blue for info
	GUICtrlSetColor($idLbl, $textColor)
	GUICtrlSetFont($idLbl, 14, 700, Default, "Segoe UI") ; Use default Windows system font

	; Make the label clickable for dismissal (especially for error messages)
	GUICtrlSetCursor($idLbl, 0) ; Hand cursor
	GUICtrlSetOnEvent($idLbl, "HideOverlayMessage")

	GUISetState(@SW_SHOW, $g_OverlayMessageGUI)
	WinSetOnTop(HWnd($g_OverlayMessageGUI), '', $WINDOWS_ONTOP)

	; Set up auto-dismiss timer for non-error messages
	If Not $isError And $autoDismiss Then
		AdlibRegister("HideOverlayMessage", 3000) ; Auto-dismiss after 3 seconds for info messages
	EndIf
EndFunc   ;==>ShowOverlayMessage

; Hides and destroys the "Please Wait" message dialog
Func HideOverlayMessage()
	If $g_OverlayMessageGUI <> 0 Then
		; Unregister any auto-dismiss timer
		AdlibUnregister("HideOverlayMessage")
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
	Local $idLabel = GUICtrlCreateLabel($label, $xLabel, $yLabel, 180, 20)
	Local $idInput = GUICtrlCreateInput(GetUserSetting($key), $xInput, $yInput, $wInput, 20)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idInput)
	If Not $g_FieldLabels.Exists($key) Then $g_FieldLabels.Add($key, $idLabel)
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
	Local $idLabel = GUICtrlCreateLabel($label, $xLabel, $yLabel, 200, 20)

	; Only create info icon if an image name is provided
	If $imageName <> "" Then
		; Build tooltip text with explanation and image reference
		Local $tooltipText = t($explainKey)
		Local $imagePath = @ScriptDir & "\images\" & $imageName

		; Check if image exists, if not use placeholder path
		If Not FileExists($imagePath) Then
			$imagePath = @ScriptDir & "\images\placeholder.jpg"
		EndIf

		; Get image dimensions for proper sizing
		Local $iImageWidth = 20, $iImageHeight = 20 ; Default fallback
		If FileExists($imagePath) Then
			_GDIPlus_Startup()
			Local $hImage = _GDIPlus_ImageLoadFromFile($imagePath)
			If $hImage <> 0 Then
				$iImageWidth = _GDIPlus_ImageGetWidth($hImage)
				$iImageHeight = _GDIPlus_ImageGetHeight($hImage)
				_GDIPlus_ImageDispose($hImage)

				; Scale down if too large (max 40px for icon)
				If $iImageWidth > 40 Or $iImageHeight > 40 Then
					Local $scale = 40 / _Max($iImageWidth, $iImageHeight)
					$iImageWidth *= $scale
					$iImageHeight *= $scale
				EndIf
			EndIf
			_GDIPlus_Shutdown()
		EndIf

		; Create info icon label with dynamic size
		Local $idInfoIcon = GUICtrlCreateLabel("[?]", $xLabel + 205, $yLabel, $iImageWidth, $iImageHeight)
		GUICtrlSetColor($idInfoIcon, 0x0066CC) ; Blue color
		GUICtrlSetFont($idInfoIcon, 10, 700) ; Bold
		GUICtrlSetCursor($idInfoIcon, 0) ; Hand cursor

		; Set standard tooltip with explanation text only
		GUICtrlSetTip($idInfoIcon, $tooltipText)

		; Store image path for custom tooltip display on click
		If Not $g_InfoIconData.Exists($idInfoIcon) Then
			$g_InfoIconData.Add($idInfoIcon, $imagePath)
		EndIf

		; Set click event to show image tooltip
		GUICtrlSetOnEvent($idInfoIcon, "_ShowImageTooltip")
	EndIf

	Local $idInput = GUICtrlCreateInput(GetUserSetting($key), $xInput, $yInput, $wInput, 20)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idInput)
	If Not $g_FieldLabels.Exists($key) Then $g_FieldLabels.Add($key, $idLabel)
	GUICtrlSetOnEvent($idInput, "CheckConfigFields")
EndFunc   ;==>_AddTextInputFieldWithTooltip

; Helper function to add text input field with label, info icon, tooltip, and lookup button
; @param $key - Settings key name
; @param $label - Display label text
; @param $xLabel,$yLabel - Label position
; @param $xInput,$yInput - Input position
; @param $wInput - Input width
; @param $explainKey - Translation key for explanation text
; @param $imageName - Image filename for tooltip
Func _AddTextInputFieldWithTooltipAndLookup($key, $label, $xLabel, $yLabel, $xInput, $yInput, $wInput, $explainKey, $imageName)
	Local $idLabel = GUICtrlCreateLabel($label, $xLabel, $yLabel, 200, 20)

	; Only create info icon if an image name is provided
	If $imageName <> "" Then
		; Build tooltip text with explanation and image reference
		Local $tooltipText = t($explainKey)
		Local $imagePath = @ScriptDir & "\images\" & $imageName

		; Check if image exists, if not use placeholder path
		If Not FileExists($imagePath) Then
			$imagePath = @ScriptDir & "\images\placeholder.jpg"
		EndIf

		; Get image dimensions for proper sizing
		Local $iImageWidth = 20, $iImageHeight = 20 ; Default fallback
		If FileExists($imagePath) Then
			_GDIPlus_Startup()
			Local $hImage = _GDIPlus_ImageLoadFromFile($imagePath)
			If $hImage <> 0 Then
				$iImageWidth = _GDIPlus_ImageGetWidth($hImage)
				$iImageHeight = _GDIPlus_ImageGetHeight($hImage)
				_GDIPlus_ImageDispose($hImage)

				; Scale down if too large (max 40px for icon)
				If $iImageWidth > 40 Or $iImageHeight > 40 Then
					Local $scale = 40 / _Max($iImageWidth, $iImageHeight)
					$iImageWidth *= $scale
					$iImageHeight *= $scale
				EndIf
			EndIf
			_GDIPlus_Shutdown()
		EndIf

		; Create info icon label with dynamic size
		Local $idInfoIcon = GUICtrlCreateLabel("[?]", $xLabel + 205, $yLabel, $iImageWidth, $iImageHeight)
		GUICtrlSetColor($idInfoIcon, 0x0066CC) ; Blue color
		GUICtrlSetFont($idInfoIcon, 10, 700) ; Bold
		GUICtrlSetCursor($idInfoIcon, 0) ; Hand cursor

		; Set standard tooltip with explanation text only
		GUICtrlSetTip($idInfoIcon, $tooltipText)

		; Store image path for custom tooltip display on click
		If Not $g_InfoIconData.Exists($idInfoIcon) Then
			$g_InfoIconData.Add($idInfoIcon, $imagePath)
		EndIf

		; Set click event to show image tooltip
		GUICtrlSetOnEvent($idInfoIcon, "_ShowImageTooltip")
	EndIf

	; Create text input field (slightly smaller width to make room for lookup button)
	Local $idInput = GUICtrlCreateInput(GetUserSetting($key), $xInput, $yInput, $wInput - 35, 20)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idInput)
	If Not $g_FieldLabels.Exists($key) Then $g_FieldLabels.Add($key, $idLabel)
	GUICtrlSetOnEvent($idInput, "CheckConfigFields")
	GUICtrlSetOnEvent($idInput, "_OnFieldFocus") ; Track when field gets focus

	; For HostToolsValue and MoreMeetingControlsValue, add immediate save functionality
	If $key = "HostToolsValue" Or $key = "MoreMeetingControlsValue" Then
		GUICtrlSetOnEvent($idInput, "_OnImmediateSaveFieldChange")
	EndIf

	_CreateLookupButton($xInput, $yInput, $wInput, $idInput)
EndFunc   ;==>_AddTextInputFieldWithTooltipAndLookup

; Creates a lookup button with magnifying glass icon for text input fields
; @param $xInput - Input field X position
; @param $yInput - Input field Y position
; @param $wInput - Input field width
; @param $idInput - Input field control ID
; @return Integer - Lookup button control ID
Func _CreateLookupButton($xInput, $yInput, $wInput, $idInput)
	Local $idLookupBtn = GUICtrlCreateButton("...", $xInput + $wInput - 30, $yInput, 25, 20)
	GUICtrlSetFont($idLookupBtn, 8, 400)
	GUICtrlSetTip($idLookupBtn, "Lookup element names from Zoom")
	GUICtrlSetOnEvent($idLookupBtn, "_OnLookupButtonClick")

	; Store the relationship between lookup button and input field
	If Not $g_InfoIconData.Exists($idLookupBtn) Then
		$g_InfoIconData.Add($idLookupBtn, $idInput)
	EndIf

	Return $idLookupBtn
EndFunc   ;==>_CreateLookupButton


; Handler for when an input field gets focus
Func _OnFieldFocus()
	$g_ActiveFieldForLookup = @GUI_CtrlId
	Debug("Field focused: " & @GUI_CtrlId, "VERBOSE")
EndFunc   ;==>_OnFieldFocus

; Handler for lookup button click
Func _OnLookupButtonClick()
	; Get the control ID of the button that was clicked
	Local $idButton = @GUI_CtrlId

	; Get the corresponding input field from our stored mapping
	Local $idInputField = 0
	If $g_InfoIconData.Exists($idButton) Then
		$idInputField = $g_InfoIconData.Item($idButton)
		$g_ActiveFieldForLookup = $idInputField ; Set this as the active field
	EndIf

	; Collect element names and show the selection GUI
	GetElementNamesForField()
EndFunc   ;==>_OnLookupButtonClick

; Collects element names and shows selection GUI for field population
Func GetElementNamesForField()
	Debug("Lookup button clicked - collecting element names for field", "VERBOSE")

	; Check if Zoom meeting is in progress
	If Not FocusZoomWindow() Then
		Return
	EndIf

	; Meeting is in progress - collect element names
	Debug("Active Zoom meeting found, collecting element names...", "VERBOSE")

	; Get Zoom window object
	if not _GetZoomWindow() then return

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

	; Show selection GUI with callback to populate the active field
	ShowElementNamesSelectionGUI($aNames, "OnFieldElementSelected")

	; Close Host Tools menu if it was opened
	If IsObj($oHostMenu) Then
		_CloseHostTools()
	EndIf

	Debug("Element names collection completed for field lookup", "VERBOSE")
EndFunc   ;==>GetElementNamesForField

; Callback function when user selects an element name for a field
Func OnFieldElementSelected($selectedName)
	; Populate the currently active field with the selected name
	If $g_ActiveFieldForLookup <> 0 Then
		GUICtrlSetData($g_ActiveFieldForLookup, $selectedName)
		; Trigger validation check
		CheckConfigFields()
		Debug("Populated field " & $g_ActiveFieldForLookup & " with: " & $selectedName, "VERBOSE")
	Else
		Debug("No active field to populate", "WARN")
	EndIf
EndFunc   ;==>OnFieldElementSelected

; Helper function to add day selection dropdown with label
; @param $key - Settings key name
; @param $label - Display label text
; @param $xLabel,$yLabel - Label position
; @param $xInput,$yInput - Input position
; @param $wInput - Input width
Func _AddDayDropdownField($key, $label, $xLabel, $yLabel, $xInput, $yInput, $wInput)
	Local $idLabel = GUICtrlCreateLabel($label, $xLabel, $yLabel, 180, 20)

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
	Debug("Day dropdown " & $key & ": Raw setting value = '" & $currentNum & "'", "VERBOSE")
	If $g_DayNumToLabel.Exists($currentNum) Then
		$currentLabel = $g_DayNumToLabel.Item($currentNum)
		Debug("Day dropdown " & $key & ": Setting selection to '" & $currentLabel & "' (day " & $currentNum & ")", "VERBOSE")
	Else
		Debug("Day dropdown " & $key & ": Day number " & $currentNum & " not found in translation map", "WARN")
		Debug("Day dropdown " & $key & ": Available translations: " & $g_DayNumToLabel.Count & " items", "VERBOSE")
	EndIf

	Local $idCombo = GUICtrlCreateCombo("", $xInput, $yInput, $wInput, 20)
	GUICtrlSetData($idCombo, $dayList, $currentLabel)
	If Not $g_FieldCtrls.Exists($key) Then $g_FieldCtrls.Add($key, $idCombo)
	If Not $g_FieldLabels.Exists($key) Then $g_FieldLabels.Add($key, $idLabel)
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
		IniWrite($CONFIG_FILE, _GetIniSectionForKey($sKey), $sKey, _StringToUTF8($val))
	Next

	; Save language setting
	Local $selDisplay = GUICtrlRead($idLanguagePicker)
	Local $selLang = _GetLanguageCodeFromDisplayName($selDisplay)
	If $selLang = "" Then $selLang = "en"  ; Fallback to English if not found
	$g_UserSettings.Item("Language") = $selLang
	IniWrite($CONFIG_FILE, "General", "Language", _StringToUTF8($selLang))
	$g_CurrentLang = $selLang

	; Save keyboard shortcut setting
	IniWrite($CONFIG_FILE, "General", "KeyboardShortcut", _StringToUTF8(GetUserSetting("KeyboardShortcut")))
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

	; Get image dimensions for proper tooltip sizing
	Local $iImageWidth = 200, $iImageHeight = 150 ; Default fallback
	If FileExists($imagePath) Then
		_GDIPlus_Startup()
		Local $hImage = _GDIPlus_ImageLoadFromFile($imagePath)
		If $hImage <> 0 Then
			$iImageWidth = _GDIPlus_ImageGetWidth($hImage)
			$iImageHeight = _GDIPlus_ImageGetHeight($hImage)
			_GDIPlus_ImageDispose($hImage)

			; Scale down if too large (max 300px for tooltip)
			If $iImageWidth > 300 Or $iImageHeight > 200 Then
				Local $scale = 300 / _Max($iImageWidth, $iImageHeight)
				$iImageWidth *= $scale
				$iImageHeight *= $scale
			EndIf
		EndIf
		_GDIPlus_Shutdown()
	EndIf

	; Create tooltip window with dynamic size based on image
	Local $iW = $iImageWidth + 20 ; Add padding
	Local $iH = $iImageHeight + 20 ; Add padding
	Local $mousePos = MouseGetPos()
	Local $iX = $mousePos[0] + 10
	Local $iY = $mousePos[1] + 10

	; Adjust position if tooltip would go off screen
	If $iX + $iW > @DesktopWidth Then $iX = @DesktopWidth - $iW - 10
	If $iY + $iH > @DesktopHeight Then $iY = @DesktopHeight - $iH - 10

	$g_TooltipGUI = GUICreate("", $iW, $iH, $iX, $iY, $WS_POPUP, $WS_EX_TOPMOST)
	GUISetBkColor(0xFFFFE1, $g_TooltipGUI) ; Light yellow background

	; Add image if it exists
	If FileExists($imagePath) Then
		GUICtrlCreatePic($imagePath, 10, 10, $iImageWidth, $iImageHeight)
	Else
		GUICtrlCreateLabel("Image not found:", 10, 10, $iImageWidth, 20)
		GUICtrlCreateLabel($imagePath, 10, 35, $iImageWidth, $iImageHeight - 25)
	EndIf

	GUISetState(@SW_SHOW, $g_TooltipGUI)

	; Auto-close after 5 seconds or when clicking anywhere
	AdlibRegister("_CloseImageTooltip", 5000)
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
	Debug("Searching for element with class: '" & $sClassName & "'", "VERBOSE")

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

	Debug("Element with class '" & $sClassName & "' found.", "VERBOSE")
	Return $oElement
EndFunc   ;==>FindElementByClassName

; Gets reference to the main Zoom meeting window
; @return Object - Zoom window element or error
Func _GetZoomWindow()
	$oZoomWindow = FindElementByClassName("ConfMultiTabContentWndClass", $TreeScope_Children)
	If Not IsObj($oZoomWindow) Then
		Debug(t("ERROR_ZOOM_WINDOW_NOT_FOUND"), "ERROR")
		Return SetError(1, 0, 0)
	EndIf
	Debug("Zoom window obtained.", "UIA")
	Return $oZoomWindow
EndFunc   ;==>_GetZoomWindow

; Internal function to find Zoom window (used by cache)
Func _FindZoomWindowInternal()
	Return FindElementByClassName("ConfMultiTabContentWndClass", $TreeScope_Children)
EndFunc   ;==>_FindZoomWindowInternal

; Focuses the main Zoom meeting window
; @return Boolean - True if successful, False otherwise
Func FocusZoomWindow()
	Debug("Focusing Zoom window...", "VERBOSE")
	Local $oZoomWindow = _GetZoomWindow()
	If Not IsObj($oZoomWindow) Then Return False

	; Get the native HWND property from the UIA element
	Local $hWnd
	$oZoomWindow.GetCurrentPropertyValue($UIA_NativeWindowHandlePropertyId, $hWnd)

	If $hWnd And $hWnd <> 0 Then
		; Convert to HWND pointer
		$hWnd = Ptr($hWnd)
		WinActivate($hWnd)
		If WinWaitActive($hWnd, "", 3) Then
			Debug("Zoom window activated and focused.", "VERBOSE")
			Return True
		Else
			Debug(t("ERROR_ZOOM_WINDOW_NOT_FOUND"), "ERROR")
		EndIf
	Else
		Debug(t("ERROR_ZOOM_WINDOW_NOT_FOUND"), "ERROR")
	EndIf

	Return False
EndFunc   ;==>FocusZoomWindow


; Snaps the Zoom window to a side of the primary monitor using Windows snap shortcuts
; @return Boolean - True if moved, False otherwise
Func _SnapZoomWindowToSide()
	Local $snapSide = GetUserSetting("SnapZoomSide")
	If StringLower($snapSide) = "disabled" Then
		Debug("Zoom window snapping disabled; skipping.", "VERBOSE")
		Return False
	EndIf

	_GetZoomWindow()
	If Not IsObj($oZoomWindow) Then Return False

	; Activate the Zoom window first using the proper focus function
	If Not FocusZoomWindow() Then Return False

	; Check if window is already snapped to the desired side (with tolerance for taskbar/borders)
	Local $aBoundingRect
	$oZoomWindow.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $aBoundingRect)

	Debug("UIA Bounding Rectangle: " & $aBoundingRect & " (raw)", "VERBOSE")

	If IsArray($aBoundingRect) And UBound($aBoundingRect) >= 4 Then
		; UIA BoundingRectangle format: [left, top, width, height]
		Local $iLeft = $aBoundingRect[0]
		Local $iTop = $aBoundingRect[1]
		Local $iWidth = $aBoundingRect[2]
		Local $iHeight = $aBoundingRect[3]

		Debug("Parsed position - X:" & $iLeft & " Y:" & $iTop & " W:" & $iWidth & " H:" & $iHeight, "VERBOSE")

		Local $iScreenWidth = @DesktopWidth
		Local $iScreenHeight = @DesktopHeight

		; Calculate expected positions for snapped windows (with 50px tolerance for borders/taskbar)
		Local $iTolerance = $SNAP_TOLERANCE_PX
		Local $iHalfWidth = $iScreenWidth / 2

		Debug("Window position check - X:" & $iLeft & " Y:" & $iTop & " W:" & $iWidth & " H:" & $iHeight & " | Screen:" & $iScreenWidth & "x" & $iScreenHeight & " | Half:" & $iHalfWidth & " | Tolerance:" & $iTolerance, "VERBOSE")

		Local $bIsLeftSnapped = ($iLeft <= $iTolerance And _
				$iTop >= -$iTolerance And $iTop <= $iTolerance And _
				Abs($iWidth - $iHalfWidth) <= $iTolerance And _
				Abs($iHeight - $iScreenHeight) <= $iTolerance)

		Local $bIsRightSnapped = ($iLeft >= $iHalfWidth - $iTolerance And _
				$iLeft <= $iHalfWidth + $iTolerance And _
				$iTop >= -$iTolerance And $iTop <= $iTolerance And _
				Abs($iWidth - $iHalfWidth) <= $iTolerance And _
				Abs($iHeight - $iScreenHeight) <= $iTolerance)

		Debug("Position analysis - LeftSnapped:" & $bIsLeftSnapped & " RightSnapped:" & $bIsRightSnapped & " | TargetSide:" & $snapSide, "VERBOSE")

		If StringLower($snapSide) = "left" And $bIsLeftSnapped Then
			Debug("Zoom window already snapped to left side; skipping.", "VERBOSE")
			Return True
		ElseIf StringLower($snapSide) = "right" And $bIsRightSnapped Then
			Debug("Zoom window already snapped to right side; skipping.", "VERBOSE")
			Return True
		EndIf
	Else
		Debug("Failed to get UIA BoundingRectangle - IsArray:" & IsArray($aBoundingRect) & " UBound:" & (IsArray($aBoundingRect) ? UBound($aBoundingRect) : "N/A"), "WARN")
	EndIf

	; Use Windows snap shortcuts instead of manual positioning
	If StringLower($snapSide) = "left" Then
		Send("#{LEFT}") ; Windows + Left arrow
		Debug("Sent Windows+Left to snap Zoom window to left half", "VERBOSE")
	ElseIf StringLower($snapSide) = "right" Then
		Send("#{RIGHT}") ; Windows + Right arrow
		Debug("Sent Windows+Right to snap Zoom window to right half", "VERBOSE")
	Else
		Debug(t("ERROR_INVALID_SNAP_SELECTION", $snapSide), "ERROR")
		Return False
	EndIf

	; Wait 0.5 seconds for snap animation to complete
	Sleep($WINDOW_SNAP_DELAY_MS)

	; Send Escape to dismiss any remaining Windows snap UI
	Send("{ESC}")
	Debug("Sent Escape to dismiss Windows snap UI", "VERBOSE")

	Debug("Zoom window snapped to " & $snapSide & " half of primary monitor using Windows shortcuts.", "VERBOSE")
	Return True
EndFunc   ;==>_SnapZoomWindowToSide


; Finds UI element by partial name match across multiple control types
; @param $sPartial - Partial text to search for in element names
; @param $aControlTypes - Array of control types to search (default: button and menu item)
; @param $oParent - Parent element to search within (default: desktop)
; @return Object - Found element or 0 if not found
Func FindElementByPartialName($sPartial, $aControlTypes = Default, $oParent = Default)
	Debug("Searching for element containing: '" & $sPartial & "'", "VERBOSE")

	; Use desktop as default parent if not specified
	Local $oSearchParent = $oDesktop
	If $oParent <> Default Then
		$oSearchParent = $oParent
		Debug("Using custom parent element for search", "VERBOSE")
	Else
		Debug("Using desktop as search parent", "VERBOSE")
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
			Debug("Found " & $iCount & " elements of this type.", "VERBOSE")

			; Check each element for partial name match
			For $i = 0 To $iCount - 1
				Local $pElement
				$oElements.GetElement($i, $pElement)

				Local $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
				If IsObj($oElement) Then
					; Get element name and check for partial match
					Local $sName
					$oElement.GetCurrentPropertyValue($UIA_NamePropertyId, $sName)
					Debug("Element found with name: '" & $sName & "'", "VERBOSE")

					If StringInStr($sName, $sPartial, $STR_NOCASESENSEBASIC) > 0 Then
						Debug("Matching element found with name: '" & $sName & "'", "VERBOSE")
						Return $oElement
					EndIf
				EndIf
			Next
		EndIf
	Next

	Debug("No element found containing: '" & $sPartial & "'", "WARN")
	Return 0
EndFunc   ;==>FindElementByPartialName

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

; Clicks a UI element using multiple methods for maximum compatibility with timeout protection
; @param $oElement - Element to click
; @param $ForceClick - If True, forces mouse click method
; @param $BoundingRectangle - If True, forces click by bounding rectangle
; @param $iTimeoutMs - Timeout in milliseconds (default: 5000)
; @return Boolean - True if successful, False otherwise
Func _ClickElement($oElement, $ForceClick = False, $BoundingRectangle = False, $iTimeoutMs = $CLICK_TIMEOUT_MS)
	ResponsiveSleep(0.5) ; Brief pause to ensure UI is ready
	Local $iStartTime = TimerInit()

	If Not IsObj($oElement) Then
		Debug(t("ERROR_INVALID_ELEMENT_OBJECT"), "WARN")
		Return False
	EndIf

	; Get element name for debugging
	Local $sElementName = GetElementName($oElement)
	Debug("Attempting to click element: '" & $sElementName & "'", "VERBOSE")

	; Method 0: Force mouse click (only when requested)
	If $ForceClick Then
		; Method 0.5: Force mouse click by bounding rectangle (only when requested)
		If $BoundingRectangle Then
			If TimerDiff($iStartTime) > $iTimeoutMs Then
				Debug("Click operation timed out after " & $iTimeoutMs & "ms", "VERBOSE")
				Return False
			EndIf
			If not ClickByBoundingRectangle($oElement) then
				Return False
			EndIf
			Debug("Element clicked via bounding rectangle.", "VERBOSE")
		Else
			If TimerDiff($iStartTime) > $iTimeoutMs Then
				Debug("Click operation timed out after " & $iTimeoutMs & "ms", "VERBOSE")
				Return False
			EndIf
			UIA_MouseClick($oElement)
			Debug("Element clicked via UIA_MouseClick.", "VERBOSE")
		EndIf
		Return True
	EndIf

	; Method 1: Try Invoke Pattern (works for most buttons)
	Local $pInvokePattern, $oInvokePattern
	$oElement.GetCurrentPattern($UIA_InvokePatternId, $pInvokePattern)
	If $pInvokePattern Then
		$oInvokePattern = ObjCreateInterface($pInvokePattern, $sIID_IUIAutomationInvokePattern, $dtagIUIAutomationInvokePattern)
		If IsObj($oInvokePattern) Then
			If TimerDiff($iStartTime) > $iTimeoutMs Then
				Debug("Click operation timed out after " & $iTimeoutMs & "ms", "VERBOSE")
				Return False
			EndIf
			Debug("Using Invoke pattern for: '" & $sElementName & "'", "VERBOSE")
			$oInvokePattern.Invoke()
			Debug("Element clicked via Invoke pattern.", "VERBOSE")
			Return True
		EndIf
	EndIf

	; Method 2: Try Legacy Accessible Pattern (works for menu items and older controls)
	Local $pLegacyPattern, $oLegacyPattern
	$oElement.GetCurrentPattern($UIA_LegacyIAccessiblePatternId, $pLegacyPattern)
	If $pLegacyPattern Then
		$oLegacyPattern = ObjCreateInterface($pLegacyPattern, $sIID_IUIAutomationLegacyIAccessiblePattern, $dtagIUIAutomationLegacyIAccessiblePattern)
		If IsObj($oLegacyPattern) Then
			If TimerDiff($iStartTime) > $iTimeoutMs Then
				Debug("Click operation timed out after " & $iTimeoutMs & "ms", "VERBOSE")
				Return False
			EndIf
			Debug("Using LegacyAccessible pattern for: '" & $sElementName & "'", "VERBOSE")
			$oLegacyPattern.DoDefaultAction()
			Debug("Element clicked via LegacyAccessible pattern.", "VERBOSE")
			Return True
		EndIf
	EndIf

	; Method 3: Try Selection Item Pattern (for selectable menu items)
	Local $pSelectionItemPattern, $oSelectionItemPattern
	$oElement.GetCurrentPattern($UIA_SelectionItemPatternId, $pSelectionItemPattern)
	If $pSelectionItemPattern Then
		$oSelectionItemPattern = ObjCreateInterface($pSelectionItemPattern, $sIID_IUIAutomationSelectionItemPattern, $dtagIUIAutomationSelectionItemPattern)
		If IsObj($oSelectionItemPattern) Then
			If TimerDiff($iStartTime) > $iTimeoutMs Then
				Debug("Click operation timed out after " & $iTimeoutMs & "ms", "VERBOSE")
				Return False
			EndIf
			Debug("Using SelectionItem pattern for: '" & $sElementName & "'", "VERBOSE")
			$oSelectionItemPattern.Select()
			Debug("Element selected via SelectionItem pattern.", "VERBOSE")
			Return True
		EndIf
	EndIf

	; Method 4: Try Toggle Pattern (for toggle buttons/menu items)
	Local $pTogglePattern, $oTogglePattern
	$oElement.GetCurrentPattern($UIA_TogglePatternId, $pTogglePattern)
	If $pTogglePattern Then
		$oTogglePattern = ObjCreateInterface($pTogglePattern, $sIID_IUIAutomationTogglePattern, $dtagIUIAutomationTogglePattern)
		If IsObj($oTogglePattern) Then
			If TimerDiff($iStartTime) > $iTimeoutMs Then
				Debug("Click operation timed out after " & $iTimeoutMs & "ms", "VERBOSE")
				Return False
			EndIf
			Debug("Using Toggle pattern for: '" & $sElementName & "'", "VERBOSE")
			$oTogglePattern.Toggle()
			Debug("Element toggled via Toggle pattern.", "VERBOSE")
			Return True
		EndIf
	EndIf

	; Method 5: Fallback - Mouse click at element center
	If TimerDiff($iStartTime) > $iTimeoutMs Then
		Debug("Click operation timed out after " & $iTimeoutMs & "ms", "VERBOSE")
		Return False
	EndIf
	if not ClickByBoundingRectangle($oElement) then
		Return False
	EndIf

	; All click methods failed
	Debug(t("ERROR_FAILED_CLICK_ELEMENT") & ": '" & $sElementName & "'", "VERBOSE")
	Return False
EndFunc   ;==>_ClickElement

Func ClickByBoundingRectangle($oElement)
	; Method 0.1: Click by bounding rectangle center
	Local $sElementName = GetElementName($oElement)
	Local $tRect
	$oElement.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $tRect)
	UIA_GetArrayPropertyValueAsString($tRect)
	Debug("Element bounding rectangle: " & $tRect, "VERBOSE")
	If Not $tRect Then
		Debug("No bounding rectangle for element: '" & $sElementName & "'", "VERBOSE")
		Return False
	EndIf

	Local $aRect = StringSplit($tRect, ",")
	If $aRect[0] < 4 Then
		Debug("Invalid rectangle format for: '" & $sElementName & "'", "VERBOSE")
		Return False
	EndIf

	Local $iLeft = Number($aRect[1])
	Local $iTop = Number($aRect[2])
	Local $iWidth = Number($aRect[3])
	Local $iHeight = Number($aRect[4])

	Local $iCenterX = $iLeft + ($iWidth / 2)
	Local $iCenterY = $iTop + ($iHeight / 2)

	Debug("Using mouse click fallback at position: " & $iCenterX & "," & $iCenterY & " for: '" & $sElementName & "'", "VERBOSE")

	; Ensure element is clickable before attempting
	Local $bIsEnabled, $bIsOffscreen
	$oElement.GetCurrentPropertyValue($UIA_IsEnabledPropertyId, $bIsEnabled)
	$oElement.GetCurrentPropertyValue($UIA_IsOffscreenPropertyId, $bIsOffscreen)

	If $bIsEnabled And Not $bIsOffscreen Then
		MouseClick("primary", $iCenterX, $iCenterY, 1, 0)
		Debug("Element clicked via mouse at center.", "VERBOSE")
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
Func _HoverElement($oElement, $iHoverTime = $HOVER_DEFAULT_MS, $SlightOffset = False)
	ResponsiveSleep(0.3) ; Small buffer before hover
	If Not IsObj($oElement) Then
		Debug("Invalid element passed to _HoverElement", "WARN")
		Return False
	EndIf

	; Get element name for debugging
	Local $sElementName = GetElementName($oElement)
	Debug("Attempting to hover element: '" & $sElementName & "'", "VERBOSE")

	; Get bounding rectangle
	Local $tRect
	$oElement.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $tRect)
	UIA_GetArrayPropertyValueAsString($tRect)
	Debug("Element bounding rectangle: " & $tRect, "VERBOSE")
	If Not $tRect Then
		Debug("No bounding rectangle for element: '" & $sElementName & "'", "WARN")
		Return False
	EndIf

	Local $aRect = StringSplit($tRect, ",")
	If $aRect[0] < 4 Then
		Debug("Invalid rectangle format for: '" & $sElementName & "'", "WARN")
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
		Debug("Applying slight offset to hover position: " & $iOffsetX & "," & $iOffsetY, "VERBOSE")
	EndIf

	Debug("Hovering at: " & $iCenterX & "," & $iCenterY & " for " & $iHoverTime & "ms", "VERBOSE")

	; Move mouse to center of element and hold position
	MouseMove($iCenterX, $iCenterY, 0)
	Sleep($iHoverTime)

	Debug("Hover completed on element: '" & $sElementName & "'", "VERBOSE")
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
	Debug("Attempting to move mouse to start of element: '" & $sElementName & "'", "VERBOSE")

	; Get bounding rectangle
	Local $tRect
	$oElement.GetCurrentPropertyValue($UIA_BoundingRectanglePropertyId, $tRect)
	UIA_GetArrayPropertyValueAsString($tRect)
	Debug("Element bounding rectangle: " & $tRect, "VERBOSE")
	If Not $tRect Then
		Debug("No bounding rectangle for element: '" & $sElementName & "'", "WARN")
		Return False
	EndIf

	Local $aRect = StringSplit($tRect, ",")
	If $aRect[0] < 4 Then
		Debug("Invalid rectangle format for: '" & $sElementName & "'", "WARN")
		Return False
	EndIf

	Local $iLeft = Number($aRect[1])
	Local $iTop = Number($aRect[2])
;~ Local $iWidth = Number($aRect[3])
	Local $iHeight = Number($aRect[4])

	; Move mouse to start (left edge, vertically centered)
	Local $iStartX = $iLeft + Random(5, 30, 1) ; Random offset from left edge
	Local $iStartY = $iTop + ($iHeight / 2) + Random(-5, 5, 1) ; Random offset from center

	Debug("Moving mouse to start position: " & $iStartX & "," & $iStartY, "VERBOSE")

	; Move mouse to start of element
	MouseMove($iStartX, $iStartY, 0)
	Sleep(200) ; Brief pause after move
	Debug("Mouse moved to start of element: '" & $sElementName & "'", "VERBOSE")

	If $Click Then
		MouseClick("primary", $iStartX, $iStartY, 1, 0)
		Debug("Element clicked at start position.", "VERBOSE")
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
		Debug("Controls might be hidden. Moving mouse to show controls.", "VERBOSE")
		Debug("Getting Zoom window.", "VERBOSE")
		if not _GetZoomWindow() then return false

		Debug("Moving mouse to start of element: 'Zoom window'", "VERBOSE")
		_MoveMouseToStartOfElement($oZoomWindow, True)
		Debug("Clicking Host Tools button.", "VERBOSE")

		; Menu not open, find and click the Host Tools button
		Local $oHostToolsButton = FindElementByPartialName(GetUserSetting("HostToolsValue"), Default, $oZoomWindow)

		; Scenario 1: Try to find Host Tools button directly
		If IsObj($oHostToolsButton) Then
			Debug("Host Tools button found directly; clicking.", "VERBOSE")
			If Not _ClickElement($oHostToolsButton) Then
				Debug(t("ERROR_FAILED_CLICK_ELEMENT") & ": 'Host Tools button'", "ERROR")
				Return False
			else 
				ResponsiveSleep(0.5)
				$oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
				Return $oHostMenu
			EndIf
		Else
			; Scenario 2: Try to find More menu, then Host Tools
			Debug("Host Tools button not found, looking for 'More' button.", "VERBOSE")
			Local $oMoreMenu = GetMoreMenu()
			If IsObj($oMoreMenu) Then
				; Now look for the Host Tools button in the More menu
				Local $oHostToolsMenuItem = FindElementByPartialName(GetUserSetting("HostToolsValue"), Default, $oMoreMenu)
				If IsObj($oHostToolsMenuItem) Then
					Debug("Found Host Tools menu item. Hovering it to open submenu.", "VERBOSE")
					If _HoverElement($oHostToolsMenuItem, 500) Then
						; Return the now-open host menu
						$oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
						Return $oHostMenu
					Else
						Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("HostToolsValue")), "ERROR")
						Return False
					EndIf
				Else
					Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("HostToolsValue")), "ERROR")
					Return False
				EndIf
			Else
				Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("MoreMeetingControlsValue")), "ERROR")
				Return False
			EndIf
		EndIf
	EndIf

	; Return the now-open host menu
	Return $oHostMenu
EndFunc   ;==>_OpenHostTools

; Internal function to find Host menu (used by cache)
Func _FindHostMenuInternal()
	Return FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
EndFunc   ;==>_FindHostMenuInternal

; Closes the Host Tools menu by clicking on the main window
Func _CloseHostTools()
	Debug(t("INFO_CLOSE_HOST_TOOLS"), "INFO")
	_MoveMouseToStartOfElement($oZoomWindow, True) ; Click at start of window to ensure menu closes
EndFunc   ;==>_CloseHostTools

; Opens the "More" menu in Zoom if available
; @return Object - More menu object or False if failed
Func GetMoreMenu()
	Debug(t("INFO_GET_MORE_MENU"), "INFO")
	If Not IsObj($oZoomWindow) Then Return False

	Local $oMoreMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)

	If Not IsObj($oMoreMenu) Then
		; Menu not open, find and click the More button
		Debug("More menu not open, attempting to open.", "VERBOSE")
		Debug("Controls might be hidden. Moving mouse to show controls.", "VERBOSE")
		_MoveMouseToStartOfElement($oZoomWindow)

		Local $oMoreButton = FindElementByPartialName(GetUserSetting("MoreMeetingControlsValue"), Default, $oZoomWindow)
		If Not IsObj($oMoreButton) Then
			Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("MoreMeetingControlsValue")), "ERROR")
			Return False
		EndIf
		Debug("Clicking More button to open menu.", "VERBOSE")
		If Not _ClickElement($oMoreButton, True) Then
			Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("MoreMeetingControlsValue")), "ERROR")
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
		Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("MoreMeetingControlsValue")), "ERROR")
	EndIf
	Return $oMoreMenu
EndFunc   ;==>GetMoreMenu

; Internal function to find More menu (used by cache)
Func _FindMoreMenuInternal()
	Return FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
EndFunc   ;==>_FindMoreMenuInternal

; Opens the Participants panel in Zoom
; @return Object - Participants panel object or False if failed
Func _OpenParticipantsPanel()
	Debug(t("INFO_OPEN_PARTICIPANTS_PANEL"), "INFO")	
	If Not IsObj($oZoomWindow) Then Return False

	; Controls might be hidden, show them by moving the mouse
	Debug("Controls might be hidden. Moving mouse to show controls.", "VERBOSE")
	_MoveMouseToStartOfElement($oZoomWindow)

	Local $ListType[1] = [$UIA_ListControlTypeId]
	Local $oParticipantsPanel = FindElementByPartialName(GetUserSetting("ParticipantValue"), $ListType, $oZoomWindow)

	If Not IsObj($oParticipantsPanel) Then
		; Panel not open, find and click the Participants button
		Debug("Participants panel not open, attempting to open.", "UIA")
		Local $oMainParticipantsButton = FindElementByPartialName(GetUserSetting("ParticipantValue"), Default, $oZoomWindow)

		; Scenario 1: Try to find Participants button directly
		If IsObj($oMainParticipantsButton) Then
			Debug("Participants button found directly; clicking.", "VERBOSE")
			If Not _ClickElement($oMainParticipantsButton) Then
				Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("ParticipantValue")), "ERROR")
				Return False
			EndIf
		Else
			; Scenario 2: Try to find More menu, then Participants
			Debug("Participants button not found, looking for 'More' button.", "VERBOSE")
			Local $oMoreMenu = GetMoreMenu()
			If IsObj($oMoreMenu) Then
				; Now look for the Participants button in the More menu
				Local $oParticipantsMenuItem = FindElementByPartialName(GetUserSetting("ParticipantValue"), Default, $oMoreMenu)
				If IsObj($oParticipantsMenuItem) Then
					Debug("Found Participants menu item. Hovering it to open submenu.", "VERBOSE")
					If _HoverElement($oParticipantsMenuItem, 1200) Then         ; 1.2s hover to ensure submenu appears
						; Now look for the Participants button again in the submenu
						Debug("Looking for Participants button again in submenu.", "VERBOSE")
						Local $oParticipantsSubMenuItem = FindElementByPartialName(GetUserSetting("ParticipantValue"), Default, $oZoomWindow)
						If IsObj($oParticipantsSubMenuItem) Then
							Debug("Final Participants button found. Clicking it.", "VERBOSE")
							_HoverElement($oParticipantsSubMenuItem, 500)
							_MoveMouseToStartOfElement($oParticipantsSubMenuItem, True)
							Debug("Participants button clicked.", "VERBOSE")
							ResponsiveSleep(0.5) ; Move mouse to start of element and click to avoid hover issues
						Else
							Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("ParticipantValue")), "ERROR")
							Return False
						EndIf
					Else
						Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("ParticipantValue")), "ERROR")
						Return False
					EndIf
				Else
					Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("ParticipantValue")), "ERROR")
					Return False
				EndIf
			Else
				Debug(t("ERROR_FAILED_CLICK_ELEMENT", GetUserSetting("ParticipantValue")), "ERROR")
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
		Debug(t("ERROR_FAILED_OPEN_PANEL", GetUserSetting("ParticipantValue")), "ERROR")
	EndIf
	Return $oParticipantsPanel
EndFunc   ;==>_OpenParticipantsPanel

; Internal function to find Participants panel (used by cache)
Func _FindParticipantsPanelInternal()
	Local $ListType[1] = [$UIA_ListControlTypeId]
	Return FindElementByPartialName(GetUserSetting("ParticipantValue"), $ListType, $oZoomWindow)
EndFunc   ;==>_FindParticipantsPanelInternal

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
		Debug("Previous keyboard shortcut unregistered: " & $g_KeyboardShortcut, "VERBOSE")
	EndIf

	; Update the global keyboard shortcut variable
	$g_KeyboardShortcut = GetUserSetting("KeyboardShortcut")

	; Register new hotkey if not empty
	If $g_KeyboardShortcut <> "" Then
		; Validate the shortcut format (basic validation)
		If StringRegExp($g_KeyboardShortcut, "^[\^\!\+\#]+[a-zA-Z0-9]$") Then
			HotKeySet($g_KeyboardShortcut, "_ManualTrigger")
			$g_HotkeyRegistered = True
			Debug("New keyboard shortcut registered: " & $g_KeyboardShortcut, "VERBOSE")
		Else
			Debug("Invalid keyboard shortcut format: " & $g_KeyboardShortcut, "VERBOSE")
			$g_KeyboardShortcut = ""
			IniWrite($CONFIG_FILE, "General", "KeyboardShortcut", "")
		EndIf
	Else
		Debug("Keyboard shortcut cleared", "VERBOSE")
	EndIf
EndFunc   ;==>_UpdateKeyboardShortcut

; Manual trigger function activated by keyboard shortcut
; Allows user to manually apply post-meeting settings
Func _ManualTrigger()    ; Show message and wait for user input before applying settings
	Debug("Manual trigger: Showing post-meeting message", "VERBOSE")

	; Show the post-meeting message and wait for Enter key
	ShowOverlayMessage	('POST_MEETING_HIT_KEY')

	; Wait for user to press Enter or ESC
	Local $userInput = _WaitForEnterOrEscape()

	; Hide the message
	HideOverlayMessage()

	; Apply settings if user pressed Enter
	If $userInput = "ENTER" Then
		Debug("User pressed Enter: Applying post-meeting settings", "VERBOSE")
		_SetPreAndPostMeetingSettings()
	Else
		Debug("User pressed Escape or closed dialog: Cancelling", "VERBOSE")
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
;~ Func GetSecuritySettingsState($aSettings)
;~ 	Local $oHostMenu = _OpenHostTools()
;~ 	If Not IsObj($oHostMenu) Then Return False

;~ 	; Create dictionary to store setting states
;~ 	Local $oDict = ObjCreate("Scripting.Dictionary")

;~ 	For $i = 0 To UBound($aSettings) - 1
;~ 		Local $sSetting = $aSettings[$i]

;~ 		; Find the setting menu item (cached)
;~ 		Local $oSetting = GetCachedElement("SecuritySetting_" & $sSetting, "_FindSecuritySettingInternal", 500, $sSetting, $oHostMenu)
;~ 		Local $bEnabled = False
;~ 		If IsObj($oSetting) Then
;~ 			; Read the setting's label to determine if it's checked
;~ 			Local $sLabel = ""
;~ 			$oSetting.GetCurrentPropertyValue($UIA_NamePropertyId, $sLabel)

;~ 			Debug("Element name: '" & $sLabel & "'", "VERBOSE")

;~ 			; Setting is enabled if label does NOT contain unchecked indicator
;~ 			Local $uncheckedValue = GetUserSetting("UncheckedValue")
;~ 			Local $sLabelLower = StringLower($sLabel)
;~ 			Local $uncheckedLower = StringLower($uncheckedValue)

;~ 			Debug("Raw label: '" & $sLabel & "'", "VERBOSE")
;~ 			Debug("Label length: " & StringLen($sLabel), "VERBOSE")
;~ 			Debug("Unchecked value: '" & $uncheckedValue & "'", "VERBOSE")
;~ 			Debug("Unchecked length: " & StringLen($uncheckedValue), "VERBOSE")
;~ 			Debug("Label lower: '" & $sLabelLower & "'", "VERBOSE")
;~ 			Debug("Unchecked lower: '" & $uncheckedLower & "'", "VERBOSE")

;~ 			; Check if unchecked value appears anywhere in the label (not just word boundaries)
;~ 			$bEnabled = (StringInStr($sLabelLower, $uncheckedLower) = 0)

;~ 			Debug("StringInStr result: " & StringInStr($sLabelLower, $uncheckedLower), "VERBOSE")
;~ 		Else
;~ 			Debug(t("ERROR_SETTING_NOT_FOUND") & ": '" & $sSetting & "'", "ERROR")
;~ 		EndIf
;~ 		$oDict.Add($sSetting, $bEnabled)
;~ 	Next

;~ 	_CloseHostTools()
;~ 	Return $oDict
;~ EndFunc   ;==>GetSecuritySettingsState

; Internal function to find security setting (used by cache)
Func _FindSecuritySettingInternal($sSetting, $oHostMenu)
	Return FindElementByPartialName($sSetting, Default, $oHostMenu)
EndFunc   ;==>_FindSecuritySettingInternal

; Sets a security setting to the desired state (enabled/disabled)
; @param $sSetting - Setting name to modify
; @param $bDesired - Desired state (True=enabled, False=disabled)
Func SetSecuritySetting($sSetting, $bDesired)
	Debug(t("INFO_SETTING_SECURITY", $sSetting), "INFO")
	ResponsiveSleep(3)
	Local $oHostMenu = _OpenHostTools()
	If Not IsObj($oHostMenu) Then Return False

	Local $oSetting = FindElementByPartialName($sSetting, Default, $oHostMenu)
	If Not IsObj($oSetting) Then Return

	; Check current state
	Local $sLabel
	$oSetting.GetCurrentPropertyValue($UIA_NamePropertyId, $sLabel)

	Debug("Element name: '" & $sLabel & "'", "VERBOSE")

	; Setting is enabled if label does NOT contain unchecked indicator
	Local $uncheckedValue = GetUserSetting("UncheckedValue")
	Local $sLabelLower = StringLower($sLabel)
	Local $uncheckedLower = StringLower($uncheckedValue)

	Debug("Raw label: '" & $sLabel & "'", "VERBOSE")
	Debug("Label length: " & StringLen($sLabel), "VERBOSE")
	Debug("Unchecked value: '" & $uncheckedValue & "'", "VERBOSE")
	Debug("Unchecked length: " & StringLen($uncheckedValue), "VERBOSE")
	Debug("Label lower: '" & $sLabelLower & "'", "VERBOSE")
	Debug("Unchecked lower: '" & $uncheckedLower & "'", "VERBOSE")

	; Check if unchecked value appears anywhere in the label (not just word boundaries)
	Local $bEnabled = (StringInStr($sLabelLower, $uncheckedLower) = 0)

	Debug("StringInStr result: " & StringInStr($sLabelLower, $uncheckedLower), "VERBOSE")
	Debug("Setting '" & $sLabel & "' | Current: " & ($bEnabled ? "True" : "False") & " | Desired: " & $bDesired, "VERBOSE")

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
	Debug(t("INFO_TOGGLE_FEED", $feedType), "INFO")
	Local $currentlyEnabled = False

	; Controls might be hidden, show them by moving the mouse
	Debug("Controls might be hidden. Moving mouse to show controls.", "VERBOSE")
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
	Debug(t("INFO_MUTE_ALL"), "INFO")

	; Open participants panel
	Local $oParticipantsPanel = _OpenParticipantsPanel()
	If Not IsObj($oParticipantsPanel) Then Return False

	; Find and click "Mute All" button
	Local $oButton = FindElementByPartialName(GetUserSetting("MuteAllValue"), Default, $oZoomWindow)
	If Not _ClickElement($oButton) Then
		Debug(t("ERROR_ELEMENT_NOT_FOUND", GetUserSetting("MuteAllValue")), "ERROR")
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
	Debug(t("ERROR_ELEMENT_NOT_FOUND", $ButtonLabel), "ERROR")
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

	if not _GetZoomWindow() then return false

	_SnapZoomWindowToSide()

	Return IsObj($oZoomWindow)
EndFunc   ;==>_LaunchZoom

; Configures settings before and after meetings
; - Enables unmute permission for participants
; - Disables screen sharing permission
; - Turns off host audio and video
Func _SetPreAndPostMeetingSettings()
	Debug(t("INFO_CONFIG_BEFORE_AFTER_START"), "INFO")
	ResponsiveSleep(3)
	If Not FocusZoomWindow() Then
		Return
	EndIf
	SetSecuritySetting(GetUserSetting("ZoomSecurityUnmuteValue"), True)          ; Allow participants to unmute
	SetSecuritySetting(GetUserSetting("ZoomSecurityShareScreenValue"), False)   ; Prevent screen sharing
	ToggleFeed("Audio", False)                  ; Turn off host audio
	ToggleFeed("Video", False)                  ; Turn off host video
	; TODO: Unmute All function
	Debug(t("INFO_CONFIG_BEFORE_AFTER_DONE"), "INFO")
EndFunc   ;==>_SetPreAndPostMeetingSettings

; Configures settings during active meetings
; - Disables unmute permission (host controls audio)
; - Disables screen sharing permission
; - Mutes all participants
; - Turns on host audio and video
Func _SetDuringMeetingSettings()
	Debug(t("INFO_MEETING_STARTING_SOON_CONFIG"), "INFO")
	ResponsiveSleep(3)
	If Not FocusZoomWindow() Then
		Return
	EndIf
	SetSecuritySetting(GetUserSetting("ZoomSecurityUnmuteValue"), False)         ; Prevent participant self-unmute
	SetSecuritySetting(GetUserSetting("ZoomSecurityShareScreenValue"), False)   ; Prevent screen sharing
	MuteAll()                                   ; Mute all participants
	ToggleFeed("Audio", True)                   ; Turn on host audio
	ToggleFeed("Video", True)                   ; Turn on host video
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

	If $nowMin >= ($meetingMin - $PRE_MEETING_MINUTES) And $nowMin < ($meetingMin - $MEETING_START_WARNING_MINUTES) Then
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

	ElseIf $nowMin = ($meetingMin - $MEETING_START_WARNING_MINUTES) Then
		; Meeting start window (1 minute before meeting)
		If Not $g_DuringMeetingSettingsConfigured Then
			if not _GetZoomWindow() then return
			_SetDuringMeetingSettings()
			$g_DuringMeetingSettingsConfigured = True
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
		Debug("New day detected. Resetting configuration flags.", "VERBOSE")
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
		Debug("GetElementNamesFromWindows: Invalid Zoom window object", "VERBOSE")
		Return $aNames
	EndIf

	; Define control types to search (same as FindElementByPartialName default)
	Local $aControlTypes[2] = [$UIA_ButtonControlTypeId, $UIA_MenuItemControlTypeId]

	; Collect names from Zoom window
	_CollectElementNames($oZoomWindow, $aControlTypes, $aNames)

	Debug("Collected these values: " & $aNames & " from Zoom window", "VERBOSE")

	; Collect names from Host menu if provided
	If IsObj($oHostMenu) Then
		_CollectElementNames($oHostMenu, $aControlTypes, $aNames)
	EndIf

	; Remove duplicates and sort
	$aNames = _ArrayUnique($aNames)
	_ArraySort($aNames)

	Debug("Collected " & UBound($aNames) & " unique element names", "VERBOSE")
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

; Shows a selectable list of collected element names for user to choose from
; @param $aNames - Array of element names to display
; @param $callbackFunc - Function to call when user makes a selection
Func ShowElementNamesSelectionGUI($aNames, $callbackFunc)
	; Close any existing selection GUI
	CloseElementNamesSelectionGUI()

	; Store callback function for when user makes selection
	$g_ElementNamesSelectionCallback = $callbackFunc

	Local $iW = 500
	Local $iH = 400
	Local $iX = (@DesktopWidth - $iW) / 2
	Local $iY = (@DesktopHeight - $iH) / 2

	; Create moveable and closeable GUI
	$g_ElementNamesSelectionGUI = GUICreate("Select Element Name", $iW, $iH, $iX, $iY, $WS_CAPTION + $WS_SYSMENU + $WS_MINIMIZEBOX, $WS_EX_TOPMOST)
	GUISetOnEvent($GUI_EVENT_CLOSE, "CloseElementNamesSelectionGUI")

	; Create list control for displaying element names
	Local $idList = GUICtrlCreateList("", 10, 10, $iW - 20, $iH - 80, $WS_VSCROLL)
	GUICtrlSetFont($idList, 9, 400, 0, "Courier New") ; Monospace font for better readability

	; Add element names to the list
	For $i = 0 To UBound($aNames) - 1
		GUICtrlSetData($idList, $aNames[$i])
	Next

	$g_ElementNamesSelectionList = $idList

	; Create selection button
	Local $idSelectBtn = GUICtrlCreateButton("Select", ($iW - 160) / 2, $iH - 60, 70, 25)
	GUICtrlSetOnEvent($idSelectBtn, "OnElementNameSelected")

	; Create cancel button
	Local $idCancelBtn = GUICtrlCreateButton("Cancel", ($iW - 160) / 2 + 90, $iH - 60, 70, 25)
	GUICtrlSetOnEvent($idCancelBtn, "CloseElementNamesSelectionGUI")

	; Make GUI moveable by dragging the title bar
	GUISetOnEvent($GUI_EVENT_PRIMARYDOWN, "_StartDragElementNamesSelectionGUI")

	GUISetState(@SW_SHOW, $g_ElementNamesSelectionGUI)

	Debug("Element names selection GUI displayed with " & UBound($aNames) & " names", "VERBOSE")
EndFunc   ;==>ShowElementNamesSelectionGUI

; Handles selection of an element name from the list
Func OnElementNameSelected()
	Local $selectedIndex = GUICtrlRead($g_ElementNamesSelectionList)
	If $selectedIndex = "" Or $selectedIndex = -1 Then Return

	; Get the selected text from the list
	Local $selectedText = GUICtrlRead($g_ElementNamesSelectionList)

	; Store the result
	$g_ElementNamesSelectionResult = $selectedText

	; Call the callback function if it exists
	If $g_ElementNamesSelectionCallback <> "" Then
		Call($g_ElementNamesSelectionCallback, $selectedText)
	EndIf

	; Close the selection GUI
	CloseElementNamesSelectionGUI()

	Debug("Element name selected: " & $selectedText, "VERBOSE")
EndFunc   ;==>OnElementNameSelected

; Closes the element names selection GUI
Func CloseElementNamesSelectionGUI()
	If $g_ElementNamesSelectionGUI <> 0 Then
		GUIDelete($g_ElementNamesSelectionGUI)
		$g_ElementNamesSelectionGUI = 0
		$g_ElementNamesSelectionList = 0
		$g_ElementNamesSelectionResult = ""
		$g_ElementNamesSelectionCallback = ""
		Debug("Element names selection GUI closed", "VERBOSE")
	EndIf
EndFunc   ;==>CloseElementNamesSelectionGUI

; Handles dragging of the element names selection GUI
Func _StartDragElementNamesSelectionGUI()
	If $g_ElementNamesSelectionGUI = 0 Then Return

	; Get mouse position relative to GUI
	Local $mousePos = MouseGetPos()
	Local $guiPos = WinGetPos($g_ElementNamesSelectionGUI)

	Local $offsetX = $mousePos[0] - $guiPos[0]
	Local $offsetY = $mousePos[1] - $guiPos[1]

	; Drag while mouse is down
	While _IsMouseDown()
		$mousePos = MouseGetPos()
		WinMove($g_ElementNamesSelectionGUI, "", $mousePos[0] - $offsetX, $mousePos[1] - $offsetY)
		Sleep(10)
	WEnd
EndFunc   ;==>_StartDragElementNamesSelectionGUI

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

	Debug("Element names GUI displayed with " & UBound($aNames) & " names", "VERBOSE")
EndFunc   ;==>ShowElementNamesGUI

; Closes the element names display GUI
Func CloseElementNamesGUI()
	If $g_ElementNamesGUI <> 0 Then
		GUIDelete($g_ElementNamesGUI)
		$g_ElementNamesGUI = 0
		$g_ElementNamesEdit = 0
		Debug("Element names GUI closed", "VERBOSE")
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
	Local $aState = DllCall("user32.dll", "int", "GetAsyncKeyState", "int", 0x01)
	If @error Or Not IsArray($aState) Then Return False
	Return BitAND($aState[0], 0x8000) <> 0
EndFunc   ;==>_IsMouseDown

; ================================================================================================
; MAIN BUTTON HANDLER
; ================================================================================================

; Button handler for "Get Element Names" button
Func GetElementNames()
	Debug("Get Element Names button clicked", "VERBOSE")

	; Check if Zoom meeting is in progress
	If Not FocusZoomWindow() Then
		Return
	EndIf

	; Meeting is in progress - collect element names
	Debug("Active Zoom meeting found, collecting element names...", "VERBOSE")

	; Get Zoom window object
	if not _GetZoomWindow() then return

	; Open Host Tools menu to collect names from it too
	Local $oHostMenu = _OpenHostTools()
	If Not IsObj($oHostMenu) Then
		Debug("Failed to open Host Tools menu, collecting names from Zoom window only", "WARN")
	EndIf

	; Collect element names from both windows
	Local $aNames = GetElementNamesFromWindows($oZoomWindow, $oHostMenu)

	If UBound($aNames) = 0 Then
		Debug(t("ERROR_ELEMENT_NOT_FOUND", t("ERROR_VARIOUS_ELEMENTS")), "ERROR")
		Return
	EndIf

	; Display the collected names in the GUI
	ShowElementNamesGUI($aNames)

	; Close Host Tools menu if it was opened
	If IsObj($oHostMenu) Then
		_CloseHostTools()
	EndIf

	Debug("Element names collection completed", "VERBOSE")
EndFunc   ;==>GetElementNames
