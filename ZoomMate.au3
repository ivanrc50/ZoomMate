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

; -------------------------
; Script configuration
; -------------------------
Global Const $CONFIG_FILE = @ScriptDir & "\zoom_config.ini"
Global Const $MAX_WAIT_TIME = 30000  ; 30 seconds

; -------------------------
; Runtime variables
; -------------------------
Global $g_PrePostSettingsConfigured = False
Global $g_DuringMeetingSettingsConfigured = False

; Default meeting configuration
Global $meetingID, $midweekDay, $midweekTime, $weekendDay, $weekendTime

; GUI globals
Global $msg, $idMeetingID, $idMidweekDay, $idMidweekTime, $idWeekendDay, $idWeekendTime, $idSaveBtn, $idQuitBtn
; -------------------------
; Initialize UIAutomation COM
; -------------------------
Global $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation)
If Not IsObj($oUIAutomation) Then
	MsgBox($MB_ICONERROR, "UIAutomation", "Failed to create UIAutomation COM object.")
	Exit
EndIf
Debug("UIAutomation COM created successfully.", "UIA")

Global $pDesktop
$oUIAutomation.GetRootElement($pDesktop)
Global $oDesktop = ObjCreateInterface($pDesktop, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
If Not IsObj($oDesktop) Then
	MsgBox($MB_ICONERROR, "UIAutomation", "Failed to get desktop element.")
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

Func Debug($string, $type = "DEBUG")
	If ($string) Then
		ConsoleWrite("[" & $type & "] " & $string & @CRLF)
		If $type = "INFO" Or $type = "ERROR" Then
			$g_StatusMsg = $string
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

	$g_ConfigGUI = GUICreate("ZoomMate Configuration", 320, 220)
	GUICtrlCreateLabel("Zoom Meeting ID:", 10, 10, 120, 20)
	$idMeetingID = GUICtrlCreateInput($meetingID, 140, 10, 160, 20)

	GUICtrlCreateLabel("Midweek Day (1=Sun,..7=Sat):", 10, 40, 180, 20)
	$idMidweekDay = GUICtrlCreateInput($midweekDay, 200, 40, 100, 20)

	GUICtrlCreateLabel("Midweek Time (HH:MM):", 10, 70, 180, 20)
	$idMidweekTime = GUICtrlCreateInput($midweekTime, 200, 70, 100, 20)

	GUICtrlCreateLabel("Weekend Day (1=Sun,..7=Sat):", 10, 100, 180, 20)
	$idWeekendDay = GUICtrlCreateInput($weekendDay, 200, 100, 100, 20)

	GUICtrlCreateLabel("Weekend Time (HH:MM):", 10, 130, 180, 20)
	$idWeekendTime = GUICtrlCreateInput($weekendTime, 200, 130, 100, 20)

	$idSaveBtn = GUICtrlCreateButton("Save", 60, 170, 80, 30)
	$idQuitBtn = GUICtrlCreateButton("Quit", 180, 170, 80, 30)

	; Save button is disabled by default
	GUICtrlSetState($idSaveBtn, $GUI_DISABLE)
	GUICtrlSetState($GUI_EVENT_CLOSE, $GUI_DISABLE)
	GUICtrlSetState($idQuitBtn, $GUI_ENABLE)

	GUICtrlSetOnEvent($idSaveBtn, "SaveConfigGUI")
	GUICtrlSetOnEvent($idQuitBtn, "QuitApp")

	; Enable Save button only if all fields have a value
	GUICtrlSetOnEvent($idMeetingID, "CheckConfigFields")
	GUICtrlSetOnEvent($idMidweekDay, "CheckConfigFields")
	GUICtrlSetOnEvent($idMidweekTime, "CheckConfigFields")
	GUICtrlSetOnEvent($idWeekendDay, "CheckConfigFields")
	GUICtrlSetOnEvent($idWeekendTime, "CheckConfigFields")

	CheckConfigFields() ; Initial check


	GUISetState(@SW_SHOW, $g_ConfigGUI)
EndFunc   ;==>ShowConfigGUI

Func CheckConfigFields()
	Local $v1 = StringStripWS(GUICtrlRead($idMeetingID), 3)
	Local $v2 = StringStripWS(GUICtrlRead($idMidweekDay), 3)
	Local $v3 = StringStripWS(GUICtrlRead($idMidweekTime), 3)
	Local $v4 = StringStripWS(GUICtrlRead($idWeekendDay), 3)
	Local $v5 = StringStripWS(GUICtrlRead($idWeekendTime), 3)
	If $v1 <> "" And $v2 <> "" And $v3 <> "" And $v4 <> "" And $v5 <> "" Then
		GUICtrlSetState($idSaveBtn, $GUI_ENABLE)
	Else
		GUICtrlSetState($idSaveBtn, $GUI_DISABLE)
	EndIf
EndFunc   ;==>CheckConfigFields

Func SaveConfigGUI()
	$meetingID = GUICtrlRead($idMeetingID)
	$midweekDay = GUICtrlRead($idMidweekDay)
	$midweekTime = GUICtrlRead($idMidweekTime)
	$weekendDay = GUICtrlRead($idWeekendDay)
	$weekendTime = GUICtrlRead($idWeekendTime)

	IniWrite($CONFIG_FILE, "ZoomSettings", "MeetingID", $meetingID)
	IniWrite($CONFIG_FILE, "Meetings", "MidweekDay", $midweekDay)
	IniWrite($CONFIG_FILE, "Meetings", "MidweekTime", $midweekTime)
	IniWrite($CONFIG_FILE, "Meetings", "WeekendDay", $weekendDay)
	IniWrite($CONFIG_FILE, "Meetings", "WeekendTime", $weekendTime)

	CloseConfigGUI()
EndFunc   ;==>SaveConfigGUI

Func CloseConfigGUI()
	GUIDelete($g_ConfigGUI)
	$g_ConfigGUI = 0
EndFunc   ;==>CloseConfigGUI

Func QuitApp()
	Exit
EndFunc   ;==>QuitApp

Func LoadMeetingConfig()
	$meetingID = IniRead($CONFIG_FILE, "ZoomSettings", "MeetingID", "")
	$midweekDay = IniRead($CONFIG_FILE, "Meetings", "MidweekDay", "")
	$midweekTime = IniRead($CONFIG_FILE, "Meetings", "MidweekTime", "")
	$weekendDay = IniRead($CONFIG_FILE, "Meetings", "WeekendDay", "")
	$weekendTime = IniRead($CONFIG_FILE, "Meetings", "WeekendTime", "")

	If $meetingID = "" Or $midweekDay = "" Or $midweekTime = "" Or $weekendDay = "" Or $weekendTime = "" Then
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

LoadMeetingConfig()

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

	If $today = Number($midweekDay) Then
		CheckMeetingWindow($midweekTime)
		ResponsiveSleep(5)
	ElseIf $today = Number($weekendDay) Then
		CheckMeetingWindow($weekendTime)
		ResponsiveSleep(5)
	Else
		Debug("Waiting for the next meeting day...", "INFO") ; Not a meeting day. Sleeping 1 hour.
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
				Debug("Error launching Zoom", "ERROR")
			Else
				Debug("Configuring settings for before and after meetings...", "INFO")
				_SetPreAndPostMeetingSettings()
				$g_PrePostSettingsConfigured = True
				Debug("Settings configured for before and after meetings.", "INFO")
			EndIf
		EndIf

	ElseIf $nowMin = ($meetingMin - 1) Then
		; 1 minute before meeting → During meeting settings
		If Not $g_DuringMeetingSettingsConfigured Then
			Local $zoomExists = IsObj(_GetZoomWindow())
			If Not $zoomExists Then
				Debug("Zoom window not found", "ERROR")
			Else
				Debug("Meeting starting soon... Configuring settings.", "INFO")
				_SetDuringMeetingSettings()
				$g_DuringMeetingSettingsConfigured = True
				Debug("Settings configured for during the meeting.", "INFO")
			EndIf
		EndIf

	ElseIf $nowMin >= $meetingMin Then
		; Meeting already started
		Local $minutesAgo = $nowMin - $meetingMin
		If $minutesAgo <= 120 Then
			Debug("Meeting started " & $minutesAgo & " minute(s) ago.", "INFO")
		Else
			Debug("Outside of meeting window. Meeting started more than 2 hours ago.", "INFO")
		EndIf

	Else
		; Too early - show countdown
		Local $minutesLeft = $meetingMin - $nowMin
		Debug("Meeting starting in " & $minutesLeft & " minute(s).", "INFO")
	EndIf
EndFunc   ;==>CheckMeetingWindow

; -------------------------
; Launch Zoom
; -------------------------
Func _LaunchZoom()
	Debug("Launching Zoom...", "INFO")

	Local $meetingID = IniRead($CONFIG_FILE, "ZoomSettings", "MeetingID", "")
	If $meetingID = "" Then
		Debug("Meeting ID not configured.", "ERROR")
		Return SetError(1, 0, 0)
	EndIf

	Local $zoomURL = "zoommtg://zoom.us/join?confno=" & $meetingID
	ShellExecute($zoomURL)
	Debug("Zoom launched: " & $meetingID, "INFO")
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
	If Not IsObj($oElement) Then
		Debug("Invalid element object.", "ERROR")
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
	Debug("Failed to click element: '" & $sElementName & "' - no suitable pattern found.", "ERROR")
	Return False
EndFunc   ;==>_ClickElement

Func _OpenHostTools()
	If Not IsObj($oZoomWindow) Then Return False
	Local $oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
	If Not IsObj($oHostMenu) Then
		Local $oButton = FindElementByPartialName("Host tools", Default, $oZoomWindow)
		If Not _ClickElement($oButton) Then
			Debug("Failed to click Host Tools.", "WARN")
			Return False
		EndIf
	EndIf
	$oHostMenu = FindElementByClassName("WCN_ModelessWnd", Default, $oZoomWindow)
	Return $oHostMenu
EndFunc   ;==>_OpenHostTools

Func _CloseHostTools()
	; Click on the main window to hide the Host tools menu
	_ClickElement($oZoomWindow, True)
EndFunc   ;==>_CloseHostTools

Func _OpenParticipantsPanel()
	If Not IsObj($oZoomWindow) Then Return False
	Local $ListType[1] = [$UIA_ListControlTypeId]
	Local $oParticipantsPanel = FindElementByPartialName("Participant", $ListType, $oZoomWindow)

	If Not IsObj($oParticipantsPanel) Then
		Local $oButton = FindElementByPartialName("Participant", Default, $oZoomWindow)
		If Not _ClickElement($oButton) Then
			Debug("Failed to click Participants.", "WARN")
			Return False
		EndIf
	EndIf
	$oParticipantsPanel = FindElementByPartialName("Participant", $ListType, $oZoomWindow)
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
			Debug("Setting '" & $sSetting & "' not found.", "ERROR")
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
		Debug("Unknown feed type: " & $feedType, "ERROR")
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
	Local $oParticipantsPanel = _OpenParticipantsPanel()
	If Not IsObj($oParticipantsPanel) Then Return False
	Local $oButton = FindElementByPartialName("Mute all", Default, $oZoomWindow)
	If Not _ClickElement($oButton) Then
		Debug("'Mute all' button not found.", "WARN")
		Return False
	EndIf

	Return DialogClick("zChangeNameWndClass", "Yes")
EndFunc   ;==>MuteAll

Func DialogClick($ClassName, $ButtonLabel)
	; Confirm dialog
	Local $oDialog = FindElementByClassName($ClassName)
	Local $oButton = FindElementByPartialName($ButtonLabel, Default, $oDialog)
	If _ClickElement($oButton) Then Return True
	Debug("Confirmation 'Yes' not found.", "WARN")
	Return False

EndFunc   ;==>DialogClick

Func _GetZoomStatus()
	Local $aSettings[2] = ["Unmute themselves", "Share screen"]
	Local $oStates = GetSecuritySettingsState($aSettings)

	; Example of accessing values
	For $sKey In $oStates.Keys
		Debug($sKey & " = " & ($oStates.Item($sKey) ? "ENABLED" : "DISABLED"), "ZOOM STATUS")
	Next
EndFunc   ;==>_GetZoomStatus

