#include-once
#include "UIA_Constants.au3"

Global $aDockPositionNames
Func DockPositionSetNames()
	If IsArray( $aDockPositionNames ) Then Return
	Local $aDockPositionsTmp = [ _
		"$DockPosition_Top", _
		"$DockPosition_Left", _
		"$DockPosition_Bottom", _
		"$DockPosition_Right", _
		"$DockPosition_Fill", _
		"$DockPosition_None" ]
	$aDockPositionNames = $aDockPositionsTmp
EndFunc

Global $aExpandCollapseStateNames
Func ExpandCollapseStateSetNames()
	If IsArray( $aExpandCollapseStateNames ) Then Return
	Local $aExpandCollapseStatesTmp = [ _
		"$ExpandCollapseState_Collapsed", _
		"$ExpandCollapseState_Expanded", _
		"$ExpandCollapseState_PartiallyExpanded", _
		"$ExpandCollapseState_LeafNode" ]
	$aExpandCollapseStateNames = $aExpandCollapseStatesTmp
EndFunc

Global $aLiveSettingNames
Func LiveSettingSetNames()
	If IsArray( $aLiveSettingNames ) Then Return
	Local $aLiveSettingsTmp = [ _
		"$LiveSetting_Off", _
		"$LiveSetting_Polite", _
		"$LiveSetting_Assertive" ]
	$aLiveSettingNames = $aLiveSettingsTmp
EndFunc

Global $aNotificationKindNames
Func NotificationKindSetNames()
	If IsArray( $aNotificationKindNames ) Then Return
	Local $aNotificationKindsTmp = [ _
		"$NotificationKind_ItemAdded", _
		"$NotificationKind_ItemRemoved ", _
		"$NotificationKind_ActionCompleted", _
		"$NotificationKind_ActionAborted", _
		"$NotificationKind_Other" ]
	$aNotificationKindNames = $aNotificationKindsTmp
EndFunc

Global $aNotificationProcessingNames
Func NotificationProcessingSetNames()
	If IsArray( $aNotificationProcessingNames ) Then Return
	Local $aNotificationProcessingsTmp = [ _
	"$NotificationProcessing_ImportantAll", _
	"$NotificationProcessing_ImportantMostRecent", _
	"$NotificationProcessing_All", _
	"$NotificationProcessing_MostRecent", _
	"$NotificationProcessing_CurrentThenMostRecent" ]
	$aNotificationProcessingNames = $aNotificationProcessingsTmp
EndFunc

Global $aOrientationTypeNames
Func OrientationTypeSetNames()
	If IsArray( $aOrientationTypeNames ) Then Return
	Local $aOrientationTypesTmp = [ _
		"$OrientationType_None", _
		"$OrientationType_Horizontal", _
		"$OrientationType_Vertical" ]
	$aOrientationTypeNames = $aOrientationTypesTmp
EndFunc

Global $aToggleStateNames
Func ToggleStateSetNames()
	If IsArray( $aToggleStateNames ) Then Return
	Local $aToggleStatesTmp = [ _
		"$ToggleState_Off", _
		"$ToggleState_On", _
		"$ToggleState_Indeterminate" ]
	$aToggleStateNames = $aToggleStatesTmp
EndFunc

Global $aWindowInteractionStateNames
Func WindowInteractionStateSetNames()
	If IsArray( $aWindowInteractionStateNames ) Then Return
	Local $aWindowInteractionStatesTmp = [ _
		"$WindowInteractionState_Running", _
		"$WindowInteractionState_Closing", _
		"$WindowInteractionState_ReadyForUserInteraction", _
		"$WindowInteractionState_BlockedByModalWindow", _
		"$WindowInteractionState_NotResponding" ]
	$aWindowInteractionStateNames = $aWindowInteractionStatesTmp
EndFunc

Global $aWindowVisualStateNames
Func WindowVisualStateSetNames()
	If IsArray( $aWindowVisualStateNames ) Then Return
	Local $aWindowVisualStatesTmp = [ _
		"$WindowVisualState_Normal", _
		"$WindowVisualState_Maximized", _
		"$WindowVisualState_Minimized" ]
	$aWindowVisualStateNames = $aWindowVisualStatesTmp
EndFunc
