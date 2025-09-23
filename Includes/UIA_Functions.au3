#include-once
#include "UIA_Constants.au3"

; Error handling
;   @error = 0: No errors
;   @error = 1: Parameter error
;            2: Result error


; Used in Sample code to find application top window
Func UIA_GetApplicationWindow( $sWinVer, $oUIAutomation, $oDesktop, ByRef $oWindow, $aPropsVals )
	If $sWinVer <> "WIN_10_LAST" And $sWinVer <> "WIN_10" And $sWinVer <> "WIN_81" And $sWinVer <> "WIN_8" And $sWinVer <> "WIN_7" Then Return SetError(1,0,1)
	If Not IsObj( $oUIAutomation ) Then Return SetError(1,0,1)
	If Not IsObj( $oDesktop ) Then Return SetError(1,0,1)
	If Not IsArray( $aPropsVals ) Then Return SetError(1,0,1)
	If UBound( $aPropsVals, 2 ) <> 2 Then Return SetError(1,0,1)

	Local $pTrueCondition
	$oUIAutomation.CreateTrueCondition( $pTrueCondition )

	Local $pElements
	$oDesktop.FindAll( $TreeScope_Children, $pTrueCondition, $pElements ) ; Top windows only
	If Not $pElements Then Return SetError(2,0,1)

	Local $oAutomationElementArray, $iLength
	$oAutomationElementArray = ObjCreateInterFace( $pElements, $sIID_IUIAutomationElementArray, $dtag_IUIAutomationElementArray )
	$oAutomationElementArray.Length( $iLength )
	If Not $iLength Then Return SetError(2,0,1)

	Local $sIIDIUIAutomationElement, $dtagIUIAutomationElement

	Switch $sWinVer
		Case "WIN_10_LAST"
			$sIIDIUIAutomationElement = $sIID_IUIAutomationElement9
			$dtagIUIAutomationElement = $dtag_IUIAutomationElement9
		Case "WIN_10"
			$sIIDIUIAutomationElement = $sIID_IUIAutomationElement4
			$dtagIUIAutomationElement = $dtag_IUIAutomationElement4
		Case "WIN_81"
			$sIIDIUIAutomationElement = $sIID_IUIAutomationElement3
			$dtagIUIAutomationElement = $dtag_IUIAutomationElement3
		Case "WIN_8"
			$sIIDIUIAutomationElement = $sIID_IUIAutomationElement2
			$dtagIUIAutomationElement = $dtag_IUIAutomationElement2
		Case "WIN_7"
			$sIIDIUIAutomationElement = $sIID_IUIAutomationElement
			$dtagIUIAutomationElement = $dtag_IUIAutomationElement
	EndSwitch

	Local $pWindow, $sValue
	Local $iProps = UBound( $aPropsVals )
	For $i = 0 To $iLength - 1
		$oAutomationElementArray.GetElement( $i, $pWindow )
		$oWindow = ObjCreateInterface( $pWindow, $sIIDIUIAutomationElement, $dtagIUIAutomationElement )
		For $j = 0 To $iProps - 1
			$oWindow.GetCurrentPropertyValue( $aPropsVals[$j][0], $sValue )
			If $sValue <> $aPropsVals[$j][1] Then ContinueLoop 2
		Next
		ExitLoop
	Next
	If $i = $iLength Then Return SetError(2,0,1)
EndFunc


; Input: Array
; Output: String
Func UIA_GetArrayPropertyValueAsString( ByRef $sArrToStr )
	If Not IsArray( $sArrToStr ) Then Return SetError(1,0,1)

	Local $sValue = $sArrToStr[0]
	For $i = 1 To UBound( $sArrToStr ) - 1
		$sValue &= "," & $sArrToStr[$i]
	Next
	$sArrToStr = $sValue
EndFunc


; $fScale is the screen scaling factor
; $fScale values: 1.25, 1.50, ... 4.00
Func UIA_MouseClick( $oElement, $bRight = False, $fScale = 0 )
	If Not IsObj( $oElement ) Then Return SetError(1,0,1)

	; Rectangle
	Local $aRect ; l, t, w, h
	$oElement.GetCurrentPropertyValue( $UIA_BoundingRectanglePropertyId, $aRect )
	If Not IsArray( $aRect ) Then Return SetError(1,0,1)
	If $fScale > 1.00 Then
		$aRect[0] = Round( $aRect[0] * $fScale )
		$aRect[1] = Round( $aRect[1] * $fScale )
		$aRect[2] = Round( $aRect[2] * $fScale )
		$aRect[3] = Round( $aRect[3] * $fScale )
	EndIf

	; Click element
	Local $aPos = MouseGetPos()
	Local $sButton = Not $bRight ? "primary" : "secondary"
	DllCall( "user32.dll", "int", "ShowCursor", "bool", False )
	MouseClick( $sButton, $aRect[0]+$aRect[2]/2, $aRect[1]+$aRect[3]/2, 1, 0 )
	MouseMove( $aPos[0], $aPos[1], 0 )
	DllCall( "user32.dll", "int", "ShowCursor", "bool", True )
EndFunc


; $oWindow
;   Window object ($oWindow)
;   Window handle ($hWindow)
Func UIA_PostMessage( $oWindow, $sAutomationId )
	If Not ( IsObj( $oWindow ) Or IsHWnd( $oWindow ) ) Then Return SetError(1,0,1)
	If Not Int( $sAutomationId ) Then Return SetError(1,0,1)
	Local $WM_COMMAND = 0x0111

	If IsHWnd( $oWindow ) Then Return _
		DllCall( "user32.dll", "bool", "PostMessage", "hwnd", $oWindow, "uint", $WM_COMMAND, "wparam", Int( $sAutomationId ), "lparam", 0 )[0]

	Local $hNativeWindowHandle
	$oWindow.GetCurrentPropertyValue( $UIA_NativeWindowHandlePropertyId, $hNativeWindowHandle )
	DllCall( "user32.dll", "bool", "PostMessage", "hwnd", HWnd( $hNativeWindowHandle ), "uint", $WM_COMMAND, "wparam", Int( $sAutomationId ), "lparam", 0 )
EndFunc


; Returns a String from a BSTR
; Copied from Variant.au3 in this example:
; https://www.autoitscript.com/forum/index.php?showtopic=200660
Func UIA_String( $pBSTR, $iLen = -1 )
	If Not $pBSTR Then Return SetError(1, 0, "")
	If $iLen < 1 Then $iLen = DllCall( "OleAut32.dll", "uint", "SysStringLen", "ptr", $pBSTR )[0]
	If $iLen < 1 Then Return SetError(2, 0, "")
	Return DllStructGetData( DllStructCreate( "wchar[" & $iLen & "]", $pBSTR ), 1 )
EndFunc


; $oWindow
;   Window object ($oWindow)
;   Window handle ($hWindow)
Func UIA_WinActivate( $oWindow )
	If Not ( IsObj( $oWindow ) Or IsHWnd( $oWindow ) ) Then Return SetError(1,0,1)

	If IsHWnd( $oWindow ) Then Return WinActivate( $oWindow )

	Local $hNativeWindowHandle
	$oWindow.GetCurrentPropertyValue( $UIA_NativeWindowHandlePropertyId, $hNativeWindowHandle )
	WinActivate( HWnd( $hNativeWindowHandle ) )
EndFunc


; $oWindow
;   Window object ($oWindow)
;   Window handle ($hWindow)
Func UIA_WinWaitActive( $oWindow, $iTimeout = 0 )
	If Not ( IsObj( $oWindow ) Or IsHWnd( $oWindow ) ) Then Return SetError(1,0,1)

	If IsHWnd( $oWindow ) Then Return WinWaitActive( $oWindow, "", $iTimeout )

	Local $hNativeWindowHandle
	$oWindow.GetCurrentPropertyValue( $UIA_NativeWindowHandlePropertyId, $hNativeWindowHandle )
	WinWaitActive( HWnd( $hNativeWindowHandle ), "", $iTimeout )
EndFunc
