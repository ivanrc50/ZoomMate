#include-once

Global $oAccVars_Object = AccVars_Init(), $hAccVars_MethodFunc


; --- The first part of the UDF creates $oAccVars_Object and implements method functions ---

Func AccVars_Init()
	Local $sCLSID_ShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
	Local $sIID_IShellWindows = "{85CB6900-4D95-11CF-960C-0080C7F4EE85}"
	Local $stag_IShellWindows = _
		"VarToVariant01 hresult(variant*);"

	; Create AccVars object (ShellWindows object)
  Local $oAccVars_Object = ObjCreateInterface( $sCLSID_ShellWindows, $sIID_IShellWindows, $stag_IShellWindows )
  If Not IsObj( $oAccVars_Object ) Then Return SetError(1,0,1)

	; Replace original methods with VarToVariantXY methods
	Local $pVarToVariant, $pAccVars_Ptr = Ptr( $oAccVars_Object() ), $sFuncName, $sFuncParams = "ptr", $iOffset, $iPtrSize = @AutoItX64 ? 8 : 4
	For $i = 1 To 1
		$sFuncName = "VarToVariant" & StringFormat( "%02i", $i )
		$sFuncParams &= ";ptr*"
		$iOffset = ( 3 + $i - 1 ) * $iPtrSize
		$pVarToVariant = DllCallbackGetPtr( DllCallbackRegister( $sFuncName, "long", $sFuncParams ) )
		ReplaceVTableFuncPtr( $pAccVars_Ptr, $iOffset, $pVarToVariant )
	Next

	Return $oAccVars_Object
EndFunc

; Copied from "Hooking into the IDispatch interface" by monoceres
; https://www.autoitscript.com/forum/index.php?showtopic=107678
Func ReplaceVTableFuncPtr( $pVTable, $iOffset, $pNewFunc )
	Local $pPointer = DllStructGetData( DllStructCreate( "ptr", $pVTable ), 1 ) + $iOffset, $PAGE_EXECUTE_READWRITE = 0x40
	Local $pOldFunc = DllStructGetData( DllStructCreate( "ptr", $pPointer ), 1 ) ; Get the original function pointer in VTable
	Local $aRet = DllCall( "Kernel32.dll", "int", "VirtualProtect", "ptr", $pPointer, "long", @AutoItX64 ? 8 : 4, "dword", $PAGE_EXECUTE_READWRITE, "dword*", 0 ) ; Unprotect memory
	DllStructSetData( DllStructCreate( "ptr", $pPointer ), 1, $pNewFunc ) ; Replace function pointer in VTable with $pNewFunc function pointer
	DllCall( "Kernel32.dll", "int", "VirtualProtect", "ptr", $pPointer, "long", @AutoItX64 ? 8 : 4, "dword", $aRet[4], "dword*", 0 ) ; Protect memory
	Return $pOldFunc ; Return original function pointer
EndFunc

Func VarToVariant01( $pSelf, $pVar01 )
	$hAccVars_MethodFunc( $pVar01 )
	Return 0 ; $S_OK (COM constant)
	#forceref $pSelf
EndFunc


; --- The last part of the UDF creates a set of easy to use functions to call the object methods ---

Func AccVars01( $hAccVars_Method, ByRef $vVar01 )
	$hAccVars_MethodFunc = $hAccVars_Method
	$oAccVars_Object.VarToVariant01( $vVar01 )
EndFunc
