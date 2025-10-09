; ================================================================================================
; ZoomMate Translation Manager - Imports all language files (5 languages supported)
; ================================================================================================

#include "en.au3"
#include "es.au3"
#include "fr.au3"
#include "ru.au3"
#include "uk.au3"

Global Const $TRANSLATIONS = ObjCreate("Scripting.Dictionary")

; Initialize translations for each language
_InitializeTranslations()

Func _InitializeTranslations()
    ; Add each language's translations to the main dictionary
    $TRANSLATIONS.Add("en", $TRANSLATIONS_EN)
    $TRANSLATIONS.Add("es", $TRANSLATIONS_ES)
    $TRANSLATIONS.Add("fr", $TRANSLATIONS_FR)
    $TRANSLATIONS.Add("ru", $TRANSLATIONS_RU)
    $TRANSLATIONS.Add("uk", $TRANSLATIONS_UK)
EndFunc   ;==>_InitializeTranslations

; Helper function to get translations for a specific language
; @param $langCode - Language code (e.g., "en", "es")
; @return Object - Dictionary containing translations for the specified language
Func _GetLanguageTranslations($langCode)
    ; Ensure translations are initialized
    If $TRANSLATIONS.Count = 0 Then _InitializeTranslations()

    If $TRANSLATIONS.Exists($langCode) Then
        Return $TRANSLATIONS.Item($langCode)
    Else
        ; Fallback to English if requested language not found
        Return $TRANSLATIONS.Item("en")
    EndIf
EndFunc   ;==>_GetLanguageTranslations

; Builds a comma-separated list of available language display names
; @return String - Comma-separated list of language names
Func _ListAvailableLanguageNames()
    Local $list = ""
    ; Ensure translations are initialized
    If $TRANSLATIONS.Count = 0 Then _InitializeTranslations()

    For $langCode In $TRANSLATIONS.Keys
        Local $translations = _GetLanguageTranslations($langCode)
        If $translations.Exists("LANGNAME") Then
            Local $langName = $translations.Item("LANGNAME")
            $list &= ($list = "" ? $langName : "|" & $langName)
        EndIf
    Next
    Return $list
EndFunc   ;==>_ListAvailableLanguageNames

; Gets the language code for a display name
; @param $displayName - Display name (e.g., "English", "Español")
; @return String - Language code or empty string if not found
Func _GetLanguageCodeFromDisplayName($displayName)
    ; Ensure translations are initialized
    If $TRANSLATIONS.Count = 0 Then _InitializeTranslations()

    For $langCode In $TRANSLATIONS.Keys
        Local $translations = _GetLanguageTranslations($langCode)
        If $translations.Exists("LANGNAME") And $translations.Item("LANGNAME") = $displayName Then
            Return $langCode
        EndIf
    Next
    Return ""  ; Not found
EndFunc   ;==>_GetLanguageCodeFromDisplayName

; Gets the display name for a language code
; @param $code - Language code (e.g., "en", "es")
; @return String - Display name or the code itself if not found
Func _GetLanguageDisplayName($code)
    Local $translations = _GetLanguageTranslations($code)
    If $translations.Exists("LANGNAME") Then
        Return $translations.Item("LANGNAME")
    EndIf
    Return $code
EndFunc   ;==>_GetLanguageDisplayName
