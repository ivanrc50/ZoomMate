; ================================================================================================
; ZoomMate Translation Data - All Languages
; ================================================================================================

Global Const $TRANSLATIONS = ObjCreate("Scripting.Dictionary")

; Initialize translations for each language
_InitializeTranslations()

Func _InitializeTranslations()
    Local $en = ObjCreate("Scripting.Dictionary")
    Local $es = ObjCreate("Scripting.Dictionary")  ; Spanish placeholder for future use

    ; Language metadata
    $en.Add("LANGNAME", "English")

    ; Configuration GUI
    $en.Add("CONFIG_TITLE", "ZoomMate Configuration")
    $en.Add("BTN_SAVE", "Save")
    $en.Add("BTN_QUIT", "Quit ZoomMate")

    ; Status messages
    $en.Add("TOOLTIP_IDLE", "Idle")
    $en.Add("INFO_ZOOM_LAUNCHING", "Launching Zoom...")
    $en.Add("INFO_ZOOM_LAUNCHED", "Zoom meeting launched")
    $en.Add("INFO_MEETING_STARTING_IN", "Meeting starting in {0} minute(s).")
    $en.Add("INFO_MEETING_STARTED_AGO", "Meeting started {0} minute(s) ago.")
    $en.Add("INFO_CONFIG_BEFORE_AFTER_START", "Configuring settings for before and after meetings...")
    $en.Add("INFO_CONFIG_BEFORE_AFTER_DONE", "Settings configured for before and after meetings.")
    $en.Add("INFO_MEETING_STARTING_SOON_CONFIG", "Meeting starting soon... Configuring settings.")
    $en.Add("INFO_CONFIG_DURING_MEETING_DONE", "Settings configured for during the meeting.")
    $en.Add("INFO_OUTSIDE_MEETING_WINDOW", "Outside of meeting window. Meeting started more than 2 hours ago.")
    $en.Add("INFO_CONFIG_LOADED", "Configuration loaded successfully.")
    $en.Add("INFO_NO_MEETING_SCHEDULED", "No meeting scheduled for today. Waiting for the next meeting day...")

    ; Labels
    $en.Add("LABEL_MEETING_ID", "Zoom Meeting ID")
    $en.Add("LABEL_MIDWEEK_DAY", "Midweek Day")
    $en.Add("LABEL_MIDWEEK_TIME", "Midweek Time (HH:MM)")
    $en.Add("LABEL_WEEKEND_DAY", "Weekend Day")
    $en.Add("LABEL_WEEKEND_TIME", "Weekend Time (HH:MM)")
    $en.Add("LABEL_LANGUAGE", "Language:")

    ; Zoom interface labels
    $en.Add("LABEL_HOST_TOOLS", "Host tools")
    $en.Add("LABEL_MORE_MEETING_CONTROLS", "More meeting controls")
    $en.Add("LABEL_PARTICIPANT", "Participant")
    $en.Add("LABEL_MUTE_ALL", "Mute All")
    $en.Add("LABEL_YES", "Yes")
    $en.Add("LABEL_UNCHECKED_VALUE", "Unchecked")
    $en.Add("LABEL_CURRENTLY_UNMUTED_VALUE", "Currently unmuted")
    $en.Add("LABEL_UNMUTE_AUDIO_VALUE", "Unmute my audio")
    $en.Add("LABEL_STOP_VIDEO_VALUE", "Stop my video")
    $en.Add("LABEL_START_VIDEO_VALUE", "Start my video")
    $en.Add("LABEL_ZOOM_SECURITY_UNMUTE", "Unmute themselves")
    $en.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN", "Share screen")

    ; Help text
    $en.Add("LABEL_HOST_TOOLS_EXPLAIN", "Enter the text that appears on the Host Tools button in your Zoom interface. This is used to locate and click the button automatically.")
    $en.Add("LABEL_PARTICIPANT_EXPLAIN", "Enter the text that appears on the Participants button in your Zoom interface. This is used to locate and open the participants panel.")
    $en.Add("LABEL_MUTE_ALL_EXPLAIN", "Enter the text that appears on the Mute All button in your Zoom interface. This is used to mute all participants automatically.")
    $en.Add("LABEL_YES_EXPLAIN", "Enter the text that appears on confirmation buttons (e.g., ""Yes"", ""OK"") in your Zoom interface. This is used to confirm actions.")
    $en.Add("LABEL_UNCHECKED_VALUE_EXPLAIN", "Enter the text that appears when a setting is unchecked/disabled in your Zoom interface. This is used to detect when settings are disabled.")
    $en.Add("LABEL_CURRENTLY_UNMUTED_VALUE_EXPLAIN", "Enter the text that appears on the audio button when you are currently unmuted in your Zoom interface. This is used to detect audio state.")
    $en.Add("LABEL_UNMUTE_AUDIO_VALUE_EXPLAIN", "Enter the text that appears on the button to unmute your audio in your Zoom interface. This is used to unmute yourself.")
    $en.Add("LABEL_STOP_VIDEO_VALUE_EXPLAIN", "Enter the text that appears on the button to stop your video in your Zoom interface. This is used to stop your video feed.")
    $en.Add("LABEL_START_VIDEO_VALUE_EXPLAIN", "Enter the text that appears on the button to start your video in your Zoom interface. This is used to start your video feed.")
    $en.Add("LABEL_ZOOM_SECURITY_UNMUTE_EXPLAIN", "Enter the text that appears on the Unmute permission setting in Zoom Security settings. This controls whether participants can unmute themselves.")
    $en.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN_EXPLAIN", "Enter the text that appears on the Share screen permission setting in Zoom Security settings. This controls whether participants can share their screen.")

    ; Settings
    $en.Add("LABEL_SNAP_ZOOM_TO", "Snap Zoom window to")
    $en.Add("SNAP_DISABLED", "Disabled")
    $en.Add("SNAP_LEFT", "Left")
    $en.Add("SNAP_RIGHT", "Right")
    $en.Add("LABEL_KEYBOARD_SHORTCUT", "Post-meeting Keyboard Shortcut")
    $en.Add("LABEL_KEYBOARD_SHORTCUT_EXPLAIN", "Enter a keyboard shortcut that will apply post-meeting settings (e.g., Ctrl+Alt+Z). Use ^ for Ctrl, ! for Alt, + for Shift, # for Win, followed by a letter or number.")

    ; Error messages
    $en.Add("ERROR_GET_DESKTOP_ELEMENT_FAILED", "Failed to get desktop element.")
    $en.Add("ERROR_ZOOM_LAUNCH", "Error launching Zoom")
    $en.Add("ERROR_ZOOM_WINDOW_NOT_FOUND", "Zoom window not found")
    $en.Add("ERROR_MEETING_ID_NOT_CONFIGURED", "Meeting ID not configured.")
    $en.Add("ERROR_MEETING_ID_FORMAT", "Enter 9–11 digits (no spaces)")
    $en.Add("ERROR_TIME_FORMAT", "Use 24h time HH:MM")
    $en.Add("ERROR_KEYBOARD_SHORTCUT_FORMAT", "Use format like ^!z (Ctrl+Alt+Z). Must include at least one modifier (^ Ctrl, ! Alt, + Shift, # Win) followed by a letter or number.")
    $en.Add("ERROR_REQUIRED", "This field is required")
    $en.Add("ERROR_FIELDS_REQUIRED", "Please complete all required fields")
    $en.Add("ERROR_INVALID_ELEMENT_OBJECT", "Invalid element object.")
    $en.Add("ERROR_FAILED_CLICK_ELEMENT", "Failed to click element")
    $en.Add("ERROR_SETTING_NOT_FOUND", "Setting not found")
    $en.Add("ERROR_UNKNOWN_FEED_TYPE", "Unknown feed type")

    ; Overlay messages
    $en.Add("PLEASE_WAIT_TITLE", "Please Wait")
    $en.Add("PLEASE_WAIT_TEXT", "Please wait...")
    $en.Add("POST_MEETING_HIT_KEY_TITLE", "Post-Meeting Settings")
    $en.Add("POST_MEETING_HIT_KEY_TEXT", "Are you ready to apply post-meeting settings? Press ENTER when the prayer is over to apply them, or ESC to cancel.")

    ; Section headers
    $en.Add("SECTION_MEETING_INFO", "Meeting Information")
    $en.Add("SECTION_ZOOM_LABELS", "Zoom Interface Labels")
    $en.Add("SECTION_GENERAL_SETTINGS", "General Settings")

    ; Day labels (1=Sunday .. 7=Saturday)
    $en.Add("DAY_1", "Sunday")
    $en.Add("DAY_2", "Monday")
    $en.Add("DAY_3", "Tuesday")
    $en.Add("DAY_4", "Wednesday")
    $en.Add("DAY_5", "Thursday")
    $en.Add("DAY_6", "Friday")
    $en.Add("DAY_7", "Saturday")

    ; Spanish translations (placeholder for future implementation)
    $es.Add("LANGNAME", "Español")
    $es.Add("CONFIG_TITLE", "Configuración de ZoomMate")
    $es.Add("BTN_SAVE", "Guardar")
    $es.Add("BTN_QUIT", "Salir de ZoomMate")
    ; Add more Spanish translations as needed...

    ; Add language dictionaries to main translations object
    $TRANSLATIONS.Add("en", $en)
    $TRANSLATIONS.Add("es", $es)
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
