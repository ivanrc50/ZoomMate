; ================================================================================================
; ZoomMate English Translation Data
; ================================================================================================

; Language metadata
Global $TRANSLATIONS_EN = ObjCreate("Scripting.Dictionary")
$TRANSLATIONS_EN.Add("LANGNAME", "English")

; Configuration GUI
$TRANSLATIONS_EN.Add("CONFIG_TITLE", "ZoomMate Configuration")
$TRANSLATIONS_EN.Add("BTN_SAVE", "Save")
$TRANSLATIONS_EN.Add("BTN_QUIT", "Quit ZoomMate")
$TRANSLATIONS_EN.Add("LABEL_LANGUAGE", "Language:")

; Status messages
$TRANSLATIONS_EN.Add("TOOLTIP_IDLE", "Idle")
$TRANSLATIONS_EN.Add("INFO_ZOOM_LAUNCHING", "Launching Zoom...")
$TRANSLATIONS_EN.Add("INFO_ZOOM_LAUNCHED", "Zoom meeting launched")
$TRANSLATIONS_EN.Add("INFO_MEETING_STARTING_IN", "Meeting starting in {0} minute(s).")
$TRANSLATIONS_EN.Add("INFO_MEETING_STARTED_AGO", "Meeting started {0} minute(s) ago.")
$TRANSLATIONS_EN.Add("INFO_CONFIG_BEFORE_AFTER_START", "Configuring settings for before and after meetings...")
$TRANSLATIONS_EN.Add("INFO_CONFIG_BEFORE_AFTER_DONE", "Settings configured for before and after meetings.")
$TRANSLATIONS_EN.Add("INFO_MEETING_STARTING_SOON_CONFIG", "Meeting starting soon... Configuring settings.")
$TRANSLATIONS_EN.Add("INFO_CONFIG_DURING_MEETING_DONE", "Settings configured for during the meeting.")
$TRANSLATIONS_EN.Add("INFO_OUTSIDE_MEETING_WINDOW", "Outside of meeting window. Meeting started more than 2 hours ago.")
$TRANSLATIONS_EN.Add("INFO_CONFIG_LOADED", "Configuration loaded successfully.")
$TRANSLATIONS_EN.Add("INFO_NO_MEETING_SCHEDULED", "No meeting scheduled for today. Waiting for the next meeting day...")
$TRANSLATIONS_EN.Add("INFO_SETTING_SECURITY", "Configuring security setting: {0}")
$TRANSLATIONS_EN.Add("INFO_TOGGLE_FEED", "Toggling {0} feed")
$TRANSLATIONS_EN.Add("INFO_MUTE_ALL", "Muting all participants")
$TRANSLATIONS_EN.Add("INFO_UNMUTE_ALL", "Unmuting all participants")
$TRANSLATIONS_EN.Add("INFO_OPEN_PARTICIPANTS_PANEL", "Opening participants panel")
$TRANSLATIONS_EN.Add("INFO_GET_MORE_MENU", "Opening More menu")
$TRANSLATIONS_EN.Add("INFO_CLOSE_HOST_TOOLS", "Closing Host tools menu")

; Labels
$TRANSLATIONS_EN.Add("LABEL_MEETING_ID", "Zoom Meeting ID")
$TRANSLATIONS_EN.Add("LABEL_MIDWEEK_DAY", "Midweek Day")
$TRANSLATIONS_EN.Add("LABEL_MIDWEEK_TIME", "Midweek Time (HH:MM)")
$TRANSLATIONS_EN.Add("LABEL_WEEKEND_DAY", "Weekend Day")
$TRANSLATIONS_EN.Add("LABEL_WEEKEND_TIME", "Weekend Time (HH:MM)")

; Zoom interface labels
$TRANSLATIONS_EN.Add("LABEL_HOST_TOOLS", "Host tools")
$TRANSLATIONS_EN.Add("LABEL_MORE_MEETING_CONTROLS", "More meeting controls")
$TRANSLATIONS_EN.Add("LABEL_PARTICIPANT", "Participant")
$TRANSLATIONS_EN.Add("LABEL_MUTE_ALL", "Mute All")
$TRANSLATIONS_EN.Add("LABEL_YES", "Yes")
$TRANSLATIONS_EN.Add("LABEL_UNCHECKED_VALUE", "Unchecked")
$TRANSLATIONS_EN.Add("LABEL_CURRENTLY_UNMUTED_VALUE", "Currently unmuted")
$TRANSLATIONS_EN.Add("LABEL_UNMUTE_AUDIO_VALUE", "Unmute my audio")
$TRANSLATIONS_EN.Add("LABEL_STOP_VIDEO_VALUE", "Stop my video")
$TRANSLATIONS_EN.Add("LABEL_START_VIDEO_VALUE", "Start my video")
$TRANSLATIONS_EN.Add("LABEL_ZOOM_SECURITY_UNMUTE", "Unmute themselves")
$TRANSLATIONS_EN.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN", "Share screen")

; Help text
$TRANSLATIONS_EN.Add("LABEL_HOST_TOOLS_EXPLAIN", "Enter the text that appears on the Host Tools button in your Zoom interface. This is used to locate and click the button automatically.")
$TRANSLATIONS_EN.Add("LABEL_PARTICIPANT_EXPLAIN", "Enter the text that appears on the Participants button in your Zoom interface. This is used to locate and open the participants panel.")
$TRANSLATIONS_EN.Add("LABEL_MUTE_ALL_EXPLAIN", "Enter the text that appears on the Mute All button in your Zoom interface. This is used to mute all participants automatically.")
$TRANSLATIONS_EN.Add("LABEL_YES_EXPLAIN", "Enter the text that appears on confirmation buttons (e.g., ""Yes"", ""OK"") in your Zoom interface. This is used to confirm actions.")
$TRANSLATIONS_EN.Add("LABEL_UNCHECKED_VALUE_EXPLAIN", "Enter the text that appears when a setting is unchecked/disabled in your Zoom interface. This is used to detect when settings are disabled.")
$TRANSLATIONS_EN.Add("LABEL_CURRENTLY_UNMUTED_VALUE_EXPLAIN", "Enter the text that appears on the audio button when you are currently unmuted in your Zoom interface. This is used to detect audio state.")
$TRANSLATIONS_EN.Add("LABEL_UNMUTE_AUDIO_VALUE_EXPLAIN", "Enter the text that appears on the button to unmute your audio in your Zoom interface. This is used to unmute yourself.")
$TRANSLATIONS_EN.Add("LABEL_STOP_VIDEO_VALUE_EXPLAIN", "Enter the text that appears on the button to stop your video in your Zoom interface. This is used to stop your video feed.")
$TRANSLATIONS_EN.Add("LABEL_START_VIDEO_VALUE_EXPLAIN", "Enter the text that appears on the button to start your video in your Zoom interface. This is used to start your video feed.")
$TRANSLATIONS_EN.Add("LABEL_ZOOM_SECURITY_UNMUTE_EXPLAIN", "Enter the text that appears on the Unmute permission setting in Zoom Security settings. This controls whether participants can unmute themselves.")
$TRANSLATIONS_EN.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN_EXPLAIN", "Enter the text that appears on the Share screen permission setting in Zoom Security settings. This controls whether participants can share their screen.")

; Settings
$TRANSLATIONS_EN.Add("LABEL_SNAP_ZOOM_TO", "Snap Zoom window to")
$TRANSLATIONS_EN.Add("SNAP_DISABLED", "Disabled")
$TRANSLATIONS_EN.Add("SNAP_LEFT", "Left")
$TRANSLATIONS_EN.Add("SNAP_RIGHT", "Right")
$TRANSLATIONS_EN.Add("LABEL_KEYBOARD_SHORTCUT", "Post-meeting Keyboard Shortcut")
$TRANSLATIONS_EN.Add("LABEL_KEYBOARD_SHORTCUT_EXPLAIN", "Enter a keyboard shortcut that will apply post-meeting settings (e.g., Ctrl+Alt+Z). Use ^ for Ctrl, ! for Alt, + for Shift, # for Win, followed by a letter or number.")

; Error messages
$TRANSLATIONS_EN.Add("ERROR_GET_DESKTOP_ELEMENT_FAILED", "Failed to get desktop element.")
$TRANSLATIONS_EN.Add("ERROR_ZOOM_LAUNCH", "Error launching Zoom")
$TRANSLATIONS_EN.Add("ERROR_ZOOM_WINDOW_NOT_FOUND", "Zoom window not found")
$TRANSLATIONS_EN.Add("ERROR_MEETING_ID_NOT_CONFIGURED", "Meeting ID not configured.")
$TRANSLATIONS_EN.Add("ERROR_MEETING_ID_FORMAT", "Enter 9–11 digits (no spaces)")
$TRANSLATIONS_EN.Add("ERROR_TIME_FORMAT", "Use 24h time HH:MM")
$TRANSLATIONS_EN.Add("ERROR_KEYBOARD_SHORTCUT_FORMAT", "Use format like ^!z (Ctrl+Alt+Z). Must include at least one modifier (^ Ctrl, ! Alt, + Shift, # Win) followed by a letter or number.")
$TRANSLATIONS_EN.Add("ERROR_REQUIRED", "This field is required")
$TRANSLATIONS_EN.Add("ERROR_FIELDS_REQUIRED", "Please complete all required fields")
$TRANSLATIONS_EN.Add("ERROR_INVALID_ELEMENT_OBJECT", "Invalid element object.")
$TRANSLATIONS_EN.Add("ERROR_FAILED_CLICK_ELEMENT", "Failed to click element")
$TRANSLATIONS_EN.Add("ERROR_SETTING_NOT_FOUND", "Setting not found")
$TRANSLATIONS_EN.Add("ERROR_UNKNOWN_FEED_TYPE", "Unknown feed type")
$TRANSLATIONS_EN.Add("ERROR_ELEMENT_NOT_FOUND", "Element not found: {0}")
$TRANSLATIONS_EN.Add("ERROR_VARIOUS_ELEMENTS", "Various elements")

; Overlay messages
$TRANSLATIONS_EN.Add("PLEASE_WAIT", "Please wait...")
$TRANSLATIONS_EN.Add("POST_MEETING_HIT_KEY", "Are you ready to apply post-meeting settings? Press ENTER when the prayer is over to apply them, or ESC to cancel.")

; Section headers
$TRANSLATIONS_EN.Add("SECTION_MEETING_INFO", "Meeting Information")
$TRANSLATIONS_EN.Add("SECTION_ZOOM_LABELS", "Zoom Interface Labels")
$TRANSLATIONS_EN.Add("SECTION_GENERAL_SETTINGS", "General Settings")

; Day labels (1=Sunday .. 7=Saturday)
$TRANSLATIONS_EN.Add("DAY_1", "Sunday")
$TRANSLATIONS_EN.Add("DAY_2", "Monday")
$TRANSLATIONS_EN.Add("DAY_3", "Tuesday")
$TRANSLATIONS_EN.Add("DAY_4", "Wednesday")
$TRANSLATIONS_EN.Add("DAY_5", "Thursday")
$TRANSLATIONS_EN.Add("DAY_6", "Friday")
$TRANSLATIONS_EN.Add("DAY_7", "Saturday")
