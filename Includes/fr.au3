; ================================================================================================
; ZoomMate French Translation Data
; ================================================================================================

; Language metadata
Global $TRANSLATIONS_FR = ObjCreate("Scripting.Dictionary")
$TRANSLATIONS_FR.Add("LANGNAME", "Français")

; Configuration GUI
$TRANSLATIONS_FR.Add("CONFIG_TITLE", "Configuration ZoomMate")
$TRANSLATIONS_FR.Add("BTN_SAVE", "Enregistrer")
$TRANSLATIONS_FR.Add("BTN_QUIT", "Quitter ZoomMate")
$TRANSLATIONS_FR.Add("LABEL_LANGUAGE", "Langue :")

; Status messages
$TRANSLATIONS_FR.Add("TOOLTIP_IDLE", "Inactif")
$TRANSLATIONS_FR.Add("INFO_ZOOM_LAUNCHING", "Lancement de Zoom...")
$TRANSLATIONS_FR.Add("INFO_ZOOM_LAUNCHED", "Réunion Zoom lancée")
$TRANSLATIONS_FR.Add("INFO_MEETING_STARTING_IN", "La réunion commence dans {0} minute(s).")
$TRANSLATIONS_FR.Add("INFO_MEETING_STARTED_AGO", "La réunion a commencé il y a {0} minute(s).")
$TRANSLATIONS_FR.Add("INFO_CONFIG_BEFORE_AFTER_START", "Configuration des paramètres avant et après les réunions...")
$TRANSLATIONS_FR.Add("INFO_CONFIG_BEFORE_AFTER_DONE", "Paramètres configurés pour avant et après les réunions.")
$TRANSLATIONS_FR.Add("INFO_MEETING_STARTING_SOON_CONFIG", "La réunion commence bientôt... Configuration des paramètres.")
$TRANSLATIONS_FR.Add("INFO_CONFIG_DURING_MEETING_DONE", "Paramètres configurés pour pendant la réunion.")
$TRANSLATIONS_FR.Add("INFO_OUTSIDE_MEETING_WINDOW", "En dehors de la fenêtre de réunion. La réunion a commencé il y a plus de 2 heures.")
$TRANSLATIONS_FR.Add("INFO_CONFIG_LOADED", "Configuration chargée avec succès.")
$TRANSLATIONS_FR.Add("INFO_NO_MEETING_SCHEDULED", "Aucune réunion prévue aujourd'hui. En attente du prochain jour de réunion...")

; Labels
$TRANSLATIONS_FR.Add("LABEL_MEETING_ID", "ID de réunion Zoom")
$TRANSLATIONS_FR.Add("LABEL_MIDWEEK_DAY", "Jour de semaine")
$TRANSLATIONS_FR.Add("LABEL_MIDWEEK_TIME", "Heure de semaine (HH:MM)")
$TRANSLATIONS_FR.Add("LABEL_WEEKEND_DAY", "Jour de week-end")
$TRANSLATIONS_FR.Add("LABEL_WEEKEND_TIME", "Heure de week-end (HH:MM)")

; Zoom interface labels
$TRANSLATIONS_FR.Add("LABEL_HOST_TOOLS", "Outils de l'hôte")
$TRANSLATIONS_FR.Add("LABEL_MORE_MEETING_CONTROLS", "Plus de contrôles de réunion")
$TRANSLATIONS_FR.Add("LABEL_PARTICIPANT", "Participant")
$TRANSLATIONS_FR.Add("LABEL_MUTE_ALL", "Tout mettre en sourdine")
$TRANSLATIONS_FR.Add("LABEL_YES", "Oui")
$TRANSLATIONS_FR.Add("LABEL_UNCHECKED_VALUE", "Non coché")
$TRANSLATIONS_FR.Add("LABEL_CURRENTLY_UNMUTED_VALUE", "Actuellement non muet")
$TRANSLATIONS_FR.Add("LABEL_UNMUTE_AUDIO_VALUE", "Activer mon audio")
$TRANSLATIONS_FR.Add("LABEL_STOP_VIDEO_VALUE", "Arrêter ma vidéo")
$TRANSLATIONS_FR.Add("LABEL_START_VIDEO_VALUE", "Démarrer ma vidéo")
$TRANSLATIONS_FR.Add("LABEL_ZOOM_SECURITY_UNMUTE", "Autoriser les participants à se réactiver")
$TRANSLATIONS_FR.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN", "Partager l'écran")

; Settings
$TRANSLATIONS_FR.Add("LABEL_SNAP_ZOOM_TO", "Ajuster la fenêtre Zoom à")
$TRANSLATIONS_FR.Add("SNAP_DISABLED", "Désactivé")
$TRANSLATIONS_FR.Add("SNAP_LEFT", "Gauche")
$TRANSLATIONS_FR.Add("SNAP_RIGHT", "Droite")
$TRANSLATIONS_FR.Add("LABEL_KEYBOARD_SHORTCUT", "Raccourci clavier après réunion")
$TRANSLATIONS_FR.Add("LABEL_KEYBOARD_SHORTCUT_EXPLAIN", "Saisissez un raccourci clavier qui appliquera les paramètres après la réunion (ex. Ctrl+Alt+Z). Utilisez ^ pour Ctrl, ! pour Alt, + pour Shift, # pour Win, suivi d'une lettre ou d'un chiffre.")

; Error messages
$TRANSLATIONS_FR.Add("ERROR_GET_DESKTOP_ELEMENT_FAILED", "Échec de l'obtention de l'élément du bureau.")
$TRANSLATIONS_FR.Add("ERROR_ZOOM_LAUNCH", "Erreur lors du lancement de Zoom")
$TRANSLATIONS_FR.Add("ERROR_ZOOM_WINDOW_NOT_FOUND", "Fenêtre Zoom introuvable")
$TRANSLATIONS_FR.Add("ERROR_MEETING_ID_NOT_CONFIGURED", "ID de réunion non configuré.")
$TRANSLATIONS_FR.Add("ERROR_MEETING_ID_FORMAT", "Saisissez 9 à 11 chiffres (sans espaces)")
$TRANSLATIONS_FR.Add("ERROR_TIME_FORMAT", "Utilisez le format 24h HH:MM")
$TRANSLATIONS_FR.Add("ERROR_KEYBOARD_SHORTCUT_FORMAT", "Utilisez le format ^!z (Ctrl+Alt+Z). Doit inclure au moins un modificateur (^ Ctrl, ! Alt, + Shift, # Win) suivi d'une lettre ou d'un chiffre.")
$TRANSLATIONS_FR.Add("ERROR_REQUIRED", "Ce champ est requis")
$TRANSLATIONS_FR.Add("ERROR_FIELDS_REQUIRED", "Veuillez remplir tous les champs requis")
$TRANSLATIONS_FR.Add("ERROR_INVALID_ELEMENT_OBJECT", "Objet élément invalide.")
$TRANSLATIONS_FR.Add("ERROR_FAILED_CLICK_ELEMENT", "Échec du clic sur l'élément")
$TRANSLATIONS_FR.Add("ERROR_SETTING_NOT_FOUND", "Paramètre introuvable")
$TRANSLATIONS_FR.Add("ERROR_UNKNOWN_FEED_TYPE", "Type de flux inconnu")

; Overlay messages
$TRANSLATIONS_FR.Add("PLEASE_WAIT_TITLE", "Veuillez patienter")
$TRANSLATIONS_FR.Add("PLEASE_WAIT_TEXT", "Veuillez patienter...")
$TRANSLATIONS_FR.Add("POST_MEETING_HIT_KEY_TITLE", "Paramètres après réunion")
$TRANSLATIONS_FR.Add("POST_MEETING_HIT_KEY_TEXT", "Êtes-vous prêt à appliquer les paramètres après la réunion ? Appuyez sur ENTRÉE quand la prière est terminée pour les appliquer, ou ÉCHAP pour annuler.")

; Section headers
$TRANSLATIONS_FR.Add("SECTION_MEETING_INFO", "Informations sur la réunion")
$TRANSLATIONS_FR.Add("SECTION_ZOOM_LABELS", "Étiquettes d'interface Zoom")
$TRANSLATIONS_FR.Add("SECTION_GENERAL_SETTINGS", "Paramètres généraux")

; Day labels (1=Sunday .. 7=Saturday)
$TRANSLATIONS_FR.Add("DAY_1", "Dimanche")
$TRANSLATIONS_FR.Add("DAY_2", "Lundi")
$TRANSLATIONS_FR.Add("DAY_3", "Mardi")
$TRANSLATIONS_FR.Add("DAY_4", "Mercredi")
$TRANSLATIONS_FR.Add("DAY_5", "Jeudi")
$TRANSLATIONS_FR.Add("DAY_6", "Vendredi")
$TRANSLATIONS_FR.Add("DAY_7", "Samedi")

; Helper function to get French translations
Func _GetFrenchTranslations()
    Return $TRANSLATIONS_FR
EndFunc
