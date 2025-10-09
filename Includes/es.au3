; ================================================================================================
; ZoomMate Spanish Translation Data
; ================================================================================================

; Language metadata
Global $TRANSLATIONS_ES = ObjCreate("Scripting.Dictionary")
$TRANSLATIONS_ES.Add("LANGNAME", "Español")

; Configuration GUI
$TRANSLATIONS_ES.Add("CONFIG_TITLE", "Configuración de ZoomMate")
$TRANSLATIONS_ES.Add("BTN_SAVE", "Guardar")
$TRANSLATIONS_ES.Add("BTN_QUIT", "Salir de ZoomMate")
$TRANSLATIONS_ES.Add("LABEL_LANGUAGE", "Idioma:")

; Status messages
$TRANSLATIONS_ES.Add("TOOLTIP_IDLE", "Inactivo")
$TRANSLATIONS_ES.Add("INFO_ZOOM_LAUNCHING", "Iniciando Zoom...")
$TRANSLATIONS_ES.Add("INFO_ZOOM_LAUNCHED", "Reunión de Zoom iniciada")
$TRANSLATIONS_ES.Add("INFO_MEETING_STARTING_IN", "La reunión comienza en {0} minuto(s).")
$TRANSLATIONS_ES.Add("INFO_MEETING_STARTED_AGO", "La reunión comenzó hace {0} minuto(s).")
$TRANSLATIONS_ES.Add("INFO_CONFIG_BEFORE_AFTER_START", "Configurando ajustes para antes y después de las reuniones...")
$TRANSLATIONS_ES.Add("INFO_CONFIG_BEFORE_AFTER_DONE", "Ajustes configurados para antes y después de las reuniones.")
$TRANSLATIONS_ES.Add("INFO_MEETING_STARTING_SOON_CONFIG", "La reunión comienza pronto... Configurando ajustes.")
$TRANSLATIONS_ES.Add("INFO_CONFIG_DURING_MEETING_DONE", "Ajustes configurados para durante la reunión.")
$TRANSLATIONS_ES.Add("INFO_OUTSIDE_MEETING_WINDOW", "Fuera de la ventana de reunión. La reunión comenzó hace más de 2 horas.")
$TRANSLATIONS_ES.Add("INFO_CONFIG_LOADED", "Configuración cargada exitosamente.")
$TRANSLATIONS_ES.Add("INFO_NO_MEETING_SCHEDULED", "No hay reunión programada para hoy. Esperando el próximo día de reunión...")

; Labels
$TRANSLATIONS_ES.Add("LABEL_MEETING_ID", "ID de reunión de Zoom")
$TRANSLATIONS_ES.Add("LABEL_MIDWEEK_DAY", "Día entre semana")
$TRANSLATIONS_ES.Add("LABEL_MIDWEEK_TIME", "Hora entre semana (HH:MM)")
$TRANSLATIONS_ES.Add("LABEL_WEEKEND_DAY", "Día de fin de semana")
$TRANSLATIONS_ES.Add("LABEL_WEEKEND_TIME", "Hora de fin de semana (HH:MM)")

; Zoom interface labels
$TRANSLATIONS_ES.Add("LABEL_HOST_TOOLS", "Herramientas del anfitrión")
$TRANSLATIONS_ES.Add("LABEL_MORE_MEETING_CONTROLS", "Más controles de reunión")
$TRANSLATIONS_ES.Add("LABEL_PARTICIPANT", "Participante")
$TRANSLATIONS_ES.Add("LABEL_MUTE_ALL", "Silenciar a todos")
$TRANSLATIONS_ES.Add("LABEL_YES", "Sí")
$TRANSLATIONS_ES.Add("LABEL_UNCHECKED_VALUE", "Sin marcar")
$TRANSLATIONS_ES.Add("LABEL_CURRENTLY_UNMUTED_VALUE", "Actualmente sin silenciar")
$TRANSLATIONS_ES.Add("LABEL_UNMUTE_AUDIO_VALUE", "Activar mi audio")
$TRANSLATIONS_ES.Add("LABEL_STOP_VIDEO_VALUE", "Detener mi video")
$TRANSLATIONS_ES.Add("LABEL_START_VIDEO_VALUE", "Iniciar mi video")
$TRANSLATIONS_ES.Add("LABEL_ZOOM_SECURITY_UNMUTE", "Permitir que se activen el micrófono")
$TRANSLATIONS_ES.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN", "Compartir pantalla")

; Help text
$TRANSLATIONS_ES.Add("LABEL_HOST_TOOLS_EXPLAIN", "Ingrese el texto que aparece en el botón Herramientas del anfitrión en su interfaz de Zoom. Esto se usa para localizar y hacer clic en el botón automáticamente.")
$TRANSLATIONS_ES.Add("LABEL_PARTICIPANT_EXPLAIN", "Ingrese el texto que aparece en el botón Participantes en su interfaz de Zoom. Esto se usa para localizar y abrir el panel de participantes.")
$TRANSLATIONS_ES.Add("LABEL_MUTE_ALL_EXPLAIN", "Ingrese el texto que aparece en el botón Silenciar a todos en su interfaz de Zoom. Esto se usa para silenciar a todos los participantes automáticamente.")
$TRANSLATIONS_ES.Add("LABEL_YES_EXPLAIN", "Ingrese el texto que aparece en los botones de confirmación (ej. ""Sí"", ""Aceptar"") en su interfaz de Zoom. Esto se usa para confirmar acciones.")
$TRANSLATIONS_ES.Add("LABEL_UNCHECKED_VALUE_EXPLAIN", "Ingrese el texto que aparece cuando una configuración está desmarcada/deshabilitada en su interfaz de Zoom. Esto se usa para detectar cuando las configuraciones están deshabilitadas.")
$TRANSLATIONS_ES.Add("LABEL_CURRENTLY_UNMUTED_VALUE_EXPLAIN", "Ingrese el texto que aparece en el botón de audio cuando actualmente no está silenciado en su interfaz de Zoom. Esto se usa para detectar el estado del audio.")
$TRANSLATIONS_ES.Add("LABEL_UNMUTE_AUDIO_VALUE_EXPLAIN", "Ingrese el texto que aparece en el botón para activar su audio en su interfaz de Zoom. Esto se usa para activar su propio audio.")
$TRANSLATIONS_ES.Add("LABEL_STOP_VIDEO_VALUE_EXPLAIN", "Ingrese el texto que aparece en el botón para detener su video en su interfaz de Zoom. Esto se usa para detener su transmisión de video.")
$TRANSLATIONS_ES.Add("LABEL_START_VIDEO_VALUE_EXPLAIN", "Ingrese el texto que aparece en el botón para iniciar su video en su interfaz de Zoom. Esto se usa para iniciar su transmisión de video.")
$TRANSLATIONS_ES.Add("LABEL_ZOOM_SECURITY_UNMUTE_EXPLAIN", "Ingrese el texto que aparece en la configuración de permiso para activar micrófono en la configuración de seguridad de Zoom. Esto controla si los participantes pueden activar su propio micrófono.")
$TRANSLATIONS_ES.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN_EXPLAIN", "Ingrese el texto que aparece en la configuración de permiso para compartir pantalla en la configuración de seguridad de Zoom. Esto controla si los participantes pueden compartir su pantalla.")

; Settings
$TRANSLATIONS_ES.Add("LABEL_SNAP_ZOOM_TO", "Ajustar ventana de Zoom a")
$TRANSLATIONS_ES.Add("SNAP_DISABLED", "Deshabilitado")
$TRANSLATIONS_ES.Add("SNAP_LEFT", "Izquierda")
$TRANSLATIONS_ES.Add("SNAP_RIGHT", "Derecha")
$TRANSLATIONS_ES.Add("LABEL_KEYBOARD_SHORTCUT", "Atajo de teclado posterior a la reunión")
$TRANSLATIONS_ES.Add("LABEL_KEYBOARD_SHORTCUT_EXPLAIN", "Ingrese un atajo de teclado que aplicará los ajustes posteriores a la reunión (ej. Ctrl+Alt+Z). Use ^ para Ctrl, ! para Alt, + para Shift, # para Win, seguido de una letra o número.")

; Error messages
$TRANSLATIONS_ES.Add("ERROR_GET_DESKTOP_ELEMENT_FAILED", "Error al obtener el elemento del escritorio.")
$TRANSLATIONS_ES.Add("ERROR_ZOOM_LAUNCH", "Error al iniciar Zoom")
$TRANSLATIONS_ES.Add("ERROR_ZOOM_WINDOW_NOT_FOUND", "Ventana de Zoom no encontrada")
$TRANSLATIONS_ES.Add("ERROR_MEETING_ID_NOT_CONFIGURED", "ID de reunión no configurado.")
$TRANSLATIONS_ES.Add("ERROR_MEETING_ID_FORMAT", "Ingrese 9-11 dígitos (sin espacios)")
$TRANSLATIONS_ES.Add("ERROR_TIME_FORMAT", "Use formato 24h HH:MM")
$TRANSLATIONS_ES.Add("ERROR_KEYBOARD_SHORTCUT_FORMAT", "Use formato como ^!z (Ctrl+Alt+Z). Debe incluir al menos un modificador (^ Ctrl, ! Alt, + Shift, # Win) seguido de una letra o número.")
$TRANSLATIONS_ES.Add("ERROR_REQUIRED", "Este campo es requerido")
$TRANSLATIONS_ES.Add("ERROR_FIELDS_REQUIRED", "Por favor complete todos los campos requeridos")
$TRANSLATIONS_ES.Add("ERROR_INVALID_ELEMENT_OBJECT", "Objeto de elemento inválido.")
$TRANSLATIONS_ES.Add("ERROR_FAILED_CLICK_ELEMENT", "Error al hacer clic en elemento")
$TRANSLATIONS_ES.Add("ERROR_SETTING_NOT_FOUND", "Configuración no encontrada")
$TRANSLATIONS_ES.Add("ERROR_UNKNOWN_FEED_TYPE", "Tipo de transmisión desconocido")

; Overlay messages
$TRANSLATIONS_ES.Add("PLEASE_WAIT_TITLE", "Espere por favor")
$TRANSLATIONS_ES.Add("PLEASE_WAIT_TEXT", "Espere por favor...")
$TRANSLATIONS_ES.Add("POST_MEETING_HIT_KEY_TITLE", "Ajustes posteriores a la reunión")
$TRANSLATIONS_ES.Add("POST_MEETING_HIT_KEY_TEXT", "¿Está listo para aplicar los ajustes posteriores a la reunión? Presione ENTER cuando termine la oración para aplicarlos, o ESC para cancelar.")

; Section headers
$TRANSLATIONS_ES.Add("SECTION_MEETING_INFO", "Información de la reunión")
$TRANSLATIONS_ES.Add("SECTION_ZOOM_LABELS", "Etiquetas de interfaz de Zoom")
$TRANSLATIONS_ES.Add("SECTION_GENERAL_SETTINGS", "Configuración general")

; Day labels (1=Sunday .. 7=Saturday)
$TRANSLATIONS_ES.Add("DAY_1", "Domingo")
$TRANSLATIONS_ES.Add("DAY_2", "Lunes")
$TRANSLATIONS_ES.Add("DAY_3", "Martes")
$TRANSLATIONS_ES.Add("DAY_4", "Miércoles")
$TRANSLATIONS_ES.Add("DAY_5", "Jueves")
$TRANSLATIONS_ES.Add("DAY_6", "Viernes")
$TRANSLATIONS_ES.Add("DAY_7", "Sábado")
