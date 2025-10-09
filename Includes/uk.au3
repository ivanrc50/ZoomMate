; ================================================================================================
; ZoomMate Ukrainian Translation Data
; ================================================================================================

; Language metadata
Global $TRANSLATIONS_UK = ObjCreate("Scripting.Dictionary")
$TRANSLATIONS_UK.Add("LANGNAME", "Українська")

; Configuration GUI
$TRANSLATIONS_UK.Add("CONFIG_TITLE", "Конфігурація ZoomMate")
$TRANSLATIONS_UK.Add("BTN_SAVE", "Зберегти")
$TRANSLATIONS_UK.Add("BTN_QUIT", "Вийти з ZoomMate")
$TRANSLATIONS_UK.Add("LABEL_LANGUAGE", "Мова:")

; Status messages
$TRANSLATIONS_UK.Add("TOOLTIP_IDLE", "Бездіяльність")
$TRANSLATIONS_UK.Add("INFO_ZOOM_LAUNCHING", "Запуск Zoom...")
$TRANSLATIONS_UK.Add("INFO_ZOOM_LAUNCHED", "Збори Zoom запущено")
$TRANSLATIONS_UK.Add("INFO_MEETING_STARTING_IN", "Збори починаються за {0} хвилину(и).")
$TRANSLATIONS_UK.Add("INFO_MEETING_STARTED_AGO", "Збори почалися {0} хвилину(и) тому.")
$TRANSLATIONS_UK.Add("INFO_CONFIG_BEFORE_AFTER_START", "Налаштування параметрів до і після зборів...")
$TRANSLATIONS_UK.Add("INFO_CONFIG_BEFORE_AFTER_DONE", "Параметри налаштовано для до і після зборів.")
$TRANSLATIONS_UK.Add("INFO_MEETING_STARTING_SOON_CONFIG", "Збори починаються скоро... Налаштування параметрів.")
$TRANSLATIONS_UK.Add("INFO_CONFIG_DURING_MEETING_DONE", "Параметри налаштовано для проведення зборів.")
$TRANSLATIONS_UK.Add("INFO_OUTSIDE_MEETING_WINDOW", "Поза вікном зборів. Збори почалися більше 2 годин тому.")
$TRANSLATIONS_UK.Add("INFO_CONFIG_LOADED", "Конфігурацію завантажено успішно.")
$TRANSLATIONS_UK.Add("INFO_NO_MEETING_SCHEDULED", "На сьогодні збори не заплановано. Очікування наступного дня зборів...")

; Labels
$TRANSLATIONS_UK.Add("LABEL_MEETING_ID", "ID зборів Zoom")
$TRANSLATIONS_UK.Add("LABEL_MIDWEEK_DAY", "День тижня")
$TRANSLATIONS_UK.Add("LABEL_MIDWEEK_TIME", "Час тижня (ГГ:ХХ)")
$TRANSLATIONS_UK.Add("LABEL_WEEKEND_DAY", "Вихідний день")
$TRANSLATIONS_UK.Add("LABEL_WEEKEND_TIME", "Час вихідного (ГГ:ХХ)")

; Zoom interface labels
$TRANSLATIONS_UK.Add("LABEL_HOST_TOOLS", "Інструменти ведучого")
$TRANSLATIONS_UK.Add("LABEL_MORE_MEETING_CONTROLS", "Додаткові елементи керування зборами")
$TRANSLATIONS_UK.Add("LABEL_PARTICIPANT", "Учасник")
$TRANSLATIONS_UK.Add("LABEL_MUTE_ALL", "Вимкнути всіх")
$TRANSLATIONS_UK.Add("LABEL_YES", "Так")
$TRANSLATIONS_UK.Add("LABEL_UNCHECKED_VALUE", "Не позначено")
$TRANSLATIONS_UK.Add("LABEL_CURRENTLY_UNMUTED_VALUE", "Поточно невимкнено")
$TRANSLATIONS_UK.Add("LABEL_UNMUTE_AUDIO_VALUE", "Увімкнути мій звук")
$TRANSLATIONS_UK.Add("LABEL_STOP_VIDEO_VALUE", "Зупинити моє відео")
$TRANSLATIONS_UK.Add("LABEL_START_VIDEO_VALUE", "Запустити моє відео")
$TRANSLATIONS_UK.Add("LABEL_ZOOM_SECURITY_UNMUTE", "Дозволити учасникам вмикати мікрофон")
$TRANSLATIONS_UK.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN", "Демонстрація екрана")

; Settings
$TRANSLATIONS_UK.Add("LABEL_SNAP_ZOOM_TO", "Прикріпити вікно Zoom до")
$TRANSLATIONS_UK.Add("SNAP_DISABLED", "Вимкнено")
$TRANSLATIONS_UK.Add("SNAP_LEFT", "Зліва")
$TRANSLATIONS_UK.Add("SNAP_RIGHT", "Справа")
$TRANSLATIONS_UK.Add("LABEL_KEYBOARD_SHORTCUT", "Гаряча клавіша після зборів")
$TRANSLATIONS_UK.Add("LABEL_KEYBOARD_SHORTCUT_EXPLAIN", "Введіть гарячу клавішу, яка застосує налаштування після зборів (наприклад, Ctrl+Alt+Z). Використовуйте ^ для Ctrl, ! для Alt, + для Shift, # для Win, за якою слідує літера або цифра.")

; Error messages
$TRANSLATIONS_UK.Add("ERROR_GET_DESKTOP_ELEMENT_FAILED", "Не вдалося отримати елемент робочого столу.")
$TRANSLATIONS_UK.Add("ERROR_ZOOM_LAUNCH", "Помилка запуску Zoom")
$TRANSLATIONS_UK.Add("ERROR_ZOOM_WINDOW_NOT_FOUND", "Вікно Zoom не знайдено")
$TRANSLATIONS_UK.Add("ERROR_MEETING_ID_NOT_CONFIGURED", "ID зборів не налаштовано.")
$TRANSLATIONS_UK.Add("ERROR_MEETING_ID_FORMAT", "Введіть 9–11 цифр (без пробілів)")
$TRANSLATIONS_UK.Add("ERROR_TIME_FORMAT", "Використовуйте формат 24г ГГ:ХХ")
$TRANSLATIONS_UK.Add("ERROR_KEYBOARD_SHORTCUT_FORMAT", "Використовуйте формат ^!z (Ctrl+Alt+Z). Має включати принаймні один модифікатор (^ Ctrl, ! Alt, + Shift, # Win) за яким слідує літера або цифра.")
$TRANSLATIONS_UK.Add("ERROR_REQUIRED", "Це поле обов'язкове")
$TRANSLATIONS_UK.Add("ERROR_FIELDS_REQUIRED", "Будь ласка, заповніть усі обов'язкові поля")
$TRANSLATIONS_UK.Add("ERROR_INVALID_ELEMENT_OBJECT", "Недійсний об'єкт елемента.")
$TRANSLATIONS_UK.Add("ERROR_FAILED_CLICK_ELEMENT", "Не вдалося натиснути на елемент")
$TRANSLATIONS_UK.Add("ERROR_SETTING_NOT_FOUND", "Налаштування не знайдено")
$TRANSLATIONS_UK.Add("ERROR_UNKNOWN_FEED_TYPE", "Невідомий тип потоку")

; Overlay messages
$TRANSLATIONS_UK.Add("PLEASE_WAIT_TITLE", "Будь ласка, зачекайте")
$TRANSLATIONS_UK.Add("PLEASE_WAIT_TEXT", "Будь ласка, зачекайте...")
$TRANSLATIONS_UK.Add("POST_MEETING_HIT_KEY_TITLE", "Налаштування після зборів")
$TRANSLATIONS_UK.Add("POST_MEETING_HIT_KEY_TEXT", "Готові застосувати налаштування після зборів? Натисніть ENTER, коли молитва закінчиться, щоб застосувати їх, або ESC для скасування.")

; Section headers
$TRANSLATIONS_UK.Add("SECTION_MEETING_INFO", "Інформація про збори")
$TRANSLATIONS_UK.Add("SECTION_ZOOM_LABELS", "Мітки інтерфейсу Zoom")
$TRANSLATIONS_UK.Add("SECTION_GENERAL_SETTINGS", "Загальні налаштування")

; Day labels (1=Sunday .. 7=Saturday)
$TRANSLATIONS_UK.Add("DAY_1", "Неділя")
$TRANSLATIONS_UK.Add("DAY_2", "Понеділок")
$TRANSLATIONS_UK.Add("DAY_3", "Вівторок")
$TRANSLATIONS_UK.Add("DAY_4", "Середа")
$TRANSLATIONS_UK.Add("DAY_5", "Четвер")
$TRANSLATIONS_UK.Add("DAY_6", "П'ятниця")
$TRANSLATIONS_UK.Add("DAY_7", "Субота")

; Helper function to get Ukrainian translations
Func _GetUkrainianTranslations()
    Return $TRANSLATIONS_UK
EndFunc
