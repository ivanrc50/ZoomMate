; ================================================================================================
; ZoomMate Russian Translation Data
; ================================================================================================

; Language metadata
Global $TRANSLATIONS_RU = ObjCreate("Scripting.Dictionary")
$TRANSLATIONS_RU.Add("LANGNAME", "Русский")

; Configuration GUI
$TRANSLATIONS_RU.Add("CONFIG_TITLE", "Конфигурация ZoomMate")
$TRANSLATIONS_RU.Add("BTN_SAVE", "Сохранить")
$TRANSLATIONS_RU.Add("BTN_QUIT", "Выйти из ZoomMate")
$TRANSLATIONS_RU.Add("LABEL_LANGUAGE", "Язык:")

; Status messages
$TRANSLATIONS_RU.Add("TOOLTIP_IDLE", "Бездействие")
$TRANSLATIONS_RU.Add("INFO_ZOOM_LAUNCHING", "Запуск Zoom...")
$TRANSLATIONS_RU.Add("INFO_ZOOM_LAUNCHED", "Собрание Zoom запущено")
$TRANSLATIONS_RU.Add("INFO_MEETING_STARTING_IN", "Собрание начинается через {0} минуту(ы).")
$TRANSLATIONS_RU.Add("INFO_MEETING_STARTED_AGO", "Собрание началось {0} минуту(ы) назад.")
$TRANSLATIONS_RU.Add("INFO_CONFIG_BEFORE_AFTER_START", "Настройка параметров до и после собраний...")
$TRANSLATIONS_RU.Add("INFO_CONFIG_BEFORE_AFTER_DONE", "Параметры настроены для до и после собраний.")
$TRANSLATIONS_RU.Add("INFO_MEETING_STARTING_SOON_CONFIG", "Собрание начинается скоро... Настройка параметров.")
$TRANSLATIONS_RU.Add("INFO_CONFIG_DURING_MEETING_DONE", "Параметры настроены для проведения собрания.")
$TRANSLATIONS_RU.Add("INFO_OUTSIDE_MEETING_WINDOW", "Вне окна собрания. Собрание началось более 2 часов назад.")
$TRANSLATIONS_RU.Add("INFO_CONFIG_LOADED", "Конфигурация загружена успешно.")
$TRANSLATIONS_RU.Add("INFO_NO_MEETING_SCHEDULED", "На сегодня собрание не запланировано. Ожидание следующего дня собрания...")

; Labels
$TRANSLATIONS_RU.Add("LABEL_MEETING_ID", "ID собрания Zoom")
$TRANSLATIONS_RU.Add("LABEL_MIDWEEK_DAY", "День недели")
$TRANSLATIONS_RU.Add("LABEL_MIDWEEK_TIME", "Время недели (ЧЧ:ММ)")
$TRANSLATIONS_RU.Add("LABEL_WEEKEND_DAY", "Выходной день")
$TRANSLATIONS_RU.Add("LABEL_WEEKEND_TIME", "Время выходного (ЧЧ:ММ)")

; Zoom interface labels
$TRANSLATIONS_RU.Add("LABEL_HOST_TOOLS", "Инструменты ведущего")
$TRANSLATIONS_RU.Add("LABEL_MORE_MEETING_CONTROLS", "Дополнительные элементы управления собранием")
$TRANSLATIONS_RU.Add("LABEL_PARTICIPANT", "Участник")
$TRANSLATIONS_RU.Add("LABEL_MUTE_ALL", "Отключить всех")
$TRANSLATIONS_RU.Add("LABEL_YES", "Да")
$TRANSLATIONS_RU.Add("LABEL_UNCHECKED_VALUE", "Не отмечено")
$TRANSLATIONS_RU.Add("LABEL_CURRENTLY_UNMUTED_VALUE", "Текущий ненастроенный")
$TRANSLATIONS_RU.Add("LABEL_UNMUTE_AUDIO_VALUE", "Включить мой звук")
$TRANSLATIONS_RU.Add("LABEL_STOP_VIDEO_VALUE", "Остановить мое видео")
$TRANSLATIONS_RU.Add("LABEL_START_VIDEO_VALUE", "Запустить мое видео")
$TRANSLATIONS_RU.Add("LABEL_ZOOM_SECURITY_UNMUTE", "Разрешить участникам включать микрофон")
$TRANSLATIONS_RU.Add("LABEL_ZOOM_SECURITY_SHARE_SCREEN", "Демонстрация экрана")

; Settings
$TRANSLATIONS_RU.Add("LABEL_SNAP_ZOOM_TO", "Прикрепить окно Zoom к")
$TRANSLATIONS_RU.Add("SNAP_DISABLED", "Отключено")
$TRANSLATIONS_RU.Add("SNAP_LEFT", "Слева")
$TRANSLATIONS_RU.Add("SNAP_RIGHT", "Справа")
$TRANSLATIONS_RU.Add("LABEL_KEYBOARD_SHORTCUT", "Горячая клавиша после собрания")
$TRANSLATIONS_RU.Add("LABEL_KEYBOARD_SHORTCUT_EXPLAIN", "Введите горячую клавишу, которая применит настройки после собрания (например, Ctrl+Alt+Z). Используйте ^ для Ctrl, ! для Alt, + для Shift, # для Win, за которой следует буква или цифра.")

; Error messages
$TRANSLATIONS_RU.Add("ERROR_GET_DESKTOP_ELEMENT_FAILED", "Не удалось получить элемент рабочего стола.")
$TRANSLATIONS_RU.Add("ERROR_ZOOM_LAUNCH", "Ошибка запуска Zoom")
$TRANSLATIONS_RU.Add("ERROR_ZOOM_WINDOW_NOT_FOUND", "Окно Zoom не найдено")
$TRANSLATIONS_RU.Add("ERROR_MEETING_ID_NOT_CONFIGURED", "ID собрания не настроен.")
$TRANSLATIONS_RU.Add("ERROR_MEETING_ID_FORMAT", "Введите 9–11 цифр (без пробелов)")
$TRANSLATIONS_RU.Add("ERROR_TIME_FORMAT", "Используйте формат 24ч ЧЧ:ММ")
$TRANSLATIONS_RU.Add("ERROR_KEYBOARD_SHORTCUT_FORMAT", "Используйте формат ^!z (Ctrl+Alt+Z). Должен включать хотя бы один модификатор (^ Ctrl, ! Alt, + Shift, # Win) за которым следует буква или цифра.")
$TRANSLATIONS_RU.Add("ERROR_REQUIRED", "Это поле обязательно")
$TRANSLATIONS_RU.Add("ERROR_FIELDS_REQUIRED", "Пожалуйста, заполните все обязательные поля")
$TRANSLATIONS_RU.Add("ERROR_INVALID_ELEMENT_OBJECT", "Недопустимый объект элемента.")
$TRANSLATIONS_RU.Add("ERROR_FAILED_CLICK_ELEMENT", "Не удалось нажать на элемент")
$TRANSLATIONS_RU.Add("ERROR_SETTING_NOT_FOUND", "Настройка не найдена")
$TRANSLATIONS_RU.Add("ERROR_UNKNOWN_FEED_TYPE", "Неизвестный тип потока")

; Overlay messages
$TRANSLATIONS_RU.Add("PLEASE_WAIT_TITLE", "Пожалуйста, подождите")
$TRANSLATIONS_RU.Add("PLEASE_WAIT_TEXT", "Пожалуйста, подождите...")
$TRANSLATIONS_RU.Add("POST_MEETING_HIT_KEY_TITLE", "Настройки после собрания")
$TRANSLATIONS_RU.Add("POST_MEETING_HIT_KEY_TEXT", "Готовы применить настройки после собрания? Нажмите ENTER, когда молитва закончится, чтобы применить их, или ESC для отмены.")

; Section headers
$TRANSLATIONS_RU.Add("SECTION_MEETING_INFO", "Информация о собрании")
$TRANSLATIONS_RU.Add("SECTION_ZOOM_LABELS", "Метки интерфейса Zoom")
$TRANSLATIONS_RU.Add("SECTION_GENERAL_SETTINGS", "Общие настройки")

; Day labels (1=Sunday .. 7=Saturday)
$TRANSLATIONS_RU.Add("DAY_1", "Воскресенье")
$TRANSLATIONS_RU.Add("DAY_2", "Понедельник")
$TRANSLATIONS_RU.Add("DAY_3", "Вторник")
$TRANSLATIONS_RU.Add("DAY_4", "Среда")
$TRANSLATIONS_RU.Add("DAY_5", "Четверг")
$TRANSLATIONS_RU.Add("DAY_6", "Пятница")
$TRANSLATIONS_RU.Add("DAY_7", "Суббота")

; Helper function to get Russian translations
Func _GetRussianTranslations()
    Return $TRANSLATIONS_RU
EndFunc
