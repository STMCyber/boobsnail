class Excel4MissingTranslationLevel():
    '''
    `Excel4MissingTranslationLevel` is used by `Excel4TranslatorConfig.missing_translation` property.
    '''

    # Raise exception Excel4TranslationMissingException if translation is missing
    EXCEPTION = 1
    # Write to the console that translation is missing
    LOG = 2
    # Do nothing
    NOTHING = 3

class Excel4TranslatorConfig(object):
    '''
    `Excel4TranslatorConfig` is used as global configuration class.
     It enables static and programmable change of behavior of `Excel4Translator` object.
    '''

    # Default language to which you want translate
    language = "en_US"

    # Directory containing translations files
    translations_directory = "excel4lib/lang/langs"

    # Extension of translations files
    translations_ext = ".json"

    # What to do if formula translation is missing
    missing_translation = Excel4MissingTranslationLevel.LOG
