from excel4lib.config import *
from excel4lib.config.excel4_translator_config import Excel4MissingTranslationLevel
from excel4lib.utils import *
from excel4lib.exception import *


class Excel4Translator(object):
    '''
    `Excel4Translator` class allows to translate english formulas to another language.

    `Excel4Translator` stores translation files in the langs directory in .json format.
    Translation files have the following format:
    ```
    {
      "arguments_separator": ",",
      "name": "LANG_NAME",
      "row_character": "ROW_CHARACTER",
      "col_character": "COL_CHARACTER",
      "translation": {
            "ENG_FORMULA":"TRANSLATION_FORMULA",
            (...)
        }
    }
    ```
    - `arguments_separator` -  stores character used to separate formula arguments;
    - `name` -  stores the name of language. It should be the same as the file name, with no extension for example, pl_PL (then file name is pl_pl.json);
    - `row_character` - stores character used to translate ROW character in RC_STYLE;
    - `col_character` - stores character used to translate COLUMN character when RC_STYLE is used;
    - `translation` - stores formulas translations in form KEY:VALUE where KEY is formula in english and VALUE
    is translation of this formula to corresponding language
    '''

    # Reference to configuration
    config = Excel4Config.translator

    # Current language - the language into which the text is to be translated
    language = config.language


    # Language from which translation is done
    # By default we use formulas in English
    # If you want to translate for example from Polish to English, then change Excel4Translator.native_language to pl_PL
    # and set Excel4Translator.language variable to en_US. Then create file en_US.json as translations file.
    # If Excel4Translator.language is equal to Excel4Translator.native_language then translation is not done
    native_language = "en_US"

    # Current language translations
    translations = {native_language:{}}

    # Default arguments separator. Returned when arguments_separator key is not defined in translations
    arguments_separator = ","

    # Default characters for rows and cols.
    row_character = "R"
    col_character = "C"

    @staticmethod
    def init():
        '''
        Initializes translator and loads `Excel4Translator.language` translation into memory.
        '''
        Excel4Translator.load_translations()

    @staticmethod
    def check_translations():
        '''
        Checks if translations have required keys. If not then `Excel4RequiredKeyMissingException` is raised.
        '''

        # Do not check if current language is equal to native
        if Excel4Translator.is_native():
            return

        req = ["translation"]
        translations_path = join_path(Excel4Translator.config.translations_directory,
                                      Excel4Translator.language + Excel4Translator.config.translations_ext)
        for k in req:
            if k not in Excel4Translator.translations[Excel4Translator.language]:
                raise Excel4RequiredKeyMissingException(
                    "{} key is missing in translations {}".format(k, translations_path))

    @staticmethod
    def load_translations(lang=None):
        '''
        Loads translation defined in `lang` into memory. If `lang` is None then `Excel4Translator.language` is loaded.
        If translation file does not exist or could not be found then `Excel4PathNotExistException` is raiesd.
        '''
        # Do not load if current language is equal to native
        if (not lang) and Excel4Translator.is_native():
            return

        if not lang:
            lang = Excel4Translator.language

        if lang in Excel4Translator.translations:
            return

        translations_path = join_path(Excel4Translator.config.translations_directory,
                                      lang + Excel4Translator.config.translations_ext)

        # Check if file with translations exists
        if not is_path(translations_path):
            raise Excel4PathNotExistException("File with translations {} does not exist".format(translations_path))
        Excel4Translator.translations[lang] = load_json_file(translations_path)

        # Check if translations have all required keys
        Excel4Translator.check_translations()

    @staticmethod
    def set_language(lang):
        '''
        Sets current language (`Excel4Translator.langauge`) to `lang` and loads translation.
        :param lang: name of the language
        '''

        # Save current language
        temp = Excel4Translator.language
        Excel4Translator.language = lang
        try:
            Excel4Translator.load_translations()
        except Exception as ex:
            # Restore language
            Excel4Translator.language = temp
            raise ex

    @staticmethod
    def is_native():
        '''
        Checks if  `Excel4Translator.language` is equal to `Excel4Translator.native_language`

        :return: True if yes and False if not
        '''
        return Excel4Translator.language == Excel4Translator.native_language

    @staticmethod
    def translate(formula, lang = None):
        '''
        Translates formula to `lang`. If `lang` is None then current language `Excel4Translator.language` is used.

        :param formula: name of formula to translate

        :param lang: name of the language

        :return: string translated formula
        '''
        lang_b = None
        # Init translations
        if not Excel4Translator.translations:
            Excel4Translator.init()
        # If formula is empty or it contains spaces then do not translate
        if (not formula) or (" " in formula):
            return formula
        if lang and (lang != Excel4Translator.language):
            lang_b = Excel4Translator.language
            Excel4Translator.set_language(lang)

        # Do not translate if current language is equal to native
        if Excel4Translator.is_native():
            return formula

        if not Excel4Translator.get_value("translation"):
            return

        if formula not in Excel4Translator.translations[Excel4Translator.language]["translation"]:

            # Raise exception if translation is missing
            if Excel4Translator.config.missing_translation == Excel4MissingTranslationLevel.EXCEPTION:
                translations_path = join_path(Excel4Translator.config.translations_directory,
                                              Excel4Translator.language + Excel4Translator.config.translations_ext)
                raise Excel4TranslationMissingException(
                    "Translation of {} formula is missing in translations {} file".format(formula, translations_path))
            # Print if translation is missing
            elif Excel4Translator.config.missing_translation == Excel4MissingTranslationLevel.LOG:
                translations_path = join_path(Excel4Translator.config.translations_directory,
                                              Excel4Translator.language + Excel4Translator.config.translations_ext)
                print("[!] Translation of {} formula is missing in translations {} file".format(formula, translations_path))

            return formula
        translation_f = Excel4Translator.translations[Excel4Translator.language]["translation"][formula]
        if lang_b:
            Excel4Translator.set_language(lang_b)
        return translation_f
    @staticmethod
    def t(formula, lang=None):
        '''
        Translates formula to `lang`. If `lang` is None then current language `Excel4Translator.language` is used.

        :param formula: name of formula to translate

        :param lang: name of the language

        :return: string translated formula
        '''
        return Excel4Translator.translate(formula, lang)

    @staticmethod
    def translate_address(address):
        '''
        Translates cell address

        :param address: address of cell to translate in RC_STYLE reference style

        :return: string translated address
        '''
        # Init translations
        if not Excel4Translator.translations:
            Excel4Translator.init()

        # Do not translate if current language is equal to native
        if Excel4Translator.is_native():
            return address

        # Do not translate if reference style is set to A1
        if not Excel4Config.rc_reference_style:
            return address

        return address.replace(Excel4Translator.row_character, Excel4Translator.get_row_character()).replace(Excel4Translator.col_character, Excel4Translator.get_col_character())

    @staticmethod
    def t_a(address):
        '''
        Translates cell address

        :param address: address of cell to translate in RC_STYLE reference style

        :return: string translated address
        '''
        return Excel4Translator.translate_address(address)

    @staticmethod
    def get_value(key_name):
        '''
        Returns value stored under `key_name` from  `Excel4Translator.translations`.
        If key does not exist then `Excel4RequiredKeyMissingException` is raised.

        :param key_name:

        :return: value stored under `key_name` in `Excel4Translator.translations` object
        '''
        if key_name not in Excel4Translator.translations[Excel4Translator.language]:
            translations_path = join_path(Excel4Translator.config.translations_directory,
                                          Excel4Translator.language + Excel4Translator.config.translations_ext)
            raise Excel4RequiredKeyMissingException(
                "{} key is missing in translations {}".format(key_name, translations_path))
        return Excel4Translator.translations[Excel4Translator.language][key_name]

    @staticmethod
    def get_arguments_separator(lang=None):
        '''
        Returns arguments separator for `lang`. If `lang` is None then current lanauge is used (`Excel4Translator.language`).

        :param lang: name of the language
        '''
        if (not lang) and Excel4Translator.is_native():
            return Excel4Translator.arguments_separator

        if not lang:
            lang = Excel4Translator.language

        if lang not in Excel4Translator.translations:
            Excel4Translator.load_translations(lang)



        return Excel4Translator.translations[lang].get("arguments_separator", Excel4Translator.arguments_separator)

    @staticmethod
    def get_row_character(lang=None):
        '''
        Returns row character for `lang`. If `lang` is None then current lanauge is used (`Excel4Translator.language`).

        :param lang: name of the language
        '''
        if (not lang) and Excel4Translator.is_native():
            return Excel4Translator.row_character

        if not lang:
            lang = Excel4Translator.language

        if lang not in Excel4Translator.translations:
            Excel4Translator.load_translations(lang)

        return Excel4Translator.translations[lang].get("row_character", Excel4Translator.row_character)

    @staticmethod
    def get_col_character(lang=None):
        '''
        Returns column character for `lang`. If `lang` is None then current lanauge is used (`Excel4Translator.language`).

        :param lang: name of the language
        '''
        if (not lang) and Excel4Translator.is_native():
            return Excel4Translator.col_character
        if not lang:
            lang = Excel4Translator.language

        if lang not in Excel4Translator.translations:
            Excel4Translator.load_translations(lang)

        return Excel4Translator.translations[lang].get("col_character", Excel4Translator.col_character)

    @staticmethod
    def get_languages():
        '''
        Returns list of available languages.
        '''
        translations_path = Excel4Translator.config.translations_directory
        langs = []
        for l in os.listdir(translations_path):
            if (Excel4Translator.config.translations_ext == l.lower().split(".")[-1]) or (Excel4Translator.config.translations_ext == "."+l.lower().split(".")[-1]):
                langs.append(".".join(l.split(".")[:-1]))
        return langs