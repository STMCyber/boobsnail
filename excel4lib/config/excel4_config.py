from excel4lib.config.excel4_obfuscator_config import *
from excel4lib.config.excel4_translator_config import *


class Excel4Config(object):
    '''
    `Excel4Config` is used as global configuration class.
     It enables static and programmable change of behavior of `Excel4Obfuscator`, `Excel4Translator` and `Excel4Macro` objects.
    '''

    # CSV Separator used during dumping macro to CSV file
    csv_separator = ";"

    # Enable R1C1 reference style
    rc_reference_style = True

    # Obfsucator configuration
    obfuscator = Excel4ObfuscatorConfig

    # Translator configuration
    translator = Excel4TranslatorConfig