from excel4lib.utils import *
from excel4lib.sheet.cell import *
from excel4lib.lang import *
from .excel4_name import *
class Excel4Instruction(Cell):
    '''
    Represents Excel4 instruction (could be formula, empty cell, variable, obfuscated formula) which should be placed
    at specified address.

    Excel4Instruction is responsible for:
    - returning the cell address;
    - translating cell address;
    - generating tag;
    - storing information about language.
    '''

    def __init__(self, x, y):
        Cell.__init__(self, x, y)
        # Characters used to indicate row and column
        self.row_character = ""
        self.col_character = ""
        # Save current language
        self.language = Excel4Translator.language
        # Generate random tag
        self.tag = random_tag(random_string(5))
        self.start_cell = None
        # Change reference style to A1
        if not Excel4Config.rc_reference_style:
            self.reference_style = CellReferenceStyle.A1_STYLE
        self.translate_address()

        # If False then obfuscation of this object is disabled
        # This flag has the highest priority.
        # The object will not be obfuscated even if obfuscation is enabled in the options
        self._obfuscate = True
        self._obfuscate_formula = True
        self._spread = True

#    def get_start_address(self):
#        if not self.start_cell:
#            return self.get_address()
#        return self.start_cell.get_address()

    def set_language(self, lang):
        '''
        Sets `language` to `lang` and translates.

        :param lang: name of the language
        '''
        self.language = lang
        self.translate_address()

    def get_reference(self, lang=None):
        '''
        Returns instruction address.
        Address is translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :param lang: language to which instruction should be translated

        :return: string representing address
        '''
        t_r = self.row_character
        t_c = self.col_character
        if lang and (lang != self.language):
            t_r = Excel4Translator.get_row_character(lang)
            t_c = Excel4Translator.get_col_character(lang)
        r = self.get_cell_address(t_r, t_c)
        return r

    def translate_address(self, lang=None):
        '''
        Translates characters indicating row and column.

        :param lang: language in which address should be returned. If None then lang is equal to self.language
        '''
        if not lang:
            lang = self.language

        self.row_character = Excel4Translator.get_row_character(lang)
        self.col_character = Excel4Translator.get_col_character(lang)

    def revert_address_translation(self):
        '''
        Reverts translation of characters indicating row and column to native language
        '''
        self.translate_address(Excel4Translator.native_language)


    def get_str(self, lang=None):
        '''
        Returns Excel4 instruction as string

        :param lang: language to which instruction should be translated

        :return: string representing address
        '''
        return self.get_reference(lang)

    def __str__(self):
        '''
        Returns Excel4 instruction address

        :return:string representing address
        '''
        return self.get_str(self.language)