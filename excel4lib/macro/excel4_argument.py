from excel4lib.utils import *
from excel4lib.macro.excel4_instruction import *
from excel4lib.lang import *

class Excel4Argument(object):
    '''
    `Excel4Argument` is a base class for Excel4 formulas that should not have address (formulas that are not placed in cell as a call but for example as formula argument)

    `Excel4Argument` class saves information about current language from `Excel4Translator.language` property.
    Language of argument could be changed by calling `Excel4Argument.set_language` function.

    Obufscation of `Excel4Argument` object could be disabled by setting `Excel4Argument._obfuscate` property.
    '''
    def __init__(self):
        # If False then obfuscation of this object is disabled
        # This flag has the highest priority.
        # The object will not be obfuscated even if obfuscation is enabled in the options
        self._obfuscate = True
        self._obfuscate_formula = True
        # Save current language
        self.language = Excel4Translator.language
        # Arguments separator
        self.args_sep = ""

    def get_length(self):
        '''
        Returns length of string representing this object.

        :return: int length of string representing this object.
        '''
        return len(str(self))

    def _get_argument(self, lang=None):
        return ""

    def get_str(self, lang=None):
        return self._get_argument(lang)

    def __str__(self):
        return self.get_str(self.language)

    def __add__(self, other):
        return self.get_str(self.language) + other

    def __radd__(self, other):
        return other + self.get_str(self.language)

    def _escape_value(self, val):
        return val.replace('"', '""')

    def translate(self):
        '''
        Translates formula
        '''
        self.args_sep = Excel4Translator.get_arguments_separator(self.language)

    def revert_translation(self):
        '''
        Reverts translation of formula to native language. Native language is stored in `Excel4Translator.native_language`.
        '''
        # Use default separator
        self.language = Excel4Translator.native_language
        self.translate()

    def set_language(self, lang):
        '''
        Sets `language` to `lang` and translates.

        :param lang: name of the language
        '''
        self.language = lang
        self.translate()

class Excel4FormulaArgument(Excel4Argument):
    '''
    Represents Excel4 formulas that should not have address (formulas that are not placed in cell as a call but for example as formula argument).

    '''
    def __init__(self, instruction, *args):
        '''

        :param instruction: name of Excel 4.0 instruction

        :param args: arguments of instruction
        '''
        Excel4Argument.__init__(self)
        # Save instruction name
        # Some instructions allow you to register a function under the given name.
        # If the instruction argument is an Excel4Instruction object, then the value returned by get_call_name will be taken as an instruction
        self.instruction = instruction

        # Instruction arguments
        self.args = list(args)

        # Save instruction name
        self.original_instruction = instruction
        if not issubclass(type(self.original_instruction),  Excel4Name):
            self.original_instruction = Excel4InstructionName(self.instruction, True)
            self.instruction = Excel4InstructionName(self.instruction, True)
        # Translate instruction
        self.translate()
        # List of objects in which self is used
        for a in args:
            try:
                a._add_reference(self)
            except:
                pass

    def get_instruction_translation(self, lang=None):
        '''
        Translates formula to `lang`. If lang is None then self.language is used.

        :param lang: Language in which instruction is returned. If lang is None then self.language is used.

        :return: translated formula
        '''
        if not lang:
            lang = self.language
        instruction = self.instruction
        if self.instruction.translate:
            instruction = Excel4InstructionName(Excel4Translator.translate(self.original_instruction.name, lang), True)
        return instruction

    def get_separator_translation(self, lang=None):
        '''
        Translates arguments separator. If lang is None then self.language is used.

        :param lang: Language in which arguments separator is returned. If lang is None then self.language is used.

        :return: translated arguments separator
        '''
        if not lang:
            lang = self.language
        return Excel4Translator.get_arguments_separator(lang)

    def translate(self):
        '''
        Translates formula language set in `language` property
        '''
        if self.instruction.translate:
            self.instruction = self.get_instruction_translation(self.language)
        self.args_sep = self.get_separator_translation(self.language)


    def revert_translation(self):
        '''
        Reverts translation of formula to native language. Native language is stored in `Excel4Translator.native_language`.
        '''
        # If Instruction is subclass of Excel4Instruction then we need to original function name
        self.instruction = self.original_instruction
        self.language = Excel4Translator.native_language
        # Use default separator
        self.args_sep = Excel4Translator.get_arguments_separator(Excel4Translator.native_language)


    def _replace_reference(self, ref, new_ref):
        '''
        Replaces `ref` with `new_ref`

        :param ref: old reference

        :param new_ref: new reference
        '''
        for i in range(0, len(self.args)):
            if self.args[i] == ref:
                self.args[i] = new_ref

    def add_argument(self, arg):
        '''
        Adds additional argument to formula

        :param arg: argument to add
        '''
        temp = list(self.args)
        temp.append(arg)
        self.args = tuple(temp)

    @staticmethod
    def parse_args(args_sep, args, lang=None):
        '''
        Parses formula arguments passed in `args` and converts them to string.
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing arguments of formula
        '''
        # Set arguments separator
        # Result
        func_args = ""
        for a in args:
            # Check type of argument
            if is_number(a):
                func_args = func_args + str(a)
            elif issubclass(type(a),  Excel4Argument):
                func_args = func_args + a.get_str(lang)
            elif issubclass(type(a), Cell):
                func_args = func_args + a.get_reference(lang)
            else:
                # Assume it's a string
                func_args = func_args + '"' + str(a) + '"'
            func_args = func_args + args_sep
        return func_args[0:-1]

    def _parse_args(self, lang=None):
        '''
        Parses formula arguments passed in `args` and converts them to string.
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing arguments of formula
        '''
        if not lang:
            lang = self.language

        args_sep = self.args_sep
        # Translate arguments separator if lang is different than self.language
        if lang and lang != self.language:
            args_sep = self.get_separator_translation(lang)
        return Excel4FormulaArgument.parse_args(args_sep, self.args, lang)

    def _get_func(self, lang=None):
        '''
        Return Excel4 instruction call
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing `Excel4FormulaArgument`
        '''
        if not lang:
            lang = self.language
        func_open = "("
        func_close = ")"
        func_args = self._parse_args(lang)
        instruction = self.instruction
        # Translate instruction if language is different than self.language
        if lang and lang != self.language:
            instruction = self.get_instruction_translation(lang)

        call = "{INSTRUCTION}{FUNCOPEN}{ARGS}{FUNCCLOSE}".format(
            **{
                "INSTRUCTION": instruction,
                "FUNCOPEN": func_open,
                "ARGS": func_args,
                "FUNCCLOSE": func_close
            }
        )
        return call

    def get_definition(self, lang=None):
        '''
        Return Excel4 instruction call
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing `Excel4FormulaArgument`
        '''
        return self._get_func(lang)

    def get_str(self, lang=None):
        '''
        Returns Excel4 instruction call
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing `Excel4FormulaArgument`
        '''
        return self._get_func(lang)

class Excel4LogicalTest(Excel4Argument):
    '''
    Represents Excel 4.0 logical test in form: VALUE1 OPERATOR VALUE2
    '''
    def __init__(self, value1, operator, value2):
        Excel4Argument.__init__(self)
        self.value1 = value1
        self.operator = operator
        self.value2 = value2

    def _get_arg_value(self, a, lang=None):
        '''
        Converts argument passed in `a` parameter to string. Argument is translated to language passed in `lang`.
        If `lang` is None then language from `language` property is taken.

        :param a: argument to convert

        :param lang: name of language

        :return: string representing argument `a`
        '''
        if not lang:
            lang = self.language

        if is_number(a):
            return str(a)
        elif issubclass(type(a), Excel4FormulaArgument):
            try:
                return a._get_func(lang)
            except:
                return str(a)

        elif issubclass(type(a), Excel4Instruction):
            return a.get_reference(lang)
        else:
            # Assume it's a string
            v = str(a)
            v = self._escape_value(v)
            return '"'+v+'"'

    def _parse_args(self, lang=None):
        '''
        Parses formula arguments and converts them to the string.
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :param lang: name of language

        :return: string representing `Excel4LogicalTest`
        '''
        func_args = "{VALUE1}{OPERATOR}{VALUE2}".format(**{
            "VALUE1":self._get_arg_value(self.value1, lang),
            "OPERATOR": self.operator,
            "VALUE2": self._get_arg_value(self.value2, lang),

        })
        return func_args

    def _get_argument(self, lang=None):
        '''
        Returns Excel4 logical test as a string.

        :return: string representing `Excel4LogicalTest`
        '''

        func_args = self._parse_args(lang)

        call = "{ARGS}".format(
            **{
                "ARGS": func_args,
            }
        )

        return call

