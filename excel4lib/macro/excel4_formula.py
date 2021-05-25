from .excel4_instruction import *
from .excel4_argument import *
from .excel4_result import *
from .excel4_variable import *
class Excel4Formula(Excel4Instruction):

    '''
    Represents Excel4 formula call (all formulas with '=' at the beginning =FORMULA(args...).)

    Excel4Formula is responsible for:
    - converting instruction and args to excel4 call;
    - translating instruction to selected language;
    - storing information about formula name and arguments.
    '''

    def __init__(self, x, y, instruction, *args):
        Excel4Instruction.__init__(self, x, y)
        # Save instruction name
        # Some instructions allow you to register a function under the given name.
        # If the instruction argument is an Excel4Instruction object, then the value returned by get_call_name will be taken as an instruction
        self.instruction = instruction

        # Arguments separator
        self.args_sep = ""
        # Instruction arguments
        self.args = list(args)

        # Save instruction name
        self.original_instruction = instruction
        if issubclass(type(self.original_instruction), Excel4Variable):
            self.original_instruction = instruction.name
            self.instruction = instruction.name
        elif not issubclass(type(self.original_instruction),  Excel4Name):
            self.original_instruction = Excel4InstructionName(self.instruction, True)
            self.instruction = Excel4InstructionName(self.instruction, True)

        # Translate instruction
        self.translate()
        # List of objects in which self is used
        self._references = []
        for a in args:
            try:
                a._add_reference(self)
            except:
                pass
    def _change_reference(self, new_ref):
        '''
        Changes all references to this object to `new_ref`.
        References are changed only in `Excel4Instruction` and `Excel4Argument` objects and their childs.

        :param new_ref: new reference
        '''
        for r in self._references:
            r._replace_reference(self, new_ref)

    def _add_reference(self, ref):
        '''
        Adds object reference to _references.

        :param ref: reference to add
        '''
        self._references.append(ref)

    def _replace_reference(self, ref, new_ref):
        '''
        Replaces `ref` with `new_ref`

        :param ref: old reference

        :param new_ref: new reference
        '''
        for i in range(0, len(self.args)):
            if self.args[i] == ref:
                self.args[i] = new_ref


    def get_instruction_translation(self, lang=None):
        '''
        Translates formula to `lang`

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
        Translates arguments separator to `lang`

        :param lang: Language in which arguments separator is returned. If lang is None then self.language is used.

        :return: translated arguments separator
        '''
        if not lang:
            lang = self.language
        return Excel4Translator.get_arguments_separator(lang)


    def translate(self):
        '''
        Translates formula to language stored in `language` property
        '''
        if self.instruction.translate:
            self.instruction = self.get_instruction_translation(self.language)
        self.args_sep = self.get_separator_translation(self.language)
        self.translate_address(self.language)

    def revert_translation(self):
        '''
        Reverts translation of formula to native language. Native language is stored in `Excel4Translator.native_language`.
        '''
        # If Instruction is subclass of Excel4Instruction then we need to original function name
        self.instruction = self.original_instruction
        self.language = Excel4Translator.native_language
        # Use default separator
        self.args_sep = Excel4Translator.get_arguments_separator(Excel4Translator.native_language)
        self.revert_address_translation()

    def add_argument(self, arg):
        '''
        Adds additional argument to formula

        :param arg: argument to add
        '''
        self.args.append(arg)

    def get_args_num(self):
        '''
        Returns number of arguments

        :return: number of arguments
        '''
        return len(self.args)

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
        Return Excel4 formula call with arguments
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing `Excel4Formula`
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

    def get_definition(self):
        '''
        Return Excel4 formula call with arguments

        :return: string representing `Excel4Formula`
        '''
        return self._get_func()


    def get_str(self, lang=None):
        '''
        Return Excel4 formula call with arguments
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing `Excel4Formula`
        '''
        return "="+self._get_func(lang)

    def __str__(self):
        '''
        Return Excel4 formula call with arguments

        :return: string representing `Excel4Formula`
        '''
        return self.get_str(self.language)


class Excel4LoopFormula(Excel4Formula):
    '''
    Represents Excel4 loop formula such as: WHILE, FOR
    '''
    def __init__(self, x, y, instruction, *args):
        Excel4Formula.__init__(self, x, y, instruction, *args)
        self._obfuscate = True
        self._obfuscate_formula = True
        self._spread = True

class Excel4ConditionFormula(Excel4Formula):
    '''
    Represents Excel4 condition formula  such as: IF
    '''
    def __init__(self, x, y, instruction, *args):
        Excel4Formula.__init__(self, x, y, instruction, *args)
        self._obfuscate = True
        self._obfuscate_formula = True
        self._spread = True

class Excel4EndLoopFormula(Excel4Formula):
    '''
    Represents Excel4 end loop formula such as: NEXT
    '''
    def __init__(self, x, y, instruction, *args):
        Excel4Formula.__init__(self, x, y, instruction, *args)
        # Do not obfuscate
        self._obfuscate = True
        self._obfuscate_formula = True
        self._spread = True


class Excel4EndConditionFormula(Excel4Formula):
    '''
    Represents Excel4 end condition formula  such as: NEXT
    '''
    def __init__(self, x, y, instruction, *args):
        Excel4Formula.__init__(self, x, y, instruction, *args)
        # Do not obfuscate
        self._obfuscate = True
        self._obfuscate_formula = True
        self._spread = True

class Excel4GoToFormula(Excel4Formula):
    '''
    Represents Excel4 GOTO formula
    '''
    def __init__(self, x, y, instruction, arg):
        Excel4Formula.__init__(self, x, y, instruction, arg)


    def _parse_args(self, lang=None):
        '''
        Parses formula arguments passed in `args` and converts them to string.
        All arguments are translated to language passed in `lang`. If `lang` is None then language from `language` property is taken.

        :return: string representing arguments of formula
        '''
        if not lang:
            lang = self.language
        # Result
        func_args = ""
        if len(self.args) < 1:
            return func_args

        a = self.args[0]
        # Check type of argument
        if is_number(a):
            func_args = func_args + str(a)
        elif issubclass(type(a),  Excel4Argument):
            func_args = func_args + a.get_str(lang)
        elif issubclass(type(a),  Excel4Result):
            func_args = func_args + a.get_reference(lang)
        elif issubclass(type(a),  Excel4Variable):
            func_args = func_args + a.get_address(lang)
        elif issubclass(type(a), Cell):
            func_args = func_args + a.get_reference(lang)
        else:
            # Assume it's a string
            func_args = func_args + '"' + str(a) + '"'
        return func_args