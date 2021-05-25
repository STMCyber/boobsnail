from .excel4_formula import *
from .excel4_name import *

class Excel4RegisterFormula(Excel4Formula):
    '''
    Represents Excel 4.0 Register formula. Also it sets basic parameters.
    ```
    REGISTER(dll_name, exported_function, type_text, function_text, "",1,9, ...)
    ```
    `Excel4RegisterFormula` class is used to register WinAPI functions.
    '''
    def __init__(self, x, y, dll_name, exported_function, type_text, function_text, *args):
        '''

        :param x: column number

        :param y: row number

        :param dll_name: name of a DLL

        :param exported_function: name of exported function that you want to import

        :param type_text:  string representing the types of return value and arguments of function that you want to import;

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        instruction = "REGISTER"
        self.function_text = Excel4InstructionName(function_text)
        Excel4Formula.__init__(self, x, y, instruction, dll_name, exported_function, type_text, self.function_text, "", 1, 9, *args)
        self.type_text = type_text
        self.dll_name = dll_name
        self.exported_function = exported_function

    def set_function_text(self, text):
        self.function_text.set_name(text)

    def get_exported_function(self):
        return self.exported_function

    def get_dll_name(self):
        return self.dll_name

    def get_type_text(self):
        return self.type_text

    def get_function_text(self):
        return self.function_text

    def get_translated_call_name(self):
        return self.get_function_text()

    def get_call_name(self):
        return self.get_function_text()