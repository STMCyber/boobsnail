from excel4lib.macro.excel4_macro_extension import *

class Excel4AntiAnalysis(Excel4MacroExtension):
    '''
    `Excel4AntiAnalysis` class allows to add anti-analysis tricks to worksheet.
    '''

    def __init__(self, macro):
        '''

        :param macro: Excel4Macro object
        '''
        Excel4MacroExtension.__init__(self)

    def get_sound_playing_check(self, formula1=None, formula2=None):
        '''
        Checks if computer is capable of playing sounds. This function does not add any object to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        if not formula1:
            formula1 = self.macro._create_argument_object("CLOSE", "TRUE")
        get_workspace = self.macro._create_argument_object("GET.WORKSPACE", 42)
        if not formula2:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, "", formula1)
        else:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, formula2, formula1)
        return result_formula


    def get_window_hidden_check(self, formula1=None, formula2=None):
        '''
        Checks if window is hidden. This function does not add any object to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)


        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        if not formula1:
            formula1 = self.macro._create_argument_object("CLOSE", "TRUE")
        get_workspace = self.macro._create_argument_object("GET.WORKSPACE", 7)
        if not formula2:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, "", formula1)
        else:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, formula2, formula1)
        return result_formula

    def get_mouse_check(self, formula1=None, formula2=None):
        '''
        Checks if mouse is present. This function does not add any object to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        if not formula1:
            formula1 = self.macro._create_argument_object("CLOSE", "TRUE")
        get_workspace = self.macro._create_argument_object("GET.WORKSPACE", 19)
        if not formula2:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, "", formula1)
        else:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, formula2, formula1)
        return result_formula

    def get_single_step_check(self, formula1=None, formula2=None):
        '''
        Checks if single-step mode is enabled. This function does not add any object to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        if not formula1:
            formula1 = self.macro._create_argument_object("CLOSE", "TRUE")
        get_workspace = self.macro._create_argument_object("GET.WORKSPACE", 31)
        if not formula2:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, formula1, "")
        else:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, formula1, formula2)
        return result_formula


    def get_sound_recording_check(self, formula1=None, formula2=None):
        '''
        Checks if computer is capable of recording sounds. This function does not add any object to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        if not formula1:
            formula1 = self.macro._create_argument_object("CLOSE", "TRUE")
        get_workspace = self.macro._create_argument_object("GET.WORKSPACE", 43)
        if not formula2:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, "", formula1)
        else:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", get_workspace, formula2, formula1)
        return result_formula

    def get_windows_check(self, formula1=None, formula2=None):
        '''
        Checks if Excel is running on Windows. This function does not add any object to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        if not formula1:
            formula1 = self.macro._create_argument_object("CLOSE", "TRUE")
        get_workspace = self.macro._create_argument_object("GET.WORKSPACE", 1)
        isnumber = self.macro._create_argument_object("ISNUMBER", self.macro._create_argument_object("SEARCH", "Windows", get_workspace))
        if not formula2:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", isnumber, "", formula1)
        else:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", isnumber, formula2, formula1)
        return result_formula

    def get_file_name_check(self, file_name, formula1=None, formula2=None):
        '''
        Checks if document name is test.xlsm. This function does not add any object to the worksheet.

        :param file_name: file name to check

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        if not formula1:
            formula1 = self.macro._create_argument_object("CLOSE", "TRUE")
        get_document = self.macro._create_argument_object("GET.DOCUMENT", 88)
        if not formula2:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", self.macro.logical(get_document, "<>", file_name), formula1, "")
        else:
            result_formula = self.macro._create_condition(self.macro.worksheet._curr_x, self.macro.worksheet._curr_y, "IF", self.macro.logical(get_document, "<>", file_name), formula1, formula2)
        return result_formula


    def add_sound_playing_check(self, formula1=None, formula2=None):
        '''
        Checks if computer is capable of playing sounds. This function adds Excel4Condition to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        formula = self.get_sound_playing_check(formula1, formula2)
        self.macro._add_to_worksheet(formula)
        return formula

    def add_window_hidden_check(self, formula1=None, formula2=None):
        '''
        Checks if window is hidden. This function adds Excel4Condition to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)


        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        formula = self.get_window_hidden_check(formula1, formula2)
        self.macro._add_to_worksheet(formula)
        return formula

    def add_mouse_check(self, formula1=None, formula2=None):
        '''
        Checks if mouse is present. This function adds Excel4Condition to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        formula = self.get_mouse_check(formula1, formula2)
        self.macro._add_to_worksheet(formula)
        return formula

    def add_single_step_check(self, formula1=None, formula2=None):
        '''
        Checks if single-step mode is enabled. This function adds Excel4Condition to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        formula = self.get_single_step_check(formula1, formula2)
        self.macro._add_to_worksheet(formula)
        return formula

    def add_sound_recording_check(self, formula1=None, formula2=None):
        '''
        Checks if computer is capable of recording sounds. This function adds Excel4Condition to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        formula = self.get_sound_recording_check(formula1, formula2)
        self.macro._add_to_worksheet(formula)
        return formula

    def add_windows_check(self, formula1=None, formula2=None):
        '''
        Checks if Excel is running on Windows. This function adds Excel4Condition to the worksheet.

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        formula = self.get_windows_check(formula1, formula2)
        self.macro._add_to_worksheet(formula)
        return formula

    def add_file_name_check(self, file_name, formula1=None, formula2=None):
        '''
        Checks if document name is test.xlsm. This function adds Excel4Condition to the worksheet.

        :param file_name: file name to check

        :param formula1: formula to execute if analysis is detected. By default CLOSE(TRUE)

        :param formula2: formula to execute if analysis is not detected

        :return: Excel4Condition object
        '''
        formula = self.get_file_name_check(file_name, formula1, formula2)
        self.macro._add_to_worksheet(formula)
        return formula