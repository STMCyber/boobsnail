from excel4lib.macro.obfuscator import Excel4Obfuscator
from excel4lib.config import Excel4Config
from excel4lib.macro.excel4_instruction import  *
from excel4lib.macro.excel4_argument import  *
from excel4lib.macro.excel4_formula import  *
from excel4lib.macro.excel4_value import *
from excel4lib.macro.excel4_variable import *
from excel4lib.macro.routine import Excel4RC4RoutineStr

class Excel4Rc4Obfuscator(Excel4Obfuscator):

    name = "rc4"
    description = "Extends standard obfuscator and allows to obfuscate macro formulas with RC4 cipher"


    '''
    Allows to obfuscate formulas with RC4 encryption.
    '''
    def __init__(self, config=Excel4Config.obfuscator):
        Excel4Obfuscator.__init__(self, config)


        # RC4 keystream, used in RC4 formula obfuscation
        self._rc4_keystream = []
        self._rc4_key_ptr = None
        self.rc4_init_func = None
        self.rc4_decrypt_func = None
        self._rc4_routine = None


    def _obfuscate_formula(self, formula):
        '''
        Obfuscates formula object. Should work for all classes that inherit from Excel4Formula.
        _obfuscate_formula works as follow:
        - convert formula with arguments to string ex: =TEST("A","B");
        - for each character in formula string:
         - obfuscate character with random function such as: MID, SUM, MOD etc.
        - end of loop
        - generate CONCATENATE formulas in order to concatenate all characters during excel 4.0 macro execution;
        - pass deobfuscated string to FORMULA call ex: =FORMULA(DEOBFUSCATED, ADDRESS_TO_SAVE_FORMULA)

        :param formula:
        :return:
        '''
        if not formula._obfuscate_formula:
            return  []
        formulas = self._obfuscate_text_rc4(str(formula), formula.tag)

        # Obfusacted formula will be deobfusacted and saved into following cell
        # So you can get result of your formula from this cell
        # This one is empty because cell will be filled after excel 4.0 macro execute
        # @HACK
        if issubclass(type(formula), Excel4LoopFormula):
            call_reference = Excel4ResultLoop(0, 0)
        elif issubclass(type(formula), Excel4ConditionFormula):
            call_reference = Excel4ResultCondition(0, 0)
        elif issubclass(type(formula), Excel4EndLoopFormula):
            call_reference = Excel4ResultEndLoop(0, 0)
        else:
            call_reference = self._create_result_formula(0, 0)
        call_reference.tag = formula.tag
        result_formula = self._create_formula(0, 0, "FORMULA", formulas[-1], call_reference)
        call_reference.start_cell = formulas[0]
        formulas.append(result_formula)
        formulas.append(call_reference)
        return formulas

    def init_rc4(self, key):
        '''
        Initializes RC4 encryption

        :param key: string key to use in encryption process

        :return:
        '''
        self._rc4_keystream = RC4.get_keystream(key)
        self._rc4_key_ptr = self.macro._create_variable(self.worksheet._curr_x, self.worksheet._curr_y, random_string(random.randint(5,20)), key)
        self._rc4_routine = Excel4RC4RoutineStr(self.macro)
        value = self.macro._create_variable(self.worksheet._curr_x, self.worksheet._curr_y, random_string(10),
                "temp")
        self._rc4_routine.set_macro_arguments(value, self._rc4_key_ptr)
        t_table, init_routine_formulas, formulas = self._rc4_routine.generate()
        self.rc4_init_func= self._rc4_routine.ref_init()
        self.rc4_decrypt_func = self._rc4_routine.ref()
        return (t_table, init_routine_formulas, formulas, self._rc4_key_ptr, self.rc4_init_func, self.rc4_decrypt_func)

    def _obfuscate_text_rc4(self, text, tag=""):
        '''
        Obfuscates `text` with RC4 encryption

        :param text: string to obfuscate

        :return: list of formulas that deobufscates `text`
        '''
        formulas = []
        encrypted = RC4.encrypt_ks(self._rc4_keystream, text)
        if len(encrypted) > 120:
            dec_calls = []
            for block in split_blocks(encrypted, 120):
                value = self.macro._create_variable(self.worksheet._curr_x, self.worksheet._curr_y, random_string(10),
                                                    block)
                self._rc4_routine.set_macro_arguments(value, self._rc4_key_ptr)
                self.rc4_decrypt_func = self._rc4_routine.ref()
                decrypt_call = self.macro._create_formula(-1,-1, self.rc4_decrypt_func.name)
                formulas.append(value)
                formulas.append(decrypt_call)
                dec_calls.append(decrypt_call)
            formulas.append(self.macro._create_formula(-1,-1, "CONCATENATE", *dec_calls))
        else:
            value = self.macro._create_variable(self.worksheet._curr_x, self.worksheet._curr_y, random_string(10), encrypted)
            self._rc4_routine.set_macro_arguments(value, self._rc4_key_ptr)
            self.rc4_decrypt_func = self._rc4_routine.ref()
            formulas.append(value)
            formulas.append(self.macro._create_formula(-1,-1, self.rc4_decrypt_func.name))
        return formulas


    def obfuscate_formulas(self, fomulas):
        '''
        Obfuscates formulas. Formulas should be ordered by execution. This function also adds obfusacted formulas to the worksheet.

        :param fomulas: list of formulas to obfuscate
        '''

        obfuscated_formulas = []
        column = self.worksheet.get_column(self.macro.trigger_x)
        if not column:
            # Raise exception
            return
        start_of_macro = None
        # Find first call in trigger column
        for cell in column.values():
            if cell.y >= self.macro.trigger_y:
                start_of_macro = cell
                break
        t_table, init_routine_formulas, decrypt_routine_formulas, rc4_key_ptr, rc4_init_func, rc4_decrypt_func = self.init_rc4(
            RC4.get_key(random.randint(16, 24)))
        obfuscated_formulas.append(rc4_key_ptr)
        obfuscated_formulas.append(rc4_init_func)
        obfuscated_formulas.append(self.macro._create_formula(-1, -1, rc4_init_func.name))
        obfuscated_formulas.append(rc4_decrypt_func)
        for f in fomulas:
            if not f._obfuscate:
                continue

            if issubclass(type(f), Excel4Formula):
                # Obfuscate formula call
                obfuscated = self.obfuscate_formula(f)
                # If formula is not obfuscated, then continue execution
                if not obfuscated:
                    continue
                # If obfuscated, then formula is splitted into multiple CONCATENATE and FORMULA calls.
                # Last element in obfuscated variable(list) is cell in which deobfuscated formula
                # will be placed during macro execution.
                obfuscated_ref = obfuscated[-1]
                # Add all formulas to worksheet
                for o in obfuscated[:-1]:
                    # Method 2
                    obfuscated_formulas.append(o)
                # Replace formula f with obfuscated_ref in worksheet
                self.worksheet.replace_cell(f, obfuscated_ref)
                # Update references
                # Formula object (f) holds references to objects in which is used.
                # We need to change these references to point to the obfuscated_ref
                obfuscated_ref.x = f.x
                obfuscated_ref.y = f.y
                f._change_reference(obfuscated_ref)
                if start_of_macro == f:
                    start_of_macro = obfuscated_ref
        # Spread obfuscated formulas across the worksheet
        self._spread_column(obfuscated_formulas, self.macro.trigger_x, self.macro.trigger_y, start_of_macro)
        # Redirect macro 4 execution to deobfuscate formulas
        self.worksheet.replace_cell(start_of_macro, self.macro._create_go_to(start_of_macro.x, start_of_macro.y, obfuscated_formulas[0]))
        # After deobfuscation redirect execution to original formula
        start_of_macro.x = obfuscated_formulas[-1].x
        start_of_macro.y = obfuscated_formulas[-1].y + 1
        return_stm = self.worksheet.get_cell(obfuscated_formulas[-1].x, obfuscated_formulas[-1].y+1)

        #self.worksheet.add_above(self._create_formula(return_stm,return_stm.y, "PAUSE"), return_stm)
        self.worksheet.add_above(start_of_macro, return_stm)

        self.macro.random_add_to_worksheet(t_table)
        self.macro.random_add_to_worksheet(init_routine_formulas)
        self.macro.random_add_to_worksheet(decrypt_routine_formulas)
