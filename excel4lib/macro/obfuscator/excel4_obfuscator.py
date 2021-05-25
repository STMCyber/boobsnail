import random

from excel4lib.macro.excel4_macro_extension import *

from excel4lib.utils import *
from excel4lib.macro.excel4_instruction import  *
from excel4lib.macro.excel4_argument import  *
from excel4lib.macro.excel4_formula import  *
from excel4lib.macro.excel4_value import *
from excel4lib.macro.excel4_variable import *
from excel4lib.exception import *
from excel4lib.config import *
from excel4lib.macro.excel4_result import *
from excel4lib.macro.excel4_register_formula import *
from excel4lib.sheet import *

class Excel4Obfuscator(Excel4MacroExtension):

    name = "standard"
    description = "Allows to obfuscate macro with standard Excel4.0 formulas suchas BITXOR, SUM, MID etc."

    '''
    Allows to obfuscate formulas,  scatter them across worksheet, obfuscate variable names and values.
    '''
    def __init__(self, config=Excel4Config.obfuscator):
        Excel4MacroExtension.__init__(self)

        # Obfuscator configuration
        self.config = config
        # List of char obfuscation methods
        self.ob_tech = []
        # Max length of cell
        self.cell_max_length = self.config.cell_limit


    def _generate_noise(self, only_empty = False):
        '''
        Adds random values to worksheet cells

        :param only_empty: flags that tells if add noise only to empty cells ( not reserved)
        '''
        # Walk through worksheet cell by cell
        for cords in self.worksheet.worksheet_iterate():
            # Choose whether add noise to this cell or not
            if random.randint(0,10) != 1:
                continue
            noise = random_string(random.randint(4, 20))
            noise_cell = self._create_value(cords[0], cords[1], noise)
            # Check if cell is reserved
            if self.worksheet.is_reserved(cords[0], cords[1]):
                if only_empty:
                    continue
                # Check if obfuscation of cell is enabled
                cell = self.worksheet.get_cell(cords[0], cords[1])
                if not cell._spread or not cell._obfuscate:
                    continue
                try:
                    # Move cell to the next cell if reserved
                    self.worksheet.move_cell(cell)
                except CouldNotMoveCellException as ex:
                    continue
            curr_cords = self.worksheet.get_current_cords()
            # Add noise
            self.worksheet.add_cell(noise_cell)
            self.worksheet.set_current_cords(curr_cords[0], curr_cords[1])

    def _spread_formulas(self, trigger_x, trigger_y):
        '''
        Spreads formulas across cells in worksheet

        :param trigger_x: number of column in which first call is placed

        :param trigger_y: number of row in which first call is placed

        '''
        # Get current cords. We need to remember current cord cause we will want to back execution to this cell.
        cords_backup = self.worksheet.get_current_cords()
        cells_cache = {}


        # Get cells to spread
        # For each column in worksheet
        for t in self.worksheet.column_iterate():
            # Get column number
            c_num = t[0]
            # Get cells in column
            cells_temp = t[1]
            if not cells_temp:
                continue
            values = cells_temp.values()
            # For each cell in column
            for cell in values:
                # Check if obfuscation of cell/formula is enabled
                if (not cell._spread) or (not cell._obfuscate):
                    continue
                # Save cell in cache
                try:
                    cells_cache[c_num][cell.y] = cell
                except KeyError:
                    cells_cache[c_num] = {cell.y : cell}

        # Remove cells from worksheet
        # For each column in cache
        for c in cells_cache.keys():
            for cell in cells_cache[c].values():
                # Remove cell from worksheet
                # x,y of cell will be changed, and cell will be placed at another cords
                self.worksheet.remove_cell(cell)

        trigger_cell = None
        # Add jump to first call
        if trigger_x in cells_cache:
            # Find first call
            for row in cells_cache[trigger_x]:
                if row >= trigger_y:
                    trigger_cell = self._go_to(trigger_x, row, cells_cache[trigger_x][row])
                    self.worksheet.add_cell(trigger_cell)
                    break

        # Spread cells across worksheet
        # For each column in cache
        for c in cells_cache.keys():
            self._spread_column(list(cells_cache[c].values()), trigger_x, trigger_y, trigger_cell)

        # Restore original cords
        self.worksheet.set_current_cords(cords_backup[0], cords_backup[1])

    def _spread_column(self, cells, trigger_x, trigger_y, trigger_cell):
        '''
        Spread `cells` across worksheet.
        :param cells: list of cells that are in the same column

        :param trigger_x: auto_open or auto_close function column

        :param trigger_y: auto_open or auto_close function row

        :param trigger_cell: auto_open or auto_close cell

        '''
        # Number of cells
        cells_num = len(cells)
        # The number of formulas that have already been spread across sheet
        cnt = 0
        fail_cnt = 0
        for_loop_cache = []
        if not cells:
            return
        while cnt < cells_num:
            # Generate random cords
            # IF all columns are reserved then add new one and place payoad there
            if fail_cnt > 1000:
                self.config.spread_x_max = self.config.spread_x_max + 1
                target_x = self.config.spread_x_max
            else:
                target_x = random.randint(self.config.spread_x_min, self.config.spread_x_max)
            target_y = random.randint(self.config.spread_y_min, self.config.spread_y_max)

            # Space between auto_open/auto_close cell and first call should be empty
            if(target_x == trigger_x) and (target_y in range(trigger_y, trigger_cell.y)):
                continue
            # If the same coordinates are drawn then randomize again
            if (target_x == self.worksheet._curr_x) and (target_y == self.worksheet._curr_y):
                # Inc failure counter
                fail_cnt = fail_cnt + 1
                continue

            height = random.randint(1, cells_num - cnt)

            # Check if cells are free
            # We need to add 1 to height since we need additional cell for GOTO formula
            if self.worksheet.is_reserved(target_x, target_y, height + 1 + 1 + 1):
                # Inc failure counter
                fail_cnt = fail_cnt + 1
                continue

            self.worksheet.set_current_cords(target_x, target_y)
            cnt_h = cnt+height
            # Add random number of cells to worksheet at random cords
            for cell in cells[cnt:cnt_h]:
                # Loops require end statement in the same column
                # So we need to place them in the same one
                if issubclass(type(cell), Excel4LoopFormula):
                    # Save column and row number of this loop
                    for_loop_cache.append((self.worksheet._curr_x, self.worksheet._curr_y + (cnt_h - cnt) + 2))
                elif issubclass(type(cell), Excel4EndLoopFormula):
                    break

                self.worksheet.add_next_cell(cell)
                cnt = cnt + 1
            # If there are more cells to spread
            if cnt < cells_num:
                # @HACK
                # If cells[cnt] is Ecel4Variable then get_reference function will return variable name
                # But if this variable name is not defined or this variable name is not storing address of another cell
                # then we can't GOTO to this formula(we need address of this formula).
                # So to bypass this we need to add an empty Excel4Value, because then get_reference function will return address of cell
                if issubclass(type(cells[cnt]), Excel4Variable):
                    empty = self._create_empty_formula(cells[cnt].x, cells[cnt].y)
                    cells.insert(cnt, empty)
                    cells_num = cells_num + 1
                # If there are more cells to spread, then redirect macro execution
                # to the next cell.
                self.worksheet.add_next_cell(self._go_to(-1, -1, cells[cnt]))
            else:
                break

            if issubclass(type(cells[cnt]), Excel4EndLoopFormula):
                if len(for_loop_cache) < 1:
                    raise Excel4LoopFormulaMissing("Excel4EndLoopFormula detected but Excel4LoopFormula is missing. Excel4 requires that the loops and NEXT() formula be in the same column.")
                cords = for_loop_cache.pop()
                cells[cnt].x = cords[0]
                cells[cnt].y = cords[1]
                self.worksheet.add_cell(cells[cnt])
                cnt = cnt + 1
                # If there are more cells to spread
                if cnt < cells_num:
                    # @HACK
                    # If cells[cnt] is Ecel4Variable then get_reference function will return variable name
                    # But if this variable name is not defined or this variable name is not storing address of another cell
                    # then we can't GOTO to this formula(we need address of this formula).
                    # So to bypass this we need to add an empty Excel4Value, because then get_reference function will return address of cell
                    if issubclass(type(cells[cnt]), Excel4Variable):
                        empty = self._create_empty_formula(cells[cnt].x, cells[cnt].y)
                        cells.insert(cnt, empty)
                        cells_num = cells_num + 1
                    # If there are more cells to spread, then redirect macro execution
                    # to the next cell.
                    self.worksheet.add_cell(self._go_to(cells[cnt-1].x, cells[cnt-1].y + 1, cells[cnt]))
                else:
                    break
        # ADD RETURN
        self.worksheet.add_next_cell(self._create_formula(-1, -1, "RETURN"))

    def _create_argument_object(self, instruction, *args):
        instruction_name = Excel4InstructionName(instruction, self.config.translate)
        o = Excel4FormulaArgument(instruction_name, *args)
        if not self.config.translate:
            # Do not translate obfuscator objects
            o.revert_translation()
        return o

    def _create_formula(self,x, y, instruction, *args):
        instruction_name = Excel4InstructionName(instruction, self.config.translate)
        o = Excel4Formula(x, y, instruction_name, *args)
        if not self.config.translate:
            # Do not translate obfuscator objects
            o.revert_translation()
        return o

    def _create_value(self, x, y, value):
        o = Excel4Value(x,y, value)
        if not self.config.translate:
            # Do not translate obfuscator objects
            o.revert_address_translation()
        return o

    def _create_empty_formula(self, x, y):
        return self._create_value(x,y,"")

    def _create_result_formula(self, x, y):
        o = Excel4Result(x,y)
        if not self.config.translate:
            # Do not translate obfuscator objects
            o.revert_address_translation()
        return o

    def _go_to(self, x, y, formula):
        instruction_name = Excel4InstructionName("GOTO", self.config.translate)
        o = Excel4GoToFormula(x, y, instruction_name, formula)
        if not self.config.translate:
            # Do not translate obfuscator objects
            o.revert_translation()
        return o

    def _char(self, s):
        '''
        Returns CHAR formula
        :param s: string, char
        :return:
        '''
        return self._create_argument_object("CHAR", s)

    def char(self, c):
        '''
        Puts c character in CHAR formula

        :param c: charcater

        :return: CHAR formula call
        '''
        if not is_number(c):
            c = ord(c)
        return self._char(c)

    def int(self, c):
        '''
        Converts c character to CHAR(INT(C)) call

        :param c: charcater

        :return:
        '''
        if not is_number(c):
            c = ord(c)

        return self._char(self._create_argument_object("INT", str(c)))

    def sum(self, c):
        '''
        Converts c character to CHAR(SUM(R, c-k/k-c) call

        :param c: charcater

        :return:
        '''
        if not is_number(c):
            c = ord(c)

        k = random.randint(1, 1000)
        while k == c:
            k = random.randint(1, 1000)

        if k < c:
            r = c - k
        else:
            r = k - c

        return self._char(self._create_argument_object("SUM", r, k))

    def mid(self, c):
        '''
        Converts c character to MID(STR, RAND_INDEX,1) call

        :param c: charcater

        :return:
        '''
        if is_number(c):
            c = chr(c)

        length = random.randint(3, 8)
        rand_str = random_string(length)
        rand_ind = random.randint(0, length)
        rand_str = rand_str[:rand_ind] + c + rand_str[rand_ind:]

        return self._create_argument_object("MID", rand_str, rand_ind+1, 1)

    def xor(self, c):
        '''
        Converts c character to CHAR(BITXOR(R,K) call

        :param c: charcater

        :return:
        '''
        if not is_number(c):
            c = ord(c)

        k = random.randint(1, 1000)
        while k == c:
            k = random.randint(1, 1000)
        r = k ^ c
        return self._char(self._create_argument_object("BITXOR", r, k))

    def mod(self, c):
        '''
        Converts c character to CHAR(MOD(R,K) call

        :param c: charcater

        :return:
        '''
        if not is_number(c):
            c = ord(c)
        k = random.randint(c + 1, 600)
        r = k + c
        return self._char(self._create_argument_object("MOD", r, k))

    def evaluate(self, c):
        #=CHAR(EVALUATE("1+64"))
        pass

    def concat(self, x, y, *args):
        '''
        Creates CONCATENATE formula
        '''
        return self._create_formula(x, y, "CONCATENATE", *args)

    def _update_obfuscation_tech(self):
        self.ob_tech = []
        if self.config.obfuscate_int:
            self.ob_tech.append(self.int)
        if self.config.obfuscate_char:
            self.ob_tech.append(self.char)
        if self.config.obfuscate_mid:
            self.ob_tech.append(self.mid)
        if self.config.obfuscate_xor:
            self.ob_tech.append(self.xor)
        if self.config.obfuscate_mod:
            self.ob_tech.append(self.mod)

    def _obfuscate_char(self, c):

        if c == '"':
            formula = self.char(c)
        else:
            self._update_obfuscation_tech()
            func = random.choice(self.ob_tech)
            formula = func(c)

        return formula

    def _split_instructions(self, objects, block_size):
        i = 0
        temp = []
        for o in objects:
            if issubclass(type(o),  Excel4FormulaArgument):
                # For arguments we need to compute length of full instruction with arguments
                o_len = o.get_length()
            elif issubclass(type(o), Cell):
                # For calls (instructions in cells) we will use references so length we be equal to address of the cell
                o_len = len(o.get_address())
            else:
                raise Excel4UnsupportedTypeException("Received unsupported type: {}".format(type(o)))

            if (i+o_len) > block_size:
                yield temp
                i = 0
                temp = []
            i = i + o_len
            temp.append(o)
        if temp:
            yield temp

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
        formulas = self._obfuscate_text(str(formula), formula.tag)

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


    def obfuscate_formula(self, formula):
        '''

        :param formula:
        :return:
        '''

        formulas = self._obfuscate_formula(formula)
        if formulas:
            formulas[-1].x = formula.x
            formulas[-1].y = formula.y
        return formulas

    def obfuscate_formulas(self, fomulas):
        '''
        Obfuscates formulas. Formulas should be ordered by execution. This function also adds obfusacted formulas to the worksheet.
        :param fomulas:
        :return:
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
                    #self.worksheet.add_above(o, f)
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


    def obfuscate_function_name(self, formula):
        if not issubclass(type(formula), Excel4RegisterFormula):
            raise Excel4WrongVariableType("Only instances of Excel4RegisterFormula could be obfuscated. Provided: {}".format(type(formula)))

        if formula._obfuscate:
            formula.set_function_text(random_string(random.randint(4, 10)))

        return formula

    def obfuscate_variable_name(self, formula):
        if not issubclass(type(formula), Excel4Variable):
            raise Excel4WrongVariableType("Only instances of Excel4Variable could be obfuscated. Provided: {}".format(type(formula)))

        if formula._obfuscate:
            formula.set_name(random_string(random.randint(4, 10)))

        return formula

    def obfuscate_variable_value(self, formula):
        if not issubclass(type(formula), Excel4Variable):
            raise Excel4WrongVariableType("Only instances of Excel4Variable could be obfuscated. Provided: {}".format(type(formula)))

        formulas = []
        if formula._obfuscate:
            # Do not obfuscate numbers
            try:
                if is_number(formula.value):
                    raise Exception("Obfuscation of numbers not supported")
                elif issubclass(type(formula.value), Cell):
                    # Obufscate address
                    raise Exception("Obfuscation of Cell objects not supported")

                formulas = self._obfuscate_text(formula.value, formula.tag)
                # Set value as address of cell in which deobfuscated value will be saved
                formula.value = formulas[-1]
            except:
                pass
        return formulas

    def _obfuscate_variable(self, formula):
        # Obfuscate variable name
        #formula.name = random_string(random.randint(4,10))
        # Obfuscate variable value if it's behaviour is similar to str
        formulas = []
        try:
            if issubclass(type(formula.value), Excel4Value):
                if formula.value.is_num():
                    raise Exception()
            formulas = self._obfuscate_text(formula.value, formula.tag)
            # Set value as address of cell in which deobfuscated value will be saved
            formula.value = formulas[-1]
        except:
            pass
        # Add variable as last call
        formulas.append(formula)
        return formulas




    def _obfuscate_text(self, text, tag=""):
        '''
        Obfuscates every char in text and returns concat formulas that allows to restore original string
        :param text: string to obfuscate
        :return:
        '''

        text_len = len(text)
        block_len = self.cell_max_length
        # Obfusacted characters
        obfuscated_chars = []
        # List of concat formulas which allows to restore original string
        concat_objects = [[]]
        # Obfusacte every char in text
        for i in range(0, text_len):
            c = text[i]
            obfuscated_chars.append(self._obfuscate_char(c))

        # Pass obfusacted chars to concat formula
        for o in self._split_instructions(obfuscated_chars, block_len):
            concat_objects[0].append(self.concat(0, 0, *o))

        # If there is more than one concatenate formula then we need to concatenate these formulas
        # Here is a little bug. At this stage we don't know what address has concatenate formula.
        # So we don't know the exact length of this formula.
        i = 0
        while True:
            concat_objects.append([])
            for o in self._split_instructions(concat_objects[i], block_len):
                concat_objects[i + 1].append(self.concat(0, 0, *o))
            if len(concat_objects[i + 1]) < 2:
                break
            i = i + 1
        r = []
        if not tag:
            tag = random_tag(text)
        for o in concat_objects:
            for x in o:
                x.tag = tag
                r.append(x)
        return r



