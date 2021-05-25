from excel4lib.macro import Excel4Variable, Excel4VariableName
from excel4lib.macro.excel4_macro_extension import *
from excel4lib.exception import *
from excel4lib.utils import *

class Excel4Routine(object):
    '''
    Base class for classes implementing predefined Excel4 macro that can be embedded in the worksheet
    '''

    def __init__(self, name, macro):
        '''

        :param name: name of the routine
        :param macro: Excel4Macro object
        '''
        # Name of routine
        self.name = name
        # First instruction of routine
        self.routine_ptr = None
        # Last instruction of routine (most often RETURN formula)
        self.routine_end_ptr = None
        # Excel4Macro object
        self.macro = macro
        # Is macro added to worksheet
        self._enabled = False
        # Name of routine in macro
        self.call_name = random_string(random.randint(10,25))
        # Routine's arguments
        self.args = {}

    def set_enabled(self):
        '''
        Sets routine as enabled - as added to the worksheet.
        '''
        self._enabled = True

    def is_enabled(self):
        '''
        Return True if routine is added to worksheet and False if not.

        :return: True if routine is added to worksheet and False if not.
        '''
        return self._enabled

    def get_argument(self, name):
        '''
        Returns argument with `name`

        :param name: name of argument to return

        :return: argument with `name`
        '''
        arg = self.args.get(name, None)
        if not arg:
            raise Excel4RoutineMissingArgument("Missing required argument {} in routine {}. Set this argument by calling set_macro_arguments function.".format(name, self.name))
        return arg["references"][0]

    def set_routine_macro_arguments(self, **kwargs):
        '''
        Sets arguments that will be used during macro execution. Note that all names of Excel4Variables passed to this function will be changed

        :param kwargs: arguments
        '''
        for key in kwargs:
            var = kwargs[key]
            if not issubclass(type(var), Excel4Variable):
                # Wrong type of argument
                continue

            if key not in self.args:
                self.args[key] = {"name": Excel4VariableName(random_string(random.randint(5,20))), "references":[var]}
            else:
                self.args[key]["references"].append(var)
            var.set_name(self.args[key]["name"].name)
            # Disable obfuscation of this name
            # @TODO In the future implement mechanism that will change all references to this name object
            var.name._obfuscate = False


    def generate(self, **kwargs):
        '''
        Generates routine formulas and returns them

        :param kwargs: arguments
        '''
        pass

    def add(self, **kwargs):
        '''
        Adds routine to the worksheet

        :param kwargs: arguments
        '''
        pass

    def ref(self):
        '''
        Returns variable that points to routine start.

        :return: Excel4Variable pointing to start of the routine in macro
        '''
        return self.macro._create_variable(-1, -1, self.call_name, self.routine_ptr)

class Excel4RC4RoutineStr(Excel4Routine):
    '''
    `Excel4RC4RoutineStr` class implements Excel4 Macro that allows to encrypt/decrypt data by using RC4 encryption.

     In order to use RC4 routine you need to:


     1. Place key and ciphertext in macro

     2. Call set_macro_arguments function on RC4 routine and pass to this function key and ciphertext variables.

     3. Get addresses of initilization and decryption functions.

     4. Call initialization function in macro.

     5. Call decrypt function in macro.


     Example code:

    ```
    from excel4lib.macro import *
    from excel4lib.macro.obfuscator import *
    from excel4lib.utils import *
    from excel4lib.macro.routine import *
    # Init routines
    routines = Excel4Routines()
    # Create macro object
    macro = Excel4Macro("test.csv",routines=routines)
    # Generate Key and encrypt plaintext
    key = RC4.get_key(10)
    cipher = RC4.encrypt(key, "ABCD"*5)
    # Add key and cipher to the worksheet
    text = macro.variable("text", cipher)
    key_ptr = macro.variable("key", key)
    # Set text and key as arguments in RC4 macro
    macro.routines.get("Excel4RC4RoutineStr").set_macro_arguments(text, key_ptr)
    # Get address of initialization and decryption function
    init_func, decrypt_func = macro.routines.get("Excel4RC4RoutineStr").add()
    # Call RC4 initilization function
    macro.formula(init_func.name)
    # Call decrypt function and display plaintext in ALERT
    macro.formula("ALERT", macro.argument(decrypt_func.name))
    macro.to_csv_file("out.csv")
    ```
    '''
    def __init__(self, macro):
        '''

        :param macro: Excel4Macro object
        '''
        Excel4Routine.__init__(self, "Excel4RC4RoutineStr", macro)
        self.initialization_routine_name = random_string(random.randint(10,25))
        self.initialization_routine_ptr = None

        self._p1 = None
        self._p2 = None

    def set_macro_arguments(self, input, key):
        '''
        Sets name of input and key variables

        :param input: Excel4Variable that stores data to decrypt

        :param key: Excel4Variable that stores key to use during decryption process
        '''
        self.set_routine_macro_arguments(input=input, key=key)


    def allocate_init_table(self):
        '''
        Generates formulas that initialize T table

        :return: list of Excel4Value representing elements of T table
        '''
        formulas = []
        for i in range(0, 256):
            val = self.macro._create_value(-1, -1, "")
            val._obfuscate = False
            formulas.append(val)
        return formulas

    def generate_init_routine(self, t_table):
        '''
        Generates RC4 initialization routine

        :param t_table: list of values generated by allocate_init_table

        :return: list of formulas initializing RC4
        '''
        formulas = []
        self.initialization_routine_ptr = self.macro._create_value(-1, -1, random_string(random.randint(2, 25)))
        formulas.append(self.initialization_routine_ptr)
        cnt = self.macro._create_variable(-1, -1, random_string(random.randint(10, 25)), 0)
        formulas.append(cnt)
        formulas.append(self.macro._create_loop(-1,-1, "FOR", cnt.name, cnt, 255, 1))
        formulas.append(self.macro._create_formula(-1,-1, "SET.VALUE", self.macro.argument("OFFSET", t_table[0], cnt, 0), cnt))
        formulas.append(self.macro._create_end_loop(-1,-1, "NEXT"))

        temp = self.macro._create_variable(-1, -1, random_string(random.randint(10, 25)), 0)
        formulas.append(temp)
        counter = self.macro._create_variable(-1, -1, random_string(random.randint(10, 25)), 0)
        formulas.append(counter)
        key_len = self.macro.argument("LEN", self.get_argument("key"))
        formulas.append(self.macro._create_loop(-1,-1, "FOR", counter.name, counter, 255, 1))
        t_table_ref = self.macro.argument("OFFSET", t_table[0], counter, 0)

        key_mod = self.macro.argument("SUM", self.macro.argument("MOD", counter, key_len), 1)
        k_table_ref = self.macro.argument("CODE", self.macro.argument("MID", self.get_argument("key"), key_mod, 1))

        sum = self.macro.argument("SUM", temp, t_table_ref, k_table_ref)
        mod = self.macro.argument("MOD", sum, 256)
        temp2 = self.macro._create_variable(-1, -1, temp.name, mod)
        formulas.append(temp2)
        t_table_ref2 = self.macro.argument("OFFSET", t_table[0], temp, 0)
        swap_temp = self.macro._create_variable(-1, -1, random_string(random.randint(10, 25)),
                                        self.macro.argument("EVALUATE", t_table_ref))
        formulas.append(swap_temp)
        formulas.append(self.macro._create_formula(-1,-1, "SET.VALUE", t_table_ref, self.macro.argument("EVALUATE", t_table_ref2)))
        formulas.append(self.macro._create_formula(-1,-1, "SET.VALUE", t_table_ref2, swap_temp))
        formulas.append(self.macro._create_end_loop(-1,-1, "NEXT"))

        self._p1 = self.macro._create_variable(-1, -1, random_string(random.randint(10, 25)), 0)
        self._p2 = self.macro._create_variable(-1, -1, random_string(random.randint(10, 25)), 0)
        formulas.append(self._p1)
        formulas.append(self._p2)
        formulas.append(self.macro._create_formula(-1,-1, "RETURN"))

        return formulas

    def generate(self):
        '''
        Generates RC4 routine formulas and returns them. Note that this function does not add formulas to the worksheet.

        :return: tuple of formulas representing RC4 macro
        '''
        formulas = []
        t_table = self.allocate_init_table()
        init_routine_formulas = self.generate_init_routine(t_table)

        self.routine_ptr = self.macro._create_value(-1, -1,random_string(random.randint(10, 25)))
        formulas.append(self.routine_ptr)
        # Initialize output
        output = self.macro._create_variable(-1, -1,random_string(random.randint(10, 25)), "")
        formulas.append(output)
        # Stores current position in input
        text_ptr = self.macro._create_variable(-1, -1,random_string(random.randint(10, 25)), 1)
        formulas.append(text_ptr)
        formulas.append(self.macro._create_loop(-1,-1,"FOR", text_ptr.name, text_ptr, self.macro.argument("LEN", self.get_argument("input")), 4))

        p1_mod = self.macro.argument("MOD", self.macro.argument("SUM", self._p1, 1), 256)
        t_table_ref3 = self.macro.argument("OFFSET", t_table[0], self._p1, 0)
        p2_mod = self.macro.argument("MOD", self.macro.argument("SUM", self._p2, t_table_ref3), 256)
        t_table_ref4 = self.macro.argument("OFFSET", t_table[0], self._p2, 0)

        p1_1 = self.macro._create_variable(-1, -1,self._p1.name, p1_mod)
        p2_2 = self.macro._create_variable(-1, -1,self._p2.name, p2_mod)
        formulas.append(p1_1)
        formulas.append(p2_2)
        char_to_dec = self.macro.argument("HEX2DEC", self.macro.argument("MID", self.get_argument("input"), text_ptr, 4))

        swap_temp = self.macro._create_variable(-1, -1, random_string(random.randint(10, 25)),
                                        self.macro.argument("EVALUATE", t_table_ref3))
        formulas.append(swap_temp)
        formulas.append(self.macro._create_formula(-1, -1, "SET.VALUE", t_table_ref3, self.macro.argument("EVALUATE", t_table_ref4)))
        formulas.append(self.macro._create_formula(-1, -1, "SET.VALUE", t_table_ref4, swap_temp))

        key = self.macro.argument("OFFSET", t_table[0], self.macro.argument("MOD", self.macro.argument("SUM",
                                                                                                       self.macro.argument(
                                                                                                           "OFFSET",
                                                                                                           t_table[0],
                                                                                                           self._p1, 0),
                                                                                                       self.macro.argument(
                                                                                                           "OFFSET",
                                                                                                           t_table[0],
                                                                                                           self._p2, 0)),
                                                                            256), 0)

        formulas.append(self.macro._create_variable(-1, -1, output.name, self.macro.operator(output, "&", self.macro.argument("UNICHAR",
                                                                                              self.macro.argument(
                                                                                                  "BITXOR", key,
                                                                                                  char_to_dec)))))
        formulas.append(self.macro._create_end_loop(-1,-1,"NEXT"))
        formulas.append(self.macro._create_formula(-1,-1, "RETURN", output))

        return (t_table, init_routine_formulas, formulas)


    def ref_init(self):
        '''
        Return variable that points to initialization routine

        :return: Excel4Variable pointing to initialization routine
        '''
        return self.macro._create_variable(-1, -1, self.initialization_routine_name, self.initialization_routine_ptr)

    def add_ref(self):
        '''
        Adds variables pointing to the beginning of init routine and decrption routine to the worksheet
        '''

        ref = self.ref()
        ref_init = self.ref_init()
        ref.x = self.macro.worksheet._curr_x
        ref.y = self.macro.worksheet._curr_y
        ref_init.x = self.macro.worksheet._curr_x
        ref_init.y = self.macro.worksheet._curr_y + 1
        self.macro._add_to_worksheet(ref)
        self.macro._add_to_worksheet(ref_init)


        return (ref_init, ref)

    def add(self):
        '''
        Adds RC4 routine to the worksheet and returns two Excel4Variables. One that points to initialization routine
        and second that points to decryption routine

        :return: tuple of variables that stores references to initialization and decryption routines
        '''

        if self.is_enabled():
            return self.add_ref()

        t_table, init_routine_formulas, formulas = self.generate()
        # Add formulas to the worksheet
        self.macro.random_add_to_worksheet(t_table)
        self.macro.random_add_to_worksheet(init_routine_formulas)
        self.macro.random_add_to_worksheet(formulas)
        self.set_enabled()
        return self.add_ref()

class Excel4Routines(Excel4MacroExtension):
    '''
    `Excel4Routines` allows to add predefined Excel4 macros to the worksheet such as: RC4 decryption routine.
    `Excel4Routines` stores information about routines, allows to add them and checks if routine is already added to the worksheet.
    `Excel4Routines` object is used by Excel4Macro
    '''

    def __init__(self):
        Excel4MacroExtension.__init__(self)
        self.routines = {}

    def set_macro(self, macro):
        '''
        Sets `macro` and worksheet
        :param macro: Excel4Macro object
        '''
        self.macro = macro
        self.worksheet = macro.worksheet
        self._init_routines()

    def _init_routines(self):
        '''
        Adds default routines
        '''
        self.add(Excel4RC4RoutineStr(self.macro))

    def add(self, routine):
        '''
        Adds routine to Excel4Routines object
        :param routine: Excel4Routine object
        '''
        if routine.name not in self.routines:
            self.routines[routine.name] = routine
        else:
            # Routine with this name already exist
            pass

    def list(self):
        '''
        Returns names of registered routines

        :return: names of registered routines
        '''
        return self.routines.keys()

    def get(self, name):
        '''
        Returns routine or None if routine does not exist

        :param name: name of routine to return

        :return: routine or None if routine does not exist
        '''
        return self.routines.get(name, None)
