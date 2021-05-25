from excel4lib.macro import *
from excel4lib.macro.obfuscator import *
import argparse
import sys

class Excel4Generator(object):
    '''
    `Excel4Generator` is a base class for Excel 4.0 Macro generators used in boobsnail package.
    '''



    def __init__(self, worksheet_name = "Excel4Generator", obfuscator = None, routines=None, desc=""):
        # Obfuscator object
        self.obfuscator = obfuscator
        # Routines object
        self.routines = routines
        # Macro object
        # For now generator support only one macro.
        self.macro = None
        self.worksheet_name = worksheet_name
        # CLI args
        self.args_parser = argparse.ArgumentParser(description=desc)
        self.args = None
        self.description = desc
        self.obfuscators = [Excel4Obfuscator, Excel4Rc4Obfuscator]

    def set_obfuscator(self, obfuscator):
        '''
        Sets obfuscator used by generator.

        :param obfuscator: string representing name of obfuscator
        '''
        for o in self.obfuscators:
            if o.name.lower() == obfuscator.lower():
                obfuscator = o
                break
        else:
            return
        self.obfuscator = obfuscator()

    def list_obfuscators(self):
        '''

        :return:
        '''
        for o in self.obfuscators:
            print("{} - {}".format(o.name, o.description))

    def init_macro(self):
        '''
        Initializes macro object
        '''
        self.macro = Excel4Macro(self.worksheet_name, obfuscator=self.obfuscator, routines=self.routines)

        self.macro.set_trigger_cords(self.args.x_auto, self.args.y_auto)
        self.macro.set_cords(self.args.x_auto, self.args.y_auto)

    def init_args(self):
        '''
        Adds command line arguments to ArgumentParser
        '''
        self.args_parser.add_argument("--obfuscator", "-ob", required=False, choices=[o.name for o in self.obfuscators], default="standard",
                                      help="obfuscator to use during obfuscation process")
        self.args_parser.add_argument("--list-obfuscators", "-lo", action="store_true", required=False, help="lists available obfuscators")
        self.args_parser.add_argument("--out", "-o", required=False, default="boobsnail.csv", help="name of output file")
        self.args_parser.add_argument("--disable-obfuscation", "-do", action="store_true", required=False,
                                      help="disable obfuscation")
        self.args_parser.add_argument("--x_auto", "-x", required=False, type=int, default=1,
                                      help="Auto_Open or Auto_Close cell column")
        self.args_parser.add_argument("--y_auto", "-y", required=False, type=int, default=1,
                                      help="Auto_Open or Auto_Close cell row")
        self.args_parser.add_argument("--language", "-l", required=False, choices=["en_US"] + Excel4Translator.get_languages(), default="en_US",
                                      help="language into witch formulas are translated")

        self.custom_args()

    def custom_args(self):
        '''
        Adds custom command line arguments to ArgumentParser
        '''
        pass

    def parse(self, args):
        '''
        Parses command line arguments.

        :param args: cli arguments
        '''
        self.args = self.args_parser.parse_args(args)

    def generate(self):
        '''
        Generates Excel4 Macro
        '''
        if self.args.list_obfuscators:
            self.list_obfuscators()
            sys.exit(1)
        if self.args.obfuscator:
            self.set_obfuscator(self.args.obfuscator)
        self.init_macro()
        if self.args.language:
            Excel4Translator.set_language(self.args.language)
        if not self.args.disable_obfuscation:
            Excel4Config.obfuscator.enable = True
        else:
            Excel4Config.obfuscator.enable = False

        self.generate_macro()

    def generate_macro(self):
        '''
        Generates Excel4 Macro
        '''
        pass

    @staticmethod
    def run(args):
        '''
        Runs generator with `args` as arguments from command line.

        :param args: cli arguments
        '''
        generator = Excel4Generator()
        generator.init_args()
        generator.parse(args)
        generator.generate()