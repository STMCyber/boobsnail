
from excel4lib.generator import *
import sys

generators = [Excel4NtDonutGenerator, Excel4ExecGenerator, Excel4DownloadExecuteGenerator]

banner = """
___.                ___.     _________             .__.__   
\_ |__   ____   ____\_ |__  /   _____/ ____ _____  |__|  |  
 | __ \ /  _ \ /  _ \| __ \ \_____  \ /    \\__  \ |  |  |  
 | \_\ (  <_> |  <_> ) \_\ \/        \   |  \/ __ \|  |  |__
 |___  /\____/ \____/|___  /_______  /___|  (____  /__|____/
     \/                  \/        \/     \/     \/         
     Author: @_mzer0 @stm_cyber
     """


def print_usage():
    print(banner)
    print("Usage: {} <generator>".format(sys.argv[0]))
    print("Generators:")
    for g in generators:
        print("{} - {}".format(g.name, g.description))

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print_usage()
        sys.exit(1)

    generator = sys.argv[1]
    for g in generators:
        if g.name.lower() == generator.lower():
            Boobsnail.print_banner()
            g.run(sys.argv[2:])
            sys.exit(1)
    else:
        print("Unknown generator {}!".format(generator))
        print_usage()
        sys.exit(1)