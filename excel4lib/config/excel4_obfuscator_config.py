

class Excel4ObfuscatorConfig(object):
    '''
    `Excel4ObfuscatorConfig` is used as global configuration class.
     It enables static and programmable change of behavior of `Excel4Obfuscator` objects.
    '''

    # Is enabled
    enable = True

    automatic_obfuscation = True

    # Enable obfuscation of variables names
    obfuscate_variable_names = False
    # Enable obfuscation of registered functions names
    obfuscate_registered_functions = True
    # Enable obfuscation of variables values
    obfuscate_variable_values = True
    # Enable obfuscation of formulas and arguments
    obfuscate_formulas = True

    # Creates cells with randomly generated strings
    generate_noise = True

    # Obfuscation techniques to use during obfuscation process
    obfuscate_int = True
    obfuscate_char = True
    obfuscate_mid = True
    obfuscate_xor = True
    obfuscate_mod = True

    # Try to spread the data across the sheet
    spread_cells = True
    spread_x_min = 1
    spread_x_max = 200
    spread_y_min = 1
    spread_y_max = 140

    # Maximum number of chars in single cell
    cell_limit = 255

    # Do not translate formulas created by obfuscator
    translate = False