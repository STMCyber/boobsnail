

class Excel4MacroExtension(object):
    '''
    Basic class for Excel4Macro extensions.
    Extensions allows to extend `Excel4Macro` functionality.
    '''
    def __init__(self):
        # Macro object
        self.macro = None
        # Worksheet object stored in macro
        self.worksheet = None

    def set_macro(self, macro):
        '''
        Sets macro and worksheet in extension

        :param macro: Object of Excel4Macro
        '''
        self.macro = macro
        self.worksheet = macro.worksheet
