class Excel4Name(object):
    '''
    Stores name of instruction or variable.
    `Excel4Name` is base class of classes that stores names of instruction or variable.
    `Excel4Name` allows to disable obfuscation for specific Excel4Name object.
    '''

    def __init__(self, name):
        self.name = name
        # If False then obfuscation of this object is disabled
        # This flag has the highest priority.
        # The object will not be obfuscated even if obfuscation is enabled in the options
        self._obfuscate = True

    def set_name(self, name):
        '''
        Sets name.

        :param name: string representing name
        '''
        self.name = name

    def __str__(self):
        return self.name

    def __getitem__(self, subscript):
        if isinstance(subscript, slice):
            return str(self)[subscript.start : subscript.stop : subscript.step]
        else:
            return str(self)[subscript]

    def get_length(self):
        '''
        Returns length of the name

        :return: int length of the name
        '''
        return len(str(self))

    def __len__(self):
        return self.get_length()

    def __add__(self, other):
        return str(self) + other

    def __radd__(self, other):
        return other + str(self)

class Excel4InstructionName(Excel4Name):
    '''
    Stores name of instruction.
    `Excel4InstructionName` is used in `Excel4Formula` and `Excel4FormulaArgument` classes to store name of the instruction.
    It allows to disable translation of specific instruction.
    '''
    def __init__(self, name, translate=False):
        Excel4Name.__init__(self, name)
        self.translate = translate



class Excel4VariableName(Excel4Name):
    '''
    Stores name of variable.
    `Excel4VariableName` is used in `Excel4Variable` class to store name of the variable.

    '''
    def __init__(self, name):
        Excel4Name.__init__(self, name)
        self.translate = False