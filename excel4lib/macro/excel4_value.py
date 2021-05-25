from .excel4_instruction import *
from excel4lib.sheet.cell import *
from .excel4_argument import *
from excel4lib.utils import *

class Excel4Value(Excel4Instruction):
    '''
    Represents Excel4 cell value (could be string, int, all cells that should store value and have specified x,y cords).

    Excel4Value is responsible for:
    - storing cell value at specified address.
    '''
    def __init__(self, x, y, value):
        '''

        :param x: column number

        :param y: row number

        :param value: value of cell
        '''
        Excel4Instruction.__init__(self, x, y)
        self.tag = random_tag(str(value))
        self.value = value
        self._type = type(value)

        # List of objects in which self is used
        self._references = []

    def _change_reference(self, new_ref):
        '''
        Changes all references to this object to `new_ref`.
        References are changed only in `Excel4Instruction` and `Excel4Argument` objects and their childs.

        :param new_ref: new reference
        '''
        for r in self._references:
            r._replace_reference(self, new_ref)

    def _add_reference(self, ref):
        '''
        Adds object reference to _references.

        :param ref: reference to add
        '''
        self._references.append(ref)

    def is_num(self):
        '''
        Checks if value is numeric `value`

        :return: True if `value` is numeric value and False if not
        '''
        return is_number(self.value)

    def is_str(self):
        '''
        Checks if value is str `value`

        :return: True if `value` is str value and False if not
        '''

        return self._type == str

    def get_str(self, lang=None):
        '''
        Returns `value` as a string

        :param lang: name of the lanuage

        :return: string representing `value`
        '''
        return '{}'.format(self.value)

    def __str__(self):
        '''
        Returns `value` as a string

        :return: string representing `value`
        '''
        return '{}'.format(self.value)

