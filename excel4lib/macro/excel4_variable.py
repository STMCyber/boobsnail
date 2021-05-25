from .excel4_value import *
from .excel4_argument import *

from excel4lib.sheet.cell import *
from excel4lib.exception import *



class Excel4Variable(Excel4Value):
    '''
    Represents Excel4 variable (all cells that store variable definition in form name=value, where value could be anything)
    '''
    def __init__(self, x, y, name, value):
        Excel4Value.__init__(self, x, y, value)
        if is_number(name):
            raise Excel4VariableWrongNameException("Variable name could not be numeric")
        # Name of variable. In Excel4 you can use this name as reference
        if not issubclass(type(name), Excel4Name):
            self.name = Excel4VariableName(name)
        else:
            self.name = name
        # List of objects in which self is used
        self._references = []
        try:
            value._add_reference(self)
        except:
            pass

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

    def _replace_reference(self, ref, new_ref):
        '''
        Replaces references `value` with `new_ref`

        :param ref: old reference

        :param new_ref: new reference
        '''
        if self.value == ref:
            self.value = new_ref

    def get_address(self, lang=None):
        '''
        Returns address of cell in which variable is placed. Address could be returned in specific language passed in `lang` parameter.
        If `lang` is None then language from `language` property is taken.

        :param lang: name of the language

        :return: address of cell
        '''
        t_r = self.row_character
        t_c = self.col_character
        if lang and (lang != self.language):
            t_r = Excel4Translator.get_row_character(lang)
            t_c = Excel4Translator.get_col_character(lang)
        r = self.get_cell_address(t_r, t_c)
        return r

    def get_reference(self, lang=None):
        '''
        Returns name of variable

        :return: name of variable as a string
        '''
        return self.name

    def set_name(self, name):
        '''
        Sets name of the variable.

        :param name: variable name
        '''
        self.name.set_name(name)

    def get_name(self):
        '''
        Returns name of the variable

        :return: name of the variable
        '''

        return self.name

    def _escape_value(self):
        '''
        Escapes quotes in `value`

        :return: escaped value
        '''
        return self.value.replace('"', '""')

    def _get_value(self, lang=None):
        '''
        Returns Excel4 variable definition. If `lang` is None then language from `language` property is taken.

        :param lang: language in which arguments should be returned

        :return: string representing variable value
        '''
        if not lang:
            lang = self.language

        value_temp = self.value

        if issubclass(type(value_temp), Excel4FormulaArgument):
            value_temp = value_temp.get_definition(lang)
        elif issubclass(type(value_temp), Excel4LogicalTest):
            value_temp = value_temp._get_argument(lang)
        elif issubclass(type(value_temp), Excel4Value):
            value_temp = value_temp.get_reference(lang)
        elif issubclass(type(value_temp), Cell):
            value_temp = value_temp.get_reference(lang)
        elif is_number(value_temp):
            value_temp = value_temp
        elif type(value_temp) == str:
            value_temp = self._escape_value()
            value_temp = '"' + str(value_temp) + '"'
        else:
            value_temp = str(self.value)

        val = "{VALUE}".format(
            **{
                "VALUE": value_temp
            }
        )
        return val

    def _get_variable(self, lang=None):
        '''
        Returns Excel4 variable definition. If `lang` is None then language from `language` property is taken.

        :param lang: language in which arguments should be returned

        :return: string representing variable
        '''
        call = "{NAME}={VALUE}".format(
            **{
                "NAME": self.name,
                "VALUE": self._get_value(lang)
            }
        )
        return call

    def get_str(self, lang=None):
        '''
        Returns Excel4 variable definition. If `lang` is None then language from `language` property is taken.

        :param lang: language in which arguments should be returned

        :return: string representing variable
        '''
        return self._get_variable(lang)

    def __str__(self):
        '''
        Returns Excel4 variable definition.

        :return: string representing variable
        '''
        return self.get_str(self.language)
