import string

class CellReferenceStyle:
    '''
    Excel reference style constants. Represents Excel references styles.
    Used by Cell object.
    '''
    RC_STYLE = 1
    A1_STYLE = 2

class Cell(object):
    '''
    Represents Excel cell.

    This class stores address of cell (`x`,`y`) and `value`.

    - `x` - represents column number;

    - `y` - represents row number.
    '''
    def __init__(self, x, y, value="", tag=""):
        # Columns
        self.x = x
        # Rows
        self.y = y
        # Cell value
        self.value = value
        # Tags are used to group cells
        self.tag = tag
        # Default R1C1 reference style
        self.reference_style = CellReferenceStyle.RC_STYLE
        # Characters representing ROW and COLUMN in RC_STYLE
        self.row_character = "R"
        self.col_character = "C"
        # If true then cell will not be placed in worksheet.
        self.block_add = False


    def get_cell_address(self, row_char=None, col_char=None):
        '''
        Returns cell address with reference style defined by `reference_style` property.

        If `reference_style` is equal to `RC_STYLE` then `row_char` and `col_char` are used as ROW and COLUMN characters.
        If `row_char` or `col_char` are None then `row_character` and `col_character` properties are used as ROW and COLUMN characters.
        By default 'R' and 'C' characters are used to represent row and column.

        :param row_char: character representing ROW

        :param col_char: character representing COLUMN

        :return: string representing cell address
        '''
        if not row_char:
            row_char = self.row_character
        if not col_char:
            col_char = self.col_character

        if self.reference_style == CellReferenceStyle.RC_STYLE:
            return "{}{}{}{}".format(row_char, self.y, col_char, self.x)
        else:
            return "{}{}".format(self.x, self.y)

    def get_address(self):
        '''
        Returns cell address with reference style defined by `reference_style` property.

        :return: string representing cell address
        '''
        return self.get_cell_address()

    def get_column_letter(self):
        '''
        Computes and returns column address in `A1_STYLE`.

        :return: column address in A!_STYLE
        '''
        r = ""
        temp = self.x
        while temp > 0:
            b = (temp - 1)%26
            r = chr(65 + b) + r
            temp = int((temp - b)/26)
        return r

    def __str__(self):
        return self.value

    def __getitem__(self, subscript):
        if isinstance(subscript, slice):
            return str(self)[subscript.start : subscript.stop : subscript.step]
        else:
            return str(self)[subscript]

    def get_length(self):
        '''
        Returns length of cell `value`.

        :return: int representing length of cell `value`
        '''
        return len(str(self))

    def __len__(self):
        return self.get_length()

    def __add__(self, other):
        return str(self) + other

    def __radd__(self, other):
        return other + str(self)