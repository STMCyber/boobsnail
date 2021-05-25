from .cell import *

class AlreadyReservedException(Exception):
    pass

class CouldNotMoveCellException(Exception):
    pass

class Worksheet(object):
    '''
    Represents Excel Worksheet
    '''
    def __init__(self, name = ""):
        # Name of worksheet
        self.name = name

        # Allocated cells {x:{y:CELL} (...)}
        self.cells = {}

        # Current coordinates
        self._curr_x = 1
        self._curr_y = 1

        # Worksheet boundaries
        self._max_x = 1
        self._max_y = 1

        # Cells sorted by tag name
        self.cells_tag = {}


    def get_column(self, x):
        '''
        Returns all cells in column x
        :param x: number of column
        :return: returns all cells in column x
        '''
        return self.cells.get(x, None)

    def worksheet_iterate(self):
        '''
        Returns cords from 1,1 to self._max_x and self._max_y
        :return: cords in form (x, y)
        '''
        i = 0
        j = 0
        while i < self._max_x:
            j = 0
            while j < self._max_y:
                yield (i+1, j+1)
                j = j + 1
            i = i + 1

    def cell_iterate(self):
        '''
        Iterate through cells
        :return: returns cell
        '''
        i = 0
        j = 0
        while i < self._max_x:
            j = 0
            while j < self._max_y:
                cell = self.get_cell(i + 1, j + 1)
                if cell:
                    yield cell
                j = j + 1
            i = i + 1

    def column_iterate(self):
        '''
        Iterate through columns
        :return: returns all cells in column
        '''
        for c in self.cells.keys():
            yield (c, self.cells[c])

    def get_current_cords(self):
        '''
        Returns current cords
        :return:
        '''
        return (self._curr_x, self._curr_y)

    def set_current_cords(self, x, y):
        '''
        Sets current cords
        :param x:
        :param y:
        :return:
        '''
        self._curr_x = x
        self._curr_y = y

    def get_cell(self, x, y):
        '''
        Returns cell
        :param x: column
        :param y: row
        :return: cell
        '''

        if (x in self.cells) and (y in self.cells[x]):
            return self.cells[x][y]
        return None

    def move_cell(self, cell):
        '''
        Moves cell to next one. If next one is reserved then moves this reserved cell to the next one.
        :param cell:
        :return:
        '''

        next_cell = self.get_cell(cell.x, cell.y + 1)
        if next_cell is not None:
            # Reserved, check if we can move this cell
            #if cell.is_fixed():
            #    raise CouldNotMoveCellException("Cell {} is fixed".format(cell.get_address()))
            self.move_cell(next_cell)

        # Move to the next one
        self.cells[cell.x][cell.y + 1] = cell
        # Remove
        del self.cells[cell.x][cell.y]
        # Overwrite y
        cell.y = cell.y + 1
        self._update_boundaries(cell)

    def is_reserved(self, x, y, height=1):
        '''
        Checks if any of celll from x, y + height is reserved. return True if yes and False if not
        :param x:
        :param y:
        :param height:
        :return:
        '''
        for i in range(0, height):
            if self.get_cell(x, y + i) is not None:
                return True

        return False

    def create_next_cell(self, value="", tag=""):
        '''
        Creates and adds cell to worksheet at current x and y
        :param value: cell value
        :param tag: cell tag
        :return:
        '''
        cell = Cell(self._curr_x, self._curr_y, value, tag)
        return self.add_cell(cell)

    def create_cell(self, x, y, value="", tag=""):
        '''
        Creates and adds cell to worksheet
        :param x: column
        :param y: row
        :param value: cell value
        :param tag: cell tag
        :return:
        '''
        cell = Cell(x, y, value, tag)
        return self.add_cell(cell)

    def add_next_cell(self, cell):
        '''
        Adds next cell to worksheet at current x and y
        :param value: cell value
        :param tag: cell tag
        :return:
        '''
        cell.x = self._curr_x
        cell.y = self._curr_y
        return self.add_cell(cell)

    def add_cell(self, cell):
        '''
        Adds cell to worksheet and moves to the next row
        :param cell: sheet.Cell object
        :return: sheet.Cell object
        '''
        cell = self._add_cell(cell)
        self._move_row()
        return cell

    def replace_cell(self, cell1 , cell2):
        '''
        Replaces cell1 with cell2
        :param cell1:
        :param cell2:
        :return:
        '''
        cell2.x = cell1.x
        cell2.y = cell1.y
        self.cells[cell1.x][cell1.y] = cell2

    def add_above(self, cell, ref):
        '''
        Adds cell above ref. If above cell is reserved then all cells under ref are moved
        :param cell:
        :param ref:
        :return:
        '''
        cords = self.get_current_cords()
        # Check if cell above is reserved
        if (ref.y == 1) or (self.is_reserved(ref.x, ref.y - 1)):
            # Copy cords
            cell.x = ref.x
            cell.y = ref.y
            # Move ref one cell down. What if cell under this one could not be moved?
            self.move_cell(ref)
        else:
            # Set cords
            cell.x = ref.x
            cell.y = ref.y - 1
        # Add cell to the worksheet
        self.add_cell(cell)
        self.set_current_cords(cords[0], cords[1])

    def get_cells_by_tag(self, tag):
        '''
        Return cells by tag
        :param tag:
        :return:
        '''
        return self.cells_tag.get(tag, [])

    def _add_cell_tag(self, cell):
        if cell.tag in self.cells_tag:
            self.cells_tag[cell.tag].append(cell)
        else:
            self.cells_tag[cell.tag] = [cell]

    def _remove_cell_tag(self, cell):
        pass

    def _add_cell(self, cell):

        '''
        Adds cell to worksheet
        :param cell: sheet.Cell object
        :return: sheet.Cell object
        '''
        if cell.block_add:
            return cell

        # Check if cell is not reserved
        if cell.x not in self.cells:
            self.cells[cell.x] = {cell.y : cell}
            self._update_boundaries(cell)
            self._add_cell_tag(cell)
            return cell
        if cell.y not in self.cells[cell.x]:
            self.cells[cell.x][cell.y] = cell
            self._update_boundaries(cell)
            self._add_cell_tag(cell)
            return cell
        raise AlreadyReservedException("Cell {} already reserved".format(cell.get_address()))

    def _update_boundaries(self, cell):
        '''
        Updates worksheet boundaries. When new cell is added this function checks if x and y are in boundaries. If not then x_max and y_max are updated.
        :param cell:
        :return:
        '''
        if cell.x > self._max_x:
            self._max_x = cell.x
        if cell.y > self._max_y:
            self._max_y = cell.y

    def remove_cell(self, cell):
        '''
        Removes cell from the worksheet
        :param cell: cell to remove
        :return: True if cell is delted and Flase if not
        '''
        if (cell.x in self.cells) and (cell.y in self.cells[cell.x]):
            # Remove
            del self.cells[cell.x][cell.y]
            return True
        return False

    def to_csv(self, separator=";"):
        '''
        Dumps all cells to CSV format
        :return: cells as CSV
        '''
        csv = ""
        i = 0
        while i < self._max_y:
            j = 0
            while j < self._max_x:
                cell = self.get_cell(j+1, i+1)
                value = ""
                if cell:
                    value = str(cell)
                csv = csv + value + separator
                j = j + 1
            csv = csv + "\n"
            i = i + 1

        return csv

    def _move_row(self):
        '''
        Moves to the next row
        '''
        self._set_curr_row(self._curr_y + 1)

    def _set_curr_row(self, y):
        '''
        Sets current row to y
        :param y: number of row
        '''
        self._curr_y = y

    def _move_col(self):
        '''
        Moves to the next column
        '''
        self._set_curr_col(self._curr_x + 1)

    def _set_curr_col(self, x):
        '''
        Sets current col to x
        :param x: number of column
        '''
        self._curr_x = x