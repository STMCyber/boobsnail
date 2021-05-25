from unittest import TestCase
from excel4lib.sheet import *

class TestWorksheet(TestCase):

    def test_column_iterate(self):
        worksheet = Worksheet("test.csv")
        worksheet.set_current_cords(1, 1)
        for i in range(1,10):
            for j in range(1, 10):
                worksheet.add_cell(Cell(i, j, "{}{}".format(i,j)))
        i = 1
        for col in worksheet.column_iterate():
            j = 1

            for c in col[1]:
                self.assertEqual(str(col[1][c]), "{}{}".format(i,j), "Should be {}{}".format(i,j))
                j = j + 1
            i = i + 1

    def test_get_cell(self):
        worksheet = Worksheet("test.csv")
        worksheet.set_current_cords(1, 1)
        worksheet.add_next_cell(Cell(-1, -1, "A"))
        worksheet.add_next_cell(Cell(-1, -1, "A"))
        worksheet.add_next_cell(Cell(-1, -1, ""))
        worksheet.add_next_cell(Cell(-1, -1, ""))
        worksheet.add_next_cell(Cell(-1, -1, "A"))
        worksheet.set_current_cords(2, 1)
        worksheet.add_next_cell(Cell(-1, -1, "B"))
        worksheet.add_next_cell(Cell(-1, -1, "B"))

        cell = worksheet.get_cell(1,1)
        self.assertEqual(str(cell), "A", "Should be: A")
        cell = worksheet.get_cell(10, 1)
        self.assertEqual(cell, None, "Should be: None")

    def test_is_reserved(self):
        worksheet = Worksheet("test.csv")
        for i in range(1,5):
            worksheet.add_cell(Cell(1,i))
        self.assertEqual(worksheet.is_reserved(1, 1, 2), True, "Should be True")
        self.assertEqual(worksheet.is_reserved(2, 1, 2), False, "Should be False")
        for i in range(8,12):
            worksheet.add_cell(Cell(1,i))
        self.assertEqual(worksheet.is_reserved(6, 8, 2), False, "Should be False")

    def test_add_next_cell(self):
        worksheet = Worksheet("test.csv")
        worksheet.set_current_cords(1,1)
        worksheet.add_next_cell(Cell(-1,-1,"A"))
        worksheet.add_next_cell(Cell(-1, -1, "A"))
        worksheet.add_next_cell(Cell(-1, -1, ""))
        worksheet.add_next_cell(Cell(-1, -1, ""))
        worksheet.add_next_cell(Cell(-1, -1, "A"))
        worksheet.set_current_cords(2, 1)
        worksheet.add_next_cell(Cell(-1, -1, "B"))
        worksheet.add_next_cell(Cell(-1, -1, "B"))
        csv = worksheet.to_csv()
        val = """A;B;\nA;B;\n;;\n;;\nA;;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))
    def test_add_cell(self):
        worksheet = Worksheet("test.csv")
        worksheet.add_cell(Cell(1,1, "A"))
        worksheet.add_cell(Cell(2,1, "B"))
        worksheet.add_cell(Cell(1,2, "A"))
        worksheet.add_cell(Cell(2,2, "B"))
        worksheet.add_cell(Cell(1,5, "A"))
        csv = worksheet.to_csv()
        val = """A;B;\nA;B;\n;;\n;;\nA;;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))

    def test_replace_cell(self):
        worksheet = Worksheet("test.csv")
        worksheet.add_cell(Cell(1,1, "A"))
        worksheet.add_cell(Cell(2,1, "B"))
        worksheet.add_cell(Cell(1,2, "A"))
        worksheet.add_cell(Cell(2,2, "B"))

        c = Cell(1,5, "A")
        c2 = Cell(1,5, "C")
        worksheet.add_cell(c)
        worksheet.replace_cell(c, c2)
        csv = worksheet.to_csv()
        val = """A;B;\nA;B;\n;;\n;;\nC;;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))

    def test_add_above(self):
        worksheet = Worksheet("test.csv")
        # Cell is in first row
        c = Cell(1, 1, "A")
        worksheet.add_cell(c)
        worksheet.add_above(Cell(1,1, "B"), c)
        csv = worksheet.to_csv()
        val = """B;\nA;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))

        # Cell above is empty
        worksheet = Worksheet("test.csv")
        c = Cell(1, 2, "A")
        worksheet.add_cell(c)
        worksheet.add_above(Cell(1,1, "B"), c)
        csv = worksheet.to_csv()
        val = """B;\nA;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))

        # Cell above is reserved but below is not
        worksheet = Worksheet("test.csv")
        c = Cell(1, 2, "A")
        worksheet.add_cell(c)
        worksheet.add_cell(Cell(1, 1, "A"))
        worksheet.add_above(Cell(1,2, "B"), c)
        csv = worksheet.to_csv()
        val = """A;\nB;\nA;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))
        # Cell above and below are reserved
        worksheet = Worksheet("test.csv")
        c = Cell(1, 2, "A")
        worksheet.add_cell(c)
        worksheet.add_cell(Cell(1, 1, "A"))
        worksheet.add_cell(Cell(1, 3, "A"))
        worksheet.add_above(Cell(1,2, "B"), c)
        csv = worksheet.to_csv()
        val = """A;\nB;\nA;\nA;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))

        # Cell above and below are reserved
        worksheet = Worksheet("test.csv")
        c = Cell(1, 2, "A")
        worksheet.add_cell(c)
        worksheet.add_cell(Cell(1, 1, "A"))
        worksheet.add_cell(Cell(1, 3, "A"))
        worksheet.add_cell(Cell(1, 4, "A"))
        worksheet.add_cell(Cell(1, 6, "A"))
        worksheet.add_above(Cell(1,2, "B"), c)
        csv = worksheet.to_csv()
        val = """A;\nB;\nA;\nA;\nA;\nA;\n"""
        self.assertEqual(csv, val, "Should be: {}".format(val))

    def test_remove_cell(self):
        worksheet = Worksheet("test.csv")
        c = Cell(1, 1, "A")
        worksheet.add_cell(c)
        c = worksheet.get_cell(1,1)
        self.assertEqual(str(c), "A", "Should be A")
        worksheet.remove_cell(c)

        c = worksheet.get_cell(1, 1)
        self.assertEqual(c, None, "Should be None")
