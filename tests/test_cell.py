from unittest import TestCase
import unittest
from excel4lib.sheet.cell import Cell

class TestCell(TestCase):
    def test__get_column_letter(self):
        self.assertEqual(Cell(677,1).get_column_letter(), "ZA", "Should be ZA")
        self.assertEqual(Cell(702, 1).get_column_letter(), "ZZ", "Should be ZZ")
        self.assertEqual(Cell(573, 1).get_column_letter(), "VA", "Should be VA")
        self.assertEqual(Cell(584, 1).get_column_letter(), "VL", "Should be VL")
        self.assertEqual(Cell(17, 1).get_column_letter(), "Q", "Should be Q")
        self.assertEqual(Cell(703, 1).get_column_letter(), "AAA", "Should be AAA")



if __name__ == '__main__':
    unittest.main()