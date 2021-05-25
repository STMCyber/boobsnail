from unittest import TestCase
from excel4lib.macro.obfuscator import *

class TestExcel4Obfuscator(TestCase):
    def test_char(self):
        obfuscator = Excel4Obfuscator()
        formula = obfuscator.char("A")
        self.assertEqual(str(formula), 'CHAR(65)', "Should be CHAR(65)")

    def test_int(self):
        obfuscator = Excel4Obfuscator()
        formula = obfuscator.int("A")
        self.assertEqual(str(formula), 'CHAR(INT("65"))', 'Should be CHAR(INT("65"))')
        self.assertEqual(formula.get_length(), 15, 'Should be 15')

