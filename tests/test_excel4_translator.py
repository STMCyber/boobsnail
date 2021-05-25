from unittest import TestCase
from excel4lib.lang import *

class TestExcel4Translator(TestCase):
    def test_translate(self):
        Excel4Translator.set_language("pl_PL")
        t = Excel4Translator.translate("GOTO")
        self.assertEqual(t, "PRZEJDŹ.DO", "Should be PRZEJDŹ.DO")
        t = Excel4Translator.translate("GOTO", "en_US")
        self.assertEqual(t, "GOTO", "Should be GOTO")
        Excel4Translator.set_language("en_US")
        t = Excel4Translator.translate("GOTO")
        self.assertEqual(t, "GOTO", "Should be GOTO")
        t = Excel4Translator.translate("GOTO", "pl_PL")
        self.assertEqual(t, "PRZEJDŹ.DO", "Should be PRZEJDŹ.DO")

    def test_get_row_character(self):
        Excel4Translator.set_language("pl_PL")
        t = Excel4Translator.get_row_character()
        self.assertEqual(t, "W", "Should be W")
        t = Excel4Translator.get_row_character("en_US")
        self.assertEqual(t, "R", "Should be R")
        Excel4Translator.set_language("en_US")
        t = Excel4Translator.get_row_character()
        self.assertEqual(t, "R", "Should be R")
        t = Excel4Translator.get_row_character("pl_PL")
        self.assertEqual(t, "W", "Should be W")

    def test_get_col_character(self):
        Excel4Translator.set_language("pl_PL")
        t = Excel4Translator.get_col_character()
        self.assertEqual(t, "K", "Should be K")
        t = Excel4Translator.get_col_character("en_US")
        self.assertEqual(t, "C", "Should be C")
        Excel4Translator.set_language("en_US")
        t = Excel4Translator.get_col_character()
        self.assertEqual(t, "C", "Should be C")
        t = Excel4Translator.get_col_character("pl_PL")
        self.assertEqual(t, "K", "Should be K")
