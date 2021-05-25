from unittest import TestCase
from excel4lib.macro import *
from excel4lib.lang import *

class TestExcel4Instruction(TestCase):

    def test_excel4instruction_str(self):
        formula = Excel4Formula(1, 2, "test", "test", 1, 2, 3)
        self.assertEqual(str(formula), '=test("test",1,2,3)', 'Should be =test("test",1,2,3)')

    def test_excel4instruction_args(self):
        argument = Excel4FormulaArgument("ISNUMBER", 1)
        formula = Excel4Formula(1, 2, "test", "test", 1, 2, argument)
        self.assertEqual(str(formula), '=test("test",1,2,ISNUMBER(1))', 'Should be =test("test",1,2,ISNUMBER(1))')

    def test_excel4instruction_if(self):
        argument = Excel4FormulaArgument("ISNUMBER", 1)
        formula = Excel4ConditionFormula(1, 1, "IF", Excel4LogicalTest(argument, "=", 1),
                                         Excel4Formula(1, 2, "GOTO", "A"), "B")
        self.assertEqual(str(formula), '=IF(ISNUMBER(1)=1,R2C1,"B")', 'Should be =IF(ISNUMBER(1)=1,R2C1,"B")')

    def test_get_reference(self):
        instruction = Excel4Instruction(1,1)
        c = "{}".format(instruction.get_reference("pl_PL"))
        self.assertEqual(c, "W1K1", "Shuold be W1K1")
        c = "{}".format(instruction.get_reference("en_US"))
        self.assertEqual(c, "R1C1", "Shuold be R1C1")

    def test_translate_address(self):
        instruction = Excel4Instruction(1,1)
        instruction.set_language("pl_PL")
        c = "{}".format(instruction.get_address())
        self.assertEqual(c, "W1K1", "Shuold be W1K1")
        instruction.set_language("en_US")
        c = "{}".format(instruction.get_address())
        self.assertEqual(c, "R1C1", "Shuold be R1C1")

    def test_revert_address_translation(self):
        Excel4Translator.native_language = "en_US"
        instruction = Excel4Instruction(1,1)
        instruction.set_language("pl_PL")
        c = "{}".format(instruction.get_address())
        self.assertEqual(c, "W1K1", "Shuold be W1K1")
        instruction.revert_address_translation()
        c = "{}".format(instruction.get_address())
        self.assertEqual(c, "R1C1", "Shuold be R1C1")
