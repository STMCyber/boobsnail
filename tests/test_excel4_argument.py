from unittest import TestCase
from excel4lib.macro import *

class TestExcel4FormulaArgument(TestCase):
    def test_get_instruction_translation(self):
        formula = Excel4FormulaArgument("GOTO")
        self.assertEqual(str(formula.instruction), "GOTO", "Should be GOTO")
        instruction = formula.get_instruction_translation("pl_PL")
        self.assertEqual(str(instruction), "PRZEJDŹ.DO", "Should be PRZEJDŹ.DO")
        self.assertEqual(str(formula.instruction), "GOTO", "Should be GOTO")

    def test_revert_translation(self):
        formula = Excel4FormulaArgument("GOTO")
        self.assertEqual(str(formula.instruction), "GOTO", "Should be GOTO")
        formula.set_language("pl_PL")
        self.assertEqual(str(formula), "PRZEJDŹ.DO()", "Should be PRZEJDŹ.DO()")
        formula.revert_translation()
        self.assertEqual(str(formula), "GOTO()", "Should be GOTO()")

    def test__get_func(self):
        # Numeric argument
        val = 1
        formula = Excel4FormulaArgument("GOTO", val)
        self.assertEqual(formula._get_func(), "GOTO(1)", "Should be GOTO(1)")
        self.assertEqual(formula._get_func("pl_PL"), "PRZEJDŹ.DO(1)", "Should be PRZEJDŹ.DO(1)")

        # String argument
        val = "TEST"
        formula = Excel4FormulaArgument("GOTO", val)
        self.assertEqual(formula._get_func(), "GOTO(\"TEST\")", "Should be GOTO(\"TEST\")")
        self.assertEqual(formula._get_func("pl_PL"), "PRZEJDŹ.DO(\"TEST\")", "Should be PRZEJDŹ.DO(\"TEST\")")

        # Cell argument
        val = Excel4Value(1,1,"test")
        formula = Excel4FormulaArgument("GOTO", val)
        self.assertEqual(formula._get_func(), "GOTO(R1C1)", "Should be GOTO(R1C1)")
        self.assertEqual(formula._get_func("pl_PL"), "PRZEJDŹ.DO(W1K1)", "Should be PRZEJDŹ.DO(W1K1)")

        # Variable argument
        val = Excel4Variable(1, 1, "test", "value")
        formula = Excel4FormulaArgument("GOTO", val)
        self.assertEqual(formula._get_func(), "GOTO(test)", "Should be GOTO(test)")
        self.assertEqual(formula._get_func("pl_PL"), "PRZEJDŹ.DO(test)", "Should be PRZEJDŹ.DO(test)")

        # Logical argument
        val = Excel4LogicalTest("val1", "=", "val2")
        formula = Excel4FormulaArgument("GOTO", val)
        self.assertEqual(formula._get_func(), "GOTO(\"val1\"=\"val2\")", "Should be GOTO(\"val1\"=\"val2\")")
        self.assertEqual(formula._get_func("pl_PL"), "PRZEJDŹ.DO(\"val1\"=\"val2\")", "Should be PRZEJDŹ.DO(\"val1\"=\"val2\")")
        # MIX
        val = Excel4LogicalTest("val1", "=", "val2")
        formula = Excel4FormulaArgument("GOTO", 1, "test", Excel4Value(1,1,"test"), val)
        self.assertEqual(formula._get_func(), "GOTO(1,\"test\",R1C1,\"val1\"=\"val2\")", "Should be GOTO(1,\"test\",R1C1,\"val1\"=\"val2\")")
        self.assertEqual(formula._get_func("pl_PL"), "PRZEJDŹ.DO(1,\"test\",W1K1,\"val1\"=\"val2\")", "Should be PRZEJDŹ.DO(1,\"test\",W1K1,\"val1\"=\"val2\")")

