from unittest import TestCase
from excel4lib.macro import *


class TestExcel4Formula(TestCase):
    def test_translate(self):
        formula = Excel4Formula(1, 1, "GOTO")
        self.assertEqual(str(formula.instruction), "GOTO", "Should be GOTO")
        instruction = formula.get_instruction_translation("pl_PL")
        self.assertEqual(str(instruction), "PRZEJDŹ.DO", "Should be PRZEJDŹ.DO")
        self.assertEqual(str(formula.instruction), "GOTO", "Should be GOTO")

    def test_get_str(self):
        formula = Excel4Formula(1, 1, "GOTO")
        self.assertEqual(str(formula), "=GOTO()", "Should be =GOTO()")
        instruction = formula.get_str("pl_PL")
        self.assertEqual(str(instruction), "=PRZEJDŹ.DO()", "Should be =PRZEJDŹ.DO()")
        self.assertEqual(str(formula.get_str()), "=GOTO()", "Should be =GOTO()")

        jump = Excel4Formula(1, 1, "GOTO")
        goto = Excel4Formula(1, 1, "GOTO", jump)
        self.assertEqual(str(goto), "=GOTO(R1C1)", "Should be: =GOTO(R1C1)")
        self.assertEqual(str(goto.get_str("pl_PL")), "=PRZEJDŹ.DO(W1K1)", "Should be: =PRZEJDŹ.DO(W1K1)")
        self.assertEqual(str(jump), "=GOTO()", "Should be: =GOTO()")
        self.assertEqual(str(jump.get_reference()), "R1C1", "Should be: R1C1")
        self.assertEqual(str(goto.get_str()), "=GOTO(R1C1)", "Should be: =GOTO(R1C1)")


class TestExcel4GoToFormula(TestCase):
    def test__parse_args(self):
        # String
        goto = Excel4GoToFormula(1,1,"GOTO","TEST")
        val = "\"TEST\""
        self.assertEqual(str(goto._parse_args()), val, "Should be: "+val)
        # Numeric
        goto = Excel4GoToFormula(1,1,"GOTO",12)
        val = "12"
        self.assertEqual(str(goto._parse_args()), val, "Should be: "+val)
        # Logical
        goto = Excel4GoToFormula(1,1,"GOTO", Excel4LogicalTest(1,"=",Excel4FormulaArgument("GOTO")))
        val = "1=GOTO()"
        self.assertEqual(str(goto._parse_args()), val, "Should be: " + val)
        val = "1=PRZEJDŹ.DO()"
        self.assertEqual(str(goto._parse_args("pl_PL")), val, "Should be: "+val)
        # Argument
        goto = Excel4GoToFormula(1,1,"GOTO", Excel4FormulaArgument("GOTO"))
        val = "GOTO()"
        self.assertEqual(str(goto._parse_args()), val, "Should be: " + val)
        val = "PRZEJDŹ.DO()"
        self.assertEqual(str(goto._parse_args("pl_PL")), val, "Should be: "+val)
        # Variable
        goto = Excel4GoToFormula(1, 1, "GOTO", Excel4Variable(1,2,"NAME","VALUE"))
        val = "R2C1"
        self.assertEqual(str(goto._parse_args()), val, "Should be: " + val)
        val = "W2K1"
        self.assertEqual(str(goto._parse_args("pl_PL")), val, "Should be: " + val)
        # Cell
        goto = Excel4GoToFormula(1, 1, "GOTO", Excel4Formula(1,2,"GOTO"))
        val = "R2C1"
        self.assertEqual(str(goto._parse_args()), val, "Should be: " + val)
        val = "W2K1"
        self.assertEqual(str(goto._parse_args("pl_PL")), val, "Should be: " + val)
