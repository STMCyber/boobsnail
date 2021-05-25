from unittest import TestCase
from excel4lib.macro import *

class TestExcel4Variable(TestCase):
    def test_get_address(self):
        var = Excel4Variable(1,1, "name", "value")
        c = "{}".format(var.get_address("pl_PL"))
        self.assertEqual(c, "W1K1", "Shuold be W1K1")
        c = "{}".format(var.get_address("en_US"))
        self.assertEqual(c, "R1C1", "Shuold be R1C1")

    def test_get_reference(self):
        var = Excel4Variable(1,1, "name", "value")
        c = "{}".format(var.get_reference("pl_PL"))
        self.assertEqual(c, "name", "Should be name")

    def test__get_value(self):
        # String as value
        var = Excel4Variable(1,1, "name", "value")
        val = var._get_value("pl_PL")
        self.assertEqual(val, '"value"', 'Shuold be "value"')
        # Numeric value
        var = Excel4Variable(1,1, "name", 12121)
        val = var._get_value("pl_PL")
        self.assertEqual(val, "12121", "Shuold be 12121")
        # Cell as value
        var = Excel4Variable(1,1, "name", Excel4Formula(1,2,"GOTO"))
        val = var._get_value("pl_PL")
        self.assertEqual(val, "W2K1", "Shuold be W2K1")
        val = var._get_value()
        self.assertEqual(val, "R2C1", "Shuold be R2C1")
        # Excel4FormulaArgument as value
        var = Excel4Variable(1,1, "name", Excel4FormulaArgument("GOTO"))
        val = var._get_value("pl_PL")
        self.assertEqual(val, "PRZEJDŹ.DO()", "Shuold be PRZEJDŹ.DO()")
        val = var._get_value()
        self.assertEqual(val, "GOTO()", "Shuold be GOTO()")


    def test__get_variable(self):
        var = Excel4Variable(1,1, "name", "value")
        val = var._get_variable()
        self.assertEqual(val, 'name="value"', 'Shuold be name="value"')
