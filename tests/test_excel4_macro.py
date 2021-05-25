from unittest import TestCase
from excel4lib.macro import  *
from excel4lib.macro.obfuscator import *

class TestExcel4Macro(TestCase):

    def test_to_csv(self):
        macro = Excel4Macro("test.csv")
        macro.worksheet.set_current_cords(2, 1)
        start = macro.value("START")
        macro.formula("ALERT", "WOW!")
        macro.worksheet.set_current_cords(1, 1)
        abcd = macro.variable("abcd", start)
        macro.formula("GOTO",abcd)

        macro.worksheet.set_current_cords(3, 1)
        jump = macro.formula(abcd)
        macro.worksheet.set_current_cords(2, 3)
        macro.condition("IF", Excel4LogicalTest(1,'=',1), Excel4FormulaArgument("GOTO", jump), "")
        csv = macro.to_csv()
        r = """abcd=R1C2;START;=abcd();
=GOTO(abcd);=ALERT("WOW!");;
;=IF(1=1,GOTO(R1C3),"");;
"""
        self.assertEqual(csv, r, "Should be \r\n"+r)

    def test_to_csv_file_spread(self):
        obfuscator =  Excel4Obfuscator()

        macro = Excel4Macro("test.csv", obfuscator)

        alert_val = macro.value(1)
        macro.formula("ALERT", alert_val)
        macro.formula("ALERT", 2)
        alert_val = macro.variable("test", 2)
        macro.formula("ALERT", macro.argument("SUM", alert_val, 1))
        macro.formula("ALERT", macro.argument("SUM", alert_val, 2))
        macro.formula("ALERT", macro.argument("SUM", alert_val, 3))
        alert_val = macro.variable("test", 6)
        macro.formula("ALERT", alert_val)
        macro.formula("ALERT", 7)
        macro.formula("ALERT", 8)
        macro.formula("ALERT", 9)
        macro.obfuscate_all()
        macro.to_csv_file("spread.csv")


    def test__reserve_cells(self):
        pass

    def test__add_to_worksheet(self):
        pass


    def test__create_logical_test(self):
        macro = Excel4Macro("test.csv")
        logical_test = macro._create_logical_test(1, '=', 1)
        logical_result = "1=1"
        self.assertEqual(str(logical_test), logical_result, "Should be {}".format(logical_result))

        logical_test = macro._create_logical_test("a", '>', "b")
        logical_result = '"a">"b"'
        self.assertEqual(str(logical_test), logical_result, "Should be {}".format(logical_result))

        arg1 = macro._create_argument_object("GET.WORKSPACE", 1)
        arg2 = macro._create_argument_object("GET.WORKSPACE", 2)
        logical_test = macro._create_logical_test(arg1, '=', arg2)
        logical_result = "GET.WORKSPACE(1)=GET.WORKSPACE(2)"
        self.assertEqual(str(logical_test), logical_result, "Should be {}".format(logical_result))

        formula1 = macro._create_formula(2,3,"TEST", 1)
        logical_test = macro._create_logical_test(arg1, '<>', formula1)
        logical_result = "GET.WORKSPACE(1)<>R3C2"
        self.assertEqual(str(logical_test), logical_result, "Should be {}".format(logical_result))

    def test__create_argument_object(self):
        macro = Excel4Macro("test.csv")
        arg1 = macro._create_argument_object("GET.WORKSPACE", 1)
        result = "GET.WORKSPACE(1)"
        self.assertEqual(str(arg1), result, "Should be {}".format(result))
        formula1 = macro._create_formula(1,3,"TEST", 1)
        arg1 = macro._create_argument_object("GET.WORKSPACE", 1, formula1)
        result = "GET.WORKSPACE(1,R3C1)"
        self.assertEqual(str(arg1), result, "Should be {}".format(result))

    def test__create_formula(self):
        macro = Excel4Macro("test.csv")
        var1 = macro._create_variable(1,2, "test", "TEST")
        formula1 = macro._create_formula(1,1,"TEST", 1,"a", macro._create_argument_object("GET.WORKSPACE", 1), var1)
        result = '=TEST(1,"a",GET.WORKSPACE(1),test)'
        self.assertEqual(str(formula1), result, "Should be {}".format(result))

    def test__spread_cells(self):
        obfusactor = Excel4Obfuscator()
        for i in range(0, 100):

            macro = Excel4Macro("test.csv", obfusactor)

            for j in range(0, 10):
                cords = macro._gen_random_cords(6)
                macro.worksheet.set_current_cords(cords[0], cords[1])
                try:
                    start = macro.value("START")
                    macro.formula("ALERT", "WOW!")
                    macro.worksheet.set_current_cords(1, 1)
                    abcd = macro.variable("abcd", start)
                    macro.formula("GOTO", abcd)
                    macro.worksheet.set_current_cords(3, 1)
                    jump = macro.formula(abcd)
                    macro.worksheet.set_current_cords(2, 3)
                    macro.condition("IF", Excel4LogicalTest(1, '=', 1), Excel4FormulaArgument("GOTO", jump), "")
                except AlreadyReservedException as ex:
                    pass

            macro._spread_cells()

    def test__generate_noise(self):
        obfusactor = Excel4Obfuscator()
        for i in range(0, 200):

            macro = Excel4Macro("test.csv", obfusactor)

            for j in range(0, 5):
                cords = macro._gen_random_cords(10)
                macro.worksheet.set_current_cords(cords[0], cords[1])
                try:
                    start = macro.value("START")
                    macro.formula("ALERT", "WOW!")
                    abcd = macro.variable("abcd", start)
                    macro.formula("GOTO", abcd)
                    macro.formula("GOTO", abcd)
                    macro.formula("GOTO", abcd)
                    macro.formula("GOTO", abcd)
                    macro.formula("GOTO", abcd)
                    macro.formula("GOTO", abcd)
                    macro.formula("GOTO", abcd)
                    jump = macro.formula(abcd)
                    macro.condition("IF", Excel4LogicalTest(1, '=', 1), Excel4FormulaArgument("GOTO", jump), "")
                except AlreadyReservedException as ex:
                    pass

            macro._generate_noise()

    def test__obfuscate_function_names(self):
        obfusactor = Excel4Obfuscator()
        for i in range(0, 10):

            macro = Excel4Macro("test.csv", obfusactor)

            for j in range(0, 5):
                cords = macro._gen_random_cords(6)
                macro.worksheet.set_current_cords(cords[0], cords[1])
                try:
                    write = macro.register_write_process_memory("WRITE")
                    create = macro.register_create_thread("CREATE")
                    macro.formula(write,1,2,3,4)
                    macro.argument(create.get_function_text(), 1,2,3)
                    macro.formula(write.get_function_text(),2,3,4,5)
                    macro.variable("test", macro.argument(create.get_function_text(), 1,2,3))
                except AlreadyReservedException as ex:
                    pass

            macro._obfuscate_function_names()

    def test__obfuscate_variable_names(self):
        obfusactor = Excel4Obfuscator()
        for i in range(0, 10):

            macro = Excel4Macro("test.csv", obfusactor)

            for j in range(0, 5):
                cords = macro._gen_random_cords(6)
                macro.worksheet.set_current_cords(cords[0], cords[1])
                try:
                    start = macro.value("START")
                    macro.variable("urt", start)
                    macro.variable("sacas", "asdsaddsadsa")
                    macro.variable("dsadsa", 1111)
                    abcd = macro.variable("abcd", start)
                    jump = macro.formula(abcd)
                    macro.variable("dsadsadsacsa", Excel4FormulaArgument("GOTO", jump))
                    macro.variable("xdqxdwqfe", Excel4FormulaArgument("GOTO", Excel4LogicalTest(1, '=', 1)))
                    macro.formula("GOTO", abcd)
                    macro.argument("TEST", abcd)
                    macro.condition("IF", macro.argument("TEST", abcd), macro.logical(abcd, "<>", "test"))
                    macro.loop("FOR", abcd.name, abcd, 1,1)
                except AlreadyReservedException as ex:
                    pass

            macro._obfuscate_variable_names()

    def test__obfuscate_variable_values(self):
        obfusactor = Excel4Obfuscator()
        for i in range(0, 10):

            macro = Excel4Macro("test.csv", obfusactor)

            for j in range(0, 5):
                cords = macro._gen_random_cords(14)
                macro.worksheet.set_current_cords(cords[0], cords[1])
                try:
                    start = macro.value("START")
                    abcd = macro.variable("abcd", start)
                    macro.formula("GOTO", abcd)
                    jump = macro.formula(abcd)
                    macro.formula("ALERT", "WOW!")
                    macro.variable("urt", start)
                    macro.variable("sacas", "asdsaddsadsa")
                    macro.variable("dsadsa", 1111)
                    macro.variable("dsadsadsacsa", Excel4FormulaArgument("GOTO", jump))
                    macro.variable("xdqxdwqfe", Excel4FormulaArgument("GOTO", Excel4LogicalTest(1, '=', 1)))
                    abcd = macro.variable("abcd", start)
                    macro.formula("GOTO", abcd)
                    jump = macro.formula(abcd)
                    macro.condition("IF", Excel4LogicalTest(1, '=', 1), Excel4FormulaArgument("GOTO", jump), "")
                except AlreadyReservedException as ex:
                    pass

            macro._obfuscate_variable_values()

    def test__obfuscate_formulas(self):
        obfusactor = Excel4Obfuscator()
        for i in range(0, 10):

            macro = Excel4Macro("test.csv", obfusactor)

            for j in range(0, 5):
                cords = macro._gen_random_cords(6)
                macro.worksheet.set_current_cords(cords[0], cords[1])
                try:
                    start = macro.value("START")
                    macro.formula("ALERT", "WOW!")
                    abcd = macro.variable("abcd", start)
                    macro.formula("GOTO", abcd)
                    jump = macro.formula(abcd)
                    macro.formula(abcd)
                    macro.formula("TEST", 123,"test", abcd)
                    macro.formula("AAAA", Excel4FormulaArgument("GOTO", jump), Excel4LogicalTest(1, '=', 1), start)
                    create_thread = macro.register_create_thread("TEST")
                    macro.formula(create_thread, 3,4,1)
                    macro.condition("IF", Excel4LogicalTest(1, '=', 1), Excel4FormulaArgument("GOTO", jump), "")
                except AlreadyReservedException as ex:
                    pass

            macro._obfuscate_formulas()

    def test__obfuscate_all(self):
        obfusactor = Excel4Obfuscator()
        for i in range(0, 200):
            obfusactor = Excel4Obfuscator()
            for i in range(0, 200):

                macro = Excel4Macro("test.csv", obfusactor)

                for j in range(0, 5):
                    cords = macro._gen_random_cords(6)
                    macro.worksheet.set_current_cords(cords[0], cords[1])
                    try:
                        start = macro.value("START")
                        macro.formula("ALERT", "WOW!")
                        macro.worksheet.set_current_cords(1, 1)
                        abcd = macro.variable("abcd", start)
                        macro.formula("GOTO", abcd)
                        macro.worksheet.set_current_cords(3, 1)
                        jump = macro.formula(abcd)
                        macro.worksheet.set_current_cords(2, 3)
                        macro.condition("IF", Excel4LogicalTest(1, '=', 1), Excel4FormulaArgument("GOTO", jump), "")
                    except AlreadyReservedException as ex:
                        pass
            macro.obfuscate_all()