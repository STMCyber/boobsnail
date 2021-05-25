from .excel4_instruction import *


class Excel4Result(Excel4Instruction):
    '''
    Represents cell in which result of another call will be placed (for example Formula((...), R1C2)
    R1C2 is address of cell in which FORMULA call will save return value in this case R1C2 should be represented as Excel4Result object)
    '''
    def __init__(self, x, y):
        Excel4Instruction.__init__(self, x, y)

    def __str__(self):
        return ""

class Excel4ResultLoop(Excel4Result):
    pass

class Excel4ResultCondition(Excel4Result):
    pass

class Excel4ResultEndLoop(Excel4Result):
    pass

