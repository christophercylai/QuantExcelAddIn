from .Calculate import calc
import quant


def GetCalculate(numlst: list = [1, 2, 3]) -> str:
    c = calc.Calc(numlst)
    addr = id(c)
    return quant.STORE_OBJ(c)

def CalculateAddNum(addr: str) -> float:
    c = quant.GET_OBJ(addr)
    return c.add()
