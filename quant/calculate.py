from .Calculate import calculate
import quant


def GetCalculate(numlst: list = [1, 2, 3]) -> str:
    c = calculate.Calc(numlst)
    addr = id(c)
    return quant.STORE_OBJ(c)

def CalculateAddNum(addr: str) -> float:
    c = quant.GET_OBJ(addr)
    return c.add()
