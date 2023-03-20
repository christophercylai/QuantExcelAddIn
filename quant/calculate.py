from .Calculate import calc
from . import qxlpy_obj


def GetCalculate(dub_list: list) -> str:
    c = calc.Calc(dub_list)
    addr = id(c)
    return qxlpy_obj.store_obj(c)

def CalculateAddNum(addr: str) -> float:
    c = qxlpy_obj.get_obj(addr)
    return c.add()
