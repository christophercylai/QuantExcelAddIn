from typing import List

from .Calculate import calc
from . import qxlpy_obj


def GetCalculate(dub_list: List[float]) -> str:
    # returns the address of the Calculate py obj
    c = calc.Calc(dub_list)
    addr = id(c)
    return qxlpy_obj.store_obj(c)

def CalculateAddNum(addr: str) -> float:
    # this func takes the address returned from Calculate
    # and make add computation
    c = qxlpy_obj.get_obj(addr)
    return c.add()
