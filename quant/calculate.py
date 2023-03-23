"""
Perform simple calculations
"""
from typing import List

from .Calculate import calc
from . import qxlpy_obj

# pylint: disable=invalid-name


def GetCalculate(dub_list: List[float]) -> str:
    """
    return a string pointing to a calc object
    """
    # returns the address of the Calculate py obj
    calobj = calc.Calc(dub_list)
    return qxlpy_obj.store_obj(calobj)

def CalculateAddNum(addr: str) -> float:
    """
    call calc.add to sum up the numbers
    """
    # this func takes the address returned from Calculate
    # and make add computation
    calobj = qxlpy_obj.get_obj(addr)
    return calobj.add()
