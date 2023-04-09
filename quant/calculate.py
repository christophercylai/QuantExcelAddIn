"""
Perform simple calculations
"""
from typing import List

from .Calculate import calc
from . import global_obj

# pylint: disable=invalid-name


def qxlpyGetCalculate(dub_list: List[float]) -> str:
    """
    return a string pointing to a calc object
    """
    # returns the address of the Calculate py obj
    calobj = calc.Calc(dub_list)
    return global_obj.store_obj(calobj)

def qxlpyCalculateAddNum(addr: str) -> float:
    """
    call calc.add to sum up the numbers
    """
    # this func takes the address returned from Calculate
    # and make add computation
    calobj = global_obj.get_obj(addr)
    return calobj.add()
