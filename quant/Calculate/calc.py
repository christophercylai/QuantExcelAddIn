"""
Calculate numbers
"""
from typing import List

from quant import py_logger


class Calc:
    """
    Calc object that store a list of numbers for computation
    """
    def __init__(self, numlst: List[float]):
        for ea_num in numlst:
            if not isinstance(ea_num, float) and not isinstance(ea_num, int):
                err = "Non-float value was supplied in numlist"
                py_logger.error(err)
                raise TypeError(err)
        self.numlist = numlst

    def add(self) -> float:
        """
        sum up the list the numbers
        """
        summation = 0
        for ea_num in self.numlist:
            summation += ea_num
        return summation

    def multiply(self) -> float:
        """
        return a product of the list of numbers
        """
        product = 1
        for ea_num in self.numlist:
            product *= ea_num
        return product
