from typing import List

from quant import py_logger


class Calc:
    def __init__(self, numlst: List(float)):
        for n in numlst:
            if not isinstance(n, float) and not isinstance(n, int):
                err = "Non-float value was supplied in numlist"
                py_logger.error(err)
                raise TypeError(err)
        self.numlist = numlst

    def add(self) -> float:
        summation = 0
        for ea_num in self.numlist:
            summation += ea_num
        return summation

    def multiply(self) -> float:
        product = 1
        for ea_num in self.numlist:
            product *= ea_num
        return product
