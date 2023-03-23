# pylint: disable=missing-module-docstring
from .. import calc


def test_calc():
    # pylint: disable=missing-function-docstring
    numlist = [19, 64, 31.8]
    calobj = calc.Calc(numlist)
    ret = 1
    for i in numlist:
        ret *= i
    assert calobj.multiply() == ret
