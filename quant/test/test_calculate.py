# pylint: disable=missing-module-docstring
from .. import calculate as C


def test_calculate():
    # pylint: disable=missing-function-docstring
    numlist = [10, 20, 31.8]
    ret = 0
    for i in numlist:
        ret += i
    cal = C.qxlpyGetCalculate(numlist, "test_calc")
    assert C.qxlpyCalculateAddNum(cal) == ret
