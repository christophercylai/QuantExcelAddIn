from .. import calc


def test_calc():
    numlist = [19, 64, 31.8]
    C = calc.Calc(numlist)
    ret = 1
    for i in numlist:
        ret *= i
    assert C.multiply() == ret
