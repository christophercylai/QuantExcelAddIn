from .. import calculate as C


def test_calculate():
    numlist = [10, 20, 31.8]
    ret = 0
    for i in numlist:
        ret += i
    c = C.GetCalculate(numlist)
    assert C.CalculateAddNum(c) == ret
