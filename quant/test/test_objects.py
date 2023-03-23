# pylint: disable=missing-module-docstring
from .. import objects as O


def test_objects():
    # pylint: disable=missing-function-docstring
    strdict = {
        "A": "Apple",
        "B": "Bee"
    }
    dic = O.StoreStrDict(strdict)

    strlist = ["a", "b"]
    lst = O.StoreStrList(strlist)

    assert dic in O.ListGlobalObjects()
    assert lst in O.ListGlobalObjects()

    ret = O.GetStrDict(dic)
    assert ret["A"] == strdict["A"]
    assert ret["B"] == strdict["B"]

    O.DeleteObject(dic)
    assert O.ObjectExists(lst) is True
    assert O.ObjectExists(dic) is False
