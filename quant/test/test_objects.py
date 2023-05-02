# pylint: disable=missing-module-docstring
from .. import objects as O


def test_objects():
    # pylint: disable=missing-function-docstring
    strdict = {
        "A": "Apple",
        "B": "Bee"
    }
    prefix = "test_obj_dict"
    dic = O.qxlpyStoreStrDict(strdict, prefix)
    assert dic.startswith(prefix)

    strlist = ["a", "b"]
    lst = O.qxlpyStoreStrList(strlist, "test_obj_list")

    assert dic in O.qxlpyListGlobalObjects()
    assert lst in O.qxlpyListGlobalObjects()

    ret = O.qxlpyGetStrDict(dic)
    assert ret["A"] == strdict["A"]
    assert ret["B"] == strdict["B"]

    O.qxlpyDeleteObject(dic)
    assert O.qxlpyObjectExists(lst) is True
    assert O.qxlpyObjectExists(dic) is False
