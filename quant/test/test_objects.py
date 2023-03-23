from .. import objects as O


def test_objects():
    strdict = {
        "A": "Apple",
        "B": "Bee"
    }
    d = O.StoreStrDict(strdict)

    strlist = ["a", "b"]
    l = O.StoreStrList(strlist)

    globallist = O.ListGlobalObjects()
    assert d in O.ListGlobalObjects()
    assert l in O.ListGlobalObjects()

    ret = O.GetStrDict(d)
    assert ret["A"] == strdict["A"]
    assert ret["B"] == strdict["B"]

    O.DeleteObject(d)
    assert O.ObjectExists(l) == True
    assert O.ObjectExists(d) == False
