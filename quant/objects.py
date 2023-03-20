from . import qxlpy_obj


def StoreStrDict(objdict: dict) -> str:
    # <key: str, value: str>
    return qxlpy_obj.store_obj(objdict)

def StoreStrList(objlist: list) -> str:
    # [str, str, ...]
    return qxlpy_obj.store_obj(objlist)

def ListGlobalObjects() -> list:
    return qxlpy_obj.list_objs()

def DeleteObject(obj_name: str) -> str:
    return qxlpy_obj.del_obj(obj_name)

def ObjectExists(obj_name: str) -> bool:
    return qxlpy_obj.obj_exists(obj_name)
