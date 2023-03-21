from typing import List, Dict

from . import qxlpy_obj


def StoreStrDict(objdict: Dict[str, str]) -> str:
    # returns the address of the Calculate py obj
    # <key: str, value: str>
    return qxlpy_obj.store_obj(objdict)

def GetStrDict(obj_name: str) -> Dict[str, str]:
    return qxlpy_obj.get_obj(obj_name)

# TODO - to be implemented by autogen
def StoreStrList(objlist: List[str]) -> str:
    return qxlpy_obj.store_obj(objlist)

def ListGlobalObjects() -> List[str]:
    # returns a list of stored objects
    return qxlpy_obj.list_objs()

# TODO - to be implemented by autogen
def DeleteObject(obj_name: str) -> str:
    return qxlpy_obj.del_obj(obj_name)

def ObjectExists(obj_name: str) -> bool:
    return qxlpy_obj.obj_exists(obj_name)
