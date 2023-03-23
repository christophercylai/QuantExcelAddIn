"""
Excel facing functions for caching python objects
"""
from typing import List, Dict

from . import qxlpy_obj


def StoreStrDict(objdict: Dict[str, str]) -> str:
    """
    store string dictionary and return the id
    """
    # returns the address of the Calculate py obj
    # <key: str, value: str>
    return qxlpy_obj.store_obj(objdict)

def GetStrDict(obj_name: str) -> Dict[str, str]:
    """
    return the dictionary object
    """
    # returns a dictionary object
    return qxlpy_obj.get_obj(obj_name)

# TODO - to be implemented by autogen
def StoreStrList(objlist: List[str]) -> str:
    """
    store string list and return the id
    """
    return qxlpy_obj.store_obj(objlist)

def ListGlobalObjects() -> List[str]:
    """
    return a list of cached object ids
    """
    # returns a list of stored objects
    return qxlpy_obj.list_objs()

# TODO - to be implemented by autogen
def DeleteObject(obj_name: str) -> str:
    """
    delete a cached object by id
    """
    return qxlpy_obj.del_obj(obj_name)

def ObjectExists(obj_name: str) -> bool:
    """
    check if an object exists by id
    """
    # check the existence of an obj
    return qxlpy_obj.obj_exists(obj_name)
