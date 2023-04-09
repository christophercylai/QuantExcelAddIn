"""
Excel facing functions for caching python objects
"""
from typing import List, Dict

from . import global_obj

# pylint: disable=invalid-name


def qxlpyStoreStrDict(objdict: Dict[str, str]) -> str:
    """
    store string dictionary and return the id
    """
    # returns the address of the Calculate py obj
    # <key: str, value: str>
    return global_obj.store_obj(objdict)

def qxlpyGetStrDict(obj_name: str) -> Dict[str, str]:
    """
    return the dictionary object
    """
    # returns a dictionary object
    return global_obj.get_obj(obj_name)

# TODO - to be implemented by autogen
def qxlpyStoreStrList(objlist: List[str]) -> str:
    """
    store string list and return the id
    """
    return global_obj.store_obj(objlist)

def qxlpyListGlobalObjects() -> List[str]:
    """
    return a list of cached object ids
    """
    # returns a list of stored objects
    return global_obj.list_objs()

# TODO - to be implemented by autogen
def qxlpyDeleteObject(obj_name: str) -> str:
    """
    delete a cached object by id
    """
    return global_obj.del_obj(obj_name)

def qxlpyObjectExists(obj_name: str) -> bool:
    """
    check if an object exists by id
    """
    # check the existence of an obj
    return global_obj.obj_exists(obj_name)
