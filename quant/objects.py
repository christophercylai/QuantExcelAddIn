"""
Excel facing functions for caching python objects
"""
from typing import List, Dict

from . import global_obj

# pylint: disable=invalid-name


def qxlpyStoreStrDict(objdict: Dict[str, str], prefix: str) -> str:
    """
    store string dictionary and return the id
    """
    # returns the address of the Calculate py obj
    # <key: str, value: str>
    return global_obj.store_obj(objdict, prefix)

def qxlpyGetStrDict(obj_name: str) -> Dict[str, str]:
    """
    return the string dictionary
    """
    # returns a dictionary object
    return global_obj.get_obj(obj_name)

def qxlpyGetObjDict(obj_name: str) -> Dict[object, object]:
    """
    return the object dictionary
    """
    return global_obj.get_obj(obj_name)

def qxlpyStoreStrList(objlist: List[str], prefix: str) -> str:
    """
    store string list and return the id
    """
    return global_obj.store_obj(objlist, prefix)

def qxlpyGetStrList(obj_name: str) -> List[str]:
    """
    return a list of strings
    """
    return global_obj.get_obj(obj_name)

def qxlpyGetObjList(obj_name: str) -> List[object]:
    """
    return a list of objects
    """
    return global_obj.get_obj(obj_name)

def qxlpyStoreDoubleList(objlist: List[float], prefix: str) -> str:
    """
    store float list and return the id
    """
    return global_obj.store_obj(objlist, prefix)

def qxlpyStoreStrTable(nested_objlist: List[List[str]], prefix: str) -> str:
    """
    store a table of strings, i.e. List[List[str]]
    """
    return global_obj.store_obj(nested_objlist, prefix)

def qxlpyGetStrTable(obj_name: str) -> List[List[str]]:
    """
    return a table of strings
    """
    return global_obj.get_obj(obj_name)

def qxlpyGetObjTable(obj_name: str) -> List[List[object]]:
    """
    return a table of objects
    """
    return global_obj.get_obj(obj_name)

def qxlpyListGlobalObjects() -> List[str]:
    """
    return a list of cached object ids
    """
    # returns a list of stored objects
    return global_obj.list_objs()

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
