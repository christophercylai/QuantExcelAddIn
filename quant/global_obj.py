"""
Core python caching
"""
from quant import py_logger


GLOBAL_OBJS = {}


def store_obj(obj, prefix: str = "") -> str:
    """
    store an object and return it's id
    """
    prefix = prefix if not prefix else f"{prefix}_"
    obj_name = prefix + str(obj.__class__).split("'")[1] + '_' + str(id(obj))
    GLOBAL_OBJS[obj_name] = obj
    return obj_name

def list_objs() -> list:
    """
    list all cached objects in GLOBAL_OBJS
    """
    obj_names = []
    for k in GLOBAL_OBJS:
        obj_names.append(k)
    return obj_names

def get_obj(obj_name: str):
    """
    return an object by it's id
    """
    if not obj_name in GLOBAL_OBJS:
        err = f"{obj_name} does not exist."
        py_logger.error(err)
        raise RuntimeError(err)
    return GLOBAL_OBJS[obj_name]

def del_obj(obj_name: str) -> str:
    """
    delete an object by it's id
    """
    if not obj_exists(obj_name):
        ret = f"{obj_name} does not exists."
    else:
        ret = GLOBAL_OBJS.pop(obj_name, None)
        ret = f"{ret} has been removed from cache."
    return ret

def obj_exists(obj_name: str) -> bool:
    """
    check if an object exists by it's id
    """
    if obj_name in GLOBAL_OBJS.keys():  # pylint: disable=consider-iterating-dictionary
        return True
    return False
