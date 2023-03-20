from pathlib import Path
import sys

from .qxlpy_log import py_logger, cs_logger


# add quant module to be searchable python
# i.e., give all modules access to this __init__.py
sys.path.append(Path(__path__[0]))


### ========== Import Modules ========== ###
# import ONLY the first layer of modules
# anything inside the submodules is should always be private
from . import calculate
from . import cslog


### ========== Global Variables ========== ###
# objects can be stored by Python and then retrived by C# later with GLOBAL_OBJS
GLOBAL_OBJS = {}


### ========== Global Functions ========== ###

def STORE_OBJ(obj) -> str:
    obj_name = str(obj.__class__).split("'")[1] + '_' + str(id(obj))
    GLOBAL_OBJS[obj_name] = obj
    return obj_name

def LIST_OBJS() -> list:
    obj_names = []
    for k in GLOBAL_OBJS:
        obj_names.append(k)
        return obj_names

def GET_OBJ(obj_name: str):
    return GLOBAL_OBJS[obj_name]

def DEL_OBJ(obj_name: str):
    GLOBAL_OBJS.pop(obj_name, None)

def OBJ_EXISTS(obj_name: str) -> bool:
    if obj_names in GLOBAL_OBJS.keys():
        return True
    return False
