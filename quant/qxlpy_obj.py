GLOBAL_OBJS = {}


def store_obj(obj) -> str:
    obj_name = str(obj.__class__).split("'")[1] + '_' + str(id(obj))
    GLOBAL_OBJS[obj_name] = obj
    return obj_name

def list_objs() -> list:
    obj_names = []
    for k in GLOBAL_OBJS:
        obj_names.append(k)
    return obj_names

def get_obj(obj_name: str):
    return GLOBAL_OBJS[obj_name]

def del_obj(obj_name: str) -> str:
    if not obj_exists(obj_name):
        ret = f"{obj_name} does not exists."
    else:
        ret = GLOBAL_OBJS.pop(obj_name, None)
        ret = f"{ret} has been removed from cache."
    return ret

def obj_exists(obj_name: str) -> bool:
    if obj_names in GLOBAL_OBJS.keys():
        return True
    return False
