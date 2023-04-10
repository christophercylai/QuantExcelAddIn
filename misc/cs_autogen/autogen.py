import os
import site
from pathlib import Path
import inspect
import re
import importlib
import typing


qxlpydir = Path(os.getenv('QXLPYDIR'))
quantdir = qxlpydir / 'quant'

site.addsitedir(str(qxlpydir))

from . import templates

# gather info on functions that start with qxlpy
autogen_info = {}
for ea_path in quantdir.glob('*.py'):
    filename = re.sub('.py', '', ea_path.name)
    mod = importlib.import_module(f'.{filename}', 'quant')
    funcs = inspect.getmembers(mod, inspect.isfunction)
    qxlpy_f = []
    for ea_f in funcs:
        if 'qxlpy' in ea_f[0]:
            qxlpy_f.append(ea_f)  # ea_f is a tuple = (func_name, func_obj)
    if qxlpy_f:
        autogen_info[filename] = funcs

# keys are python types, values are C# types
type_map = {
    bool: 'bool',
    str: 'string',
    int: 'long',
    float: 'double',
    list: 'object[]',
    dict: 'object[,]'
}

to_type_map = {
    str: 'ToString',
    int: 'long',
    float: 'double'
}

# autogen scripts variables
main_cs = ''
python_cs = ''

for key, value in autogen_info.items():
    for funcs in value:
        ret_pye = templates.RET_PYE
        main_f = templates.MAIN_F
        return_s = templates.RETURN_S
        main_list = templates.MAIN_LIST

        # sub function name
        main_f = re.sub('_FUNCTION_NAME_', funcs[0], main_f)
        ret_pye = re.sub('_FUNCTION_NAME_', funcs[0], ret_pye)

        # params and function contents
        argspec = inspect.getfullargspec(funcs[1])
        defaults = argspec.defaults
        arg_default = {}
        args = argspec.args
        args.reverse()  # params with defaults cannot precede params without
        if defaults:
            defaults = list(defaults)
            defaults.reverse()
            for num in range(len(defaults)):
                arg_default[args[num]] = defaults[num]
        args_str = ''
        for arg in args:
            args_str += f'{arg}, '
            if not arg in arg_default:
                arg_default[arg] = None
        ret_pye = re.sub('_ARGS_', args_str[:-2], ret_pye)

        # return type
        type_checks = ''
        annotations = argspec.annotations  # annotations is a dictionary
        if annotations['return'] in type_map:
            main_f = re.sub('_EXCEL_RETURN_TYPE_', type_map[annotations['return']], main_f)
            ret_pye = re.sub('_PY_RETURN_TYPE_', type_map[annotations['return']], ret_pye)
            return_s = re.sub('_RET_', 'ret', return_s)
            main_list = ''
        elif 'List' == annotations['return']._name:
            # list and dict return string 'SUCCESS' to the func cell
            # results are printed below the function
            main_f = re.sub('_EXCEL_RETURN_TYPE_', 'string', main_f)
            ret_pye = re.sub('_PY_RETURN_TYPE_', type_map[list], ret_pye)
            type_checks += f'            CheckEmpty(func_pos)\n'
            return_s = re.sub('_RET_', '"SUCCESS"', return_s)
            list_type = type_map[annotations['return'].__args__[0]]
        elif 'Dict' == annotations['return']._name:
            main_f = re.sub('_EXCEL_RETURN_TYPE_', 'string', main_f)
            ret_pye = re.sub('_PY_RETURN_TYPE_', 'Dictionary<string, List<string>>', ret_pye)
            type_checks += f'            CheckEmpty(func_pos)\n'
            return_s = re.sub('_RET_', '"SUCCESS"', return_s)

        # sub params
        params = ''
        for ky, vlu in arg_default.items():
            # params
            p_type = annotations[ky]
            if p_type in type_map:
                params += f'{type_map[p_type]} '
                type_checks += f'            CheckEmpty({ky})\n'
            elif 'List' in str(p_type):
                params += f'{type_map[list]} '
                type_checks += f'            ListCheckEmpty({ky})\n'
            elif 'Dict' in str(p_type):
                params += f'{type_map[dict]} '
                type_checks += f'            DictCheckEmpty({ky})\n'
            else:
                raise KeyError(f'{key}.{func[0]}: {type_map[p_type]} is not a valid C# type')
            params += ky
            if arg_default[ky]:
                params += f' = "{arg_default[ky]}"'
            params += ', '

        params += 'string func_pos = ""'
        main_f = re.sub('_PARAMETERS_', params, main_f)
        main_array = [
            f'{main_f}\n',
            '        {\n',
            type_checks,
            '            PyExecutor pye = new();\n',
            f'{ret_pye}\n',
            main_list,
            f'{return_s}\n',
            '        }\n\n'
        ]
        for ea_line in main_array:
            main_cs += ea_line

        # python_cs
        python_ipt = '                dynamic imp = SCOPE.Import("quant._MODULE_NAME_");'
        python_ipt = re.sub('_MODULE_NAME_', key, python_ipt)

#print(main_cs)





python_cs = """
        public string _FUNCTION_NAME_(string logmsg, string level)
        {
            using (Py.GIL())
            {
                string ret = imp._FUNCTION_NAME_(logmsg, level);
                return ret;
            }
        }
"""
