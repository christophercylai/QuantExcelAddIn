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

# autogen scripts variables
main_cs = ''
python_cs = ''

main_func = '        public static _RETURN_TYPE_ _FUNCTION_NAME_(_PARAMETERS_)'
python_import = '                dynamic imp = SCOPE.Import("quant._MODULE_NAME_");'
for key, value in autogen_info.items():
    for funcs in value:
        # sub function name
        main_f = re.sub('_FUNCTION_NAME_', funcs[0], main_func)

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
        for arg in args:
            if not arg in arg_default:
                arg_default[arg] = None

        # return type
        annotations = argspec.annotations  # annotations is a dictionary
        if annotations['return'] in type_map:
            main_f = re.sub('_RETURN_TYPE_', type_map[annotations['return']], main_f)
        else:
            # list and dict return string 'SUCCESS' to the func cell
            # results are printed below the function
            main_f = re.sub('_RETURN_TYPE_', 'string', main_f)

        # sub params
        params = ''
        type_checks = ''
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
        main_cs += f'{main_f}\n'
        main_cs += '        {\n'
        main_cs += type_checks
        main_cs += '            PyExecutor pye = new();'
        main_cs += '        }\n\n'

        # python_cs
        python_ipt = re.sub('_MODULE_NAME_', key, python_import)

print(main_cs)




main_cs = """
        {
            CheckEmpty(logmsg);
            CheckEmpty(level);
            PyExecutor pye = new();
            string ret = pye._FUNCTION_NAME_(logmsg, level);
            return ret;
        }
"""

main_excelfunc = '        [ExcelFunction(Name = "_FUNCTION_NAME_")]'





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
