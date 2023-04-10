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


def autogen(gen_main = True, gen_python = True):
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

    funcs_list = []
    # key is file name and value is function name
    for key, value in autogen_info.items():
        if value in funcs_list:
            continue
        funcs_list.append(value)
        for funcs in value:
            main_ret_pye = templates.MAIN_RET_PYE
            main_f = templates.MAIN_F
            main_return_s = templates.MAIN_RETURN_S
            main_excel = templates.MAIN_EXCEL
            main_ld = ''

            python_func = templates.PYTHON_FUNC
            python_call = templates.PYTHON_CALL

            # function name
            main_f = re.sub('_FUNCTION_NAME_', funcs[0], main_f)
            main_ret_pye = re.sub('_FUNCTION_NAME_', funcs[0], main_ret_pye)
            main_excel = re.sub('_FUNCTION_NAME_', funcs[0], main_excel)
            python_func  = re.sub('_FUNCTION_NAME_', funcs[0], python_func)
            python_call  = re.sub('_FUNCTION_NAME_', funcs[0], python_call)

            # function contents
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
            args.reverse()
            for arg in args:
                args_str += f'{arg}, '
                if not arg in arg_default:
                    arg_default[arg] = None
            main_ret_pye = re.sub('_ARGS_', args_str[:-2], main_ret_pye)
            python_call = re.sub('_ARGS_', args_str[:-2], python_call)

            # return type
            type_checks = ''
            annotations = argspec.annotations  # annotations is a dictionary
            ret_type = annotations['return']
            if ret_type in type_map:
                main_f = re.sub('_EXCEL_RETURN_TYPE_', type_map[ret_type], main_f)
                main_ret_pye = re.sub('_PY_RETURN_TYPE_', type_map[ret_type], main_ret_pye)
                main_return_s = re.sub('_RET_', 'ret', main_return_s)
            elif 'List' == ret_type._name:
                # list and dict return string 'SUCCESS' to the func cell
                # results are printed below the function
                main_f = re.sub('_EXCEL_RETURN_TYPE_', 'string', main_f)
                type_checks += f'            CheckEmpty(func_pos);\n'
                main_return_s = re.sub('_RET_', '"SUCCESS"', main_return_s)
###list_type = type_map[ret_type.__args__[0]]
                main_ret_pye = re.sub('_PY_RETURN_TYPE_', type_map[list], main_ret_pye)
                main_ld = templates.MAIN_LIST
            elif 'Dict' == ret_type._name:
                main_f = re.sub('_EXCEL_RETURN_TYPE_', 'string', main_f)
                type_checks += f'            CheckEmpty(func_pos);\n'
                main_return_s = re.sub('_RET_', '"SUCCESS"', main_return_s)
                main_ret_pye = re.sub('_PY_RETURN_TYPE_', 'Dictionary<string, List<object>>', main_ret_pye)
                main_ld = templates.MAIN_DICT
            else:
                raise KeyError(f'{key}.{func[0]}: {ret_type} is not a valid type for C# autogen')

            # parameters
            params = ''
            for ea_arg in args:
                # params
                p_type = annotations[ea_arg]
                if p_type in type_map:
                    params += f'{type_map[p_type]} '
                    type_checks += f'            CheckEmpty({ea_arg});\n'
                elif 'List' in str(p_type):
                    params += f'{type_map[list]} '
                    type_checks += f'            ListCheckEmpty({ea_arg});\n'
                elif 'Dict' in str(p_type):
                    params += f'{type_map[dict]} '
                    type_checks += f'            DictCheckEmpty({ea_arg});\n'
                else:
                    raise KeyError(f'{key}.{func[0]}: {p_type} is not a valid type for C# autogen')
                params += ea_arg
                if arg_default[ea_arg]:
                    params += f' = "{arg_default[ea_arg]}"'
                params += ', '

            python_func = re.sub('_PARAMETERS_', params[:-2], python_func)
            params += 'string func_pos = ""'
            main_f = re.sub('_PARAMETERS_', params, main_f)
            main_array = [
                f'{main_excel}\n',
                f'{main_f}\n',
                '        {\n',
                type_checks,
                '            PyExecutor pye = new();\n',
                f'{main_ret_pye}\n',
                main_ld,
                f'{main_return_s}\n',
                '        }\n\n'
            ]
            for ea_line in main_array:
                main_cs += ea_line

            # python_cs
            python_ipt = templates.PYTHON_IPT
            python_ipt = re.sub('_MODULE_NAME_', key, python_ipt)

            python_array = [
                f'{python_ipt}\n',
                f'{python_call}',
            ]
            python_gil = ''
            for ea_line in python_array:
                python_gil += ea_line
            python_cs += f'{python_func}'
            python_cs += re.sub('_BODY_', python_gil, templates.PYTHON_GIL)

    if gen_main:
        main_body = re.sub('_BODY_', main_cs, templates.MAIN_BODY)
        with open(str(qxlpydir / 'main.cs'), 'w') as main_cs_f:
            main_cs_f.write(main_body)

    if gen_python:
        print(python_cs)
