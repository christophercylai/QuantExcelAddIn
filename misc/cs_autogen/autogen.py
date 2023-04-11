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


def autogen(gen_main = True, gen_python = True, dryrun = False):
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

    # Reference:
    # https://learn.microsoft.com/en-us/dotnet/api/system.convert?redirectedfrom=MSDN&view=net-6.0
    to_type_map = {
        bool: 'ToBool',
        str: 'ToString',
        int: 'ToInt64',
        float: 'ToDouble'
    }

    # autogen scripts variables
    main_cs = ''
    python_cs = ''

    funcs_list = []
    # key is file name and value is function name(s)
    for key, value in autogen_info.items():
        if value in funcs_list:
            continue
        funcs_list.append(value)
        for func in value:
            main_ret_pye = templates.MAIN_RET_PYE
            main_f = templates.MAIN_F
            main_return_s = templates.MAIN_RETURN_S
            main_excel = templates.MAIN_EXCEL
            main_ld = ''

            python_func = templates.PYTHON_FUNC
            python_call = templates.PYTHON_CALL
            python_dl_input = ''
            python_dl_return = ''

            # function name
            # func[0] is the function name and func[2] are the parameters
            main_f = re.sub('_FUNCTION_NAME_', func[0], main_f)
            main_ret_pye = re.sub('_FUNCTION_NAME_', func[0], main_ret_pye)
            main_excel = re.sub('_FUNCTION_NAME_', func[0], main_excel)
            python_func  = re.sub('_FUNCTION_NAME_', func[0], python_func)
            python_call = re.sub('_FUNCTION_NAME_', func[0], python_call)

            # function contents
            argspec = inspect.getfullargspec(func[1])
            defaults = argspec.defaults
            arg_default = {}
            args = argspec.args
            args.reverse()  # params with defaults cannot precede params without
            if defaults:
                defaults = list(defaults)
                defaults.reverse()
                for num in range(len(defaults)):
                    arg_default[args[num]] = defaults[num]
            main_args_str = ''
            args.reverse()
            for arg in args:
                main_args_str += f'{arg}, '
                if not arg in arg_default:
                    arg_default[arg] = None
            main_ret_pye = re.sub('_ARGS_', main_args_str[:-2], main_ret_pye)

            # return type
            type_checks = ''
            annotations = argspec.annotations  # annotations is a dictionary
            ret_type = annotations['return']
            if ret_type in type_map:
                main_f = re.sub('_EXCEL_RETURN_TYPE_', type_map[ret_type], main_f)
                main_ret_pye = re.sub('_PY_RETURN_TYPE_', type_map[ret_type], main_ret_pye)
                main_return_s = re.sub('_RET_', 'ret', main_return_s)

                python_func  = re.sub('_FUNC_TYPE_', type_map[ret_type], python_func)
                python_call  = re.sub('_FUNC_TYPE_', type_map[ret_type], python_call)
            elif 'List' == ret_type._name:
                # list and dict return string 'SUCCESS' to the func cell
                # results are printed below the function
                main_f = re.sub('_EXCEL_RETURN_TYPE_', 'string', main_f)
                type_checks += f'            CheckEmpty(func_pos);\n'
                main_return_s = re.sub('_RET_', '"SUCCESS"', main_return_s)
                main_ret_pye = re.sub('_PY_RETURN_TYPE_', type_map[list], main_ret_pye)
                main_ld = templates.MAIN_LIST

                list_type = type_map[ret_type.__args__[0]]
                python_dl_return = templates.PYTHON_LIST_RETURN
                python_dl_return = re.sub('_LIST_TYPE_', type_map[ret_type.__args__[0]], python_dl_return)
                python_dl_return = re.sub('_TO_TYPE_', to_type_map[ret_type.__args__[0]], python_dl_return)
                python_func  = re.sub('_FUNC_TYPE_', 'object[]', python_func)
                python_call  = re.sub('_FUNC_TYPE_', 'object[]', python_call)
            elif 'Dict' == ret_type._name:
                main_f = re.sub('_EXCEL_RETURN_TYPE_', 'string', main_f)
                type_checks += f'            CheckEmpty(func_pos);\n'
                main_return_s = re.sub('_RET_', '"SUCCESS"', main_return_s)
                main_ret_pye = re.sub('_PY_RETURN_TYPE_', 'List<List<object>>', main_ret_pye)
                main_ld = templates.MAIN_DICT

                python_func  = re.sub('_FUNC_TYPE_', 'List<List<object>>', python_func)
                python_call  = re.sub('_FUNC_TYPE_', 'List<List<object>>', python_call)
                python_dl_return = templates.PYTHON_DICT_RETURN
                python_dl_return = re.sub('_FUNC_NAME_', func[0], python_dl_return)
                python_dl_return = re.sub('_TO_KEY_TYPE_', to_type_map[ret_type.__args__[0]], python_dl_return)
                python_dl_return = re.sub('_TO_VAL_TYPE_', to_type_map[ret_type.__args__[1]], python_dl_return)
            else:
                raise KeyError(f'{key}.{func[0]}: {ret_type} is not a valid type for C# autogen')

            # parameters
            main_params = ''
            py_params = ''
            for ea_arg in args:
                # params
                p_type = annotations[ea_arg]
                if p_type in type_map:
                    main_params += f'{type_map[p_type]} '
                    type_checks += f'            CheckEmpty({ea_arg});\n'

                    py_params += ea_arg
                elif 'List' in str(p_type):
                    main_params += f'{type_map[list]} '
                    type_checks += f'            ListCheckEmpty({ea_arg});\n'

                    python_dl_input = templates.PYTHON_LIST_INPUT
                    python_dl_input = re.sub('_ARG_NAME_', ea_arg, python_dl_input)
                    python_dl_input = re.sub('_ARG_TYPE_', type_map[p_type.__args__[0]], python_dl_input)
                    python_dl_input = re.sub('_TO_TYPE_', to_type_map[p_type.__args__[0]], python_dl_input)
                    py_params += f'pylist_{ea_arg}'
                elif 'Dict' in str(p_type):
                    main_params += f'{type_map[dict]} '
                    type_checks += f'            DictCheckEmpty({ea_arg});\n'
                else:
                    raise KeyError(f'{key}.{func[0]}: {p_type} is not a valid type for C# autogen')
                main_params += ea_arg
                if arg_default[ea_arg]:
                    main_params += f' = "{arg_default[ea_arg]}"'
                main_params += ', '
                py_params += ', '

            main_params += 'string func_pos = ""'
            main_f = re.sub('_PARAMETERS_', main_params, main_f)
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
            python_func = re.sub('_PARAMETERS_', main_params.replace('string func_pos = ""', '')[:-2], python_func)
            python_call = re.sub('_ARGS_', py_params, python_call)
            python_ipt = templates.PYTHON_IPT
            python_ipt = re.sub('_MODULE_NAME_', key, python_ipt)

            python_array = [
                f'{python_ipt}\n',
                f'{python_call}',
            ]
            python_f_body = ''
            for ea_line in python_array:
                python_f_body += ea_line
            python_gil = templates.PYTHON_GIL
            python_gil = re.sub('_DL_INPUTS_', python_dl_input, python_gil)
            python_gil = re.sub('_BODY_', python_f_body, python_gil)
            python_gil = re.sub('_DL_RETURN_', python_dl_return, python_gil)
            python_cs += f'{python_func}'
            python_cs += f'{python_gil}'

    if gen_main:
        main_body = re.sub('_BODY_', main_cs, templates.MAIN_BODY)
        if dryrun:
            print(main_body)
        else:
            with open(str(qxlpydir / 'main.cs'), 'w') as main_cs_f:
                main_cs_f.write(main_body)

    if gen_python:
        if dryrun:
            print(python_cs)
        else:
            with open(str(qxlpydir / 'python.cs'), 'w') as python_cs_f:
                python_cs_f.write(python_cs)
