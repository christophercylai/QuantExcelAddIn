"""
Auto generate Qxlpy C# code for the Excel Addin
Files generated: main.cs, python.cs
"""
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
            # only func that start with qxlpy will be processed
            if ea_f[0].startswith('qxlpy'):
                qxlpy_f.append(ea_f)  # ea_f is a tuple = (func_name, func_obj)
        if qxlpy_f:
            autogen_info[filename] = qxlpy_f

    # keys are python types, values are C# types
    type_map = {
        bool: 'bool',
        str: 'string',
        int: 'long',
        float: 'double',
        object: 'dynamic',
        list: 'object[]',
        dict: 'object[,]'
    }

    # Reference:
    # https://learn.microsoft.com/en-us/dotnet/api/system.convert?redirectedfrom=MSDN&view=net-6.0
    to_type_map = {
        bool: 'Convert.ToBoolean',
        str: 'Convert.ToString',
        int: 'Convert.ToInt64',
        float: 'Convert.ToDouble',
        object: 'GetToTypeByValue'
    }

    # Reference
    # https://pythonnet.github.io/pythonnet/reference.html
    py_type_map = {
        str: 'new PyString',
        int: 'new PyInt',
        float: 'new PyFloat',
        object: 'GetPyTypeByValue'
    }

    # autogen scripts variables
    main_cs = ''
    main_docstring_func = ''
    python_cs = ''
    python_module_list = []

    funcs_list = []
    # key is file name and value is list of function names
    for key, value in autogen_info.items():
        for func in value:
            if func[0] in funcs_list:
                continue
            funcs_list.append(func[0])

            main_ret_pye = templates.MAIN_RET_PYE
            main_f = templates.MAIN_F
            main_return_s = templates.MAIN_RETURN_S
            main_excel = templates.MAIN_EXCEL

            python_func = templates.PYTHON_FUNC
            python_call = templates.PYTHON_CALL
            python_dl_inputs = ''
            python_dl_return = ''

            # function name
            # func[0] is the function name and func[1] are the parameters
            main_f = re.sub('_FUNCTIONNAME_', func[0], main_f)
            main_ret_pye = re.sub('_FUNCTIONNAME_', func[0], main_ret_pye)
            main_excel = re.sub('_FUNCTIONNAME_', func[0], main_excel)
            python_func  = re.sub('_FUNCTIONNAME_', func[0], python_func)
            python_call = re.sub('_FUNCTIONNAME_', func[0], python_call)
            python_call = re.sub('_PYTHONIMPORT_', key.upper(), python_call)

            # docstring
            docstring = inspect.getdoc(func[1])
            main_docstring = re.sub("_FUNCTIONNAME_", func[0], templates.MAIN_DOCSTRING)
            main_docstring = re.sub("_DOCSTRING_", docstring, main_docstring)
            main_docstring_func += main_docstring

### argspec example ###
# >>> import typing
# >>> import inspect
# >>> def blah(s: str, l: typing.List[str] = [], d: typing.Dict[int, float] = {5: 5.5, 6: 6.6}) -> typing.Dict[str, int]:
# ...   return {'a': 1}
# ...
# >>>
# >>> argspec = inspect.getfullargspec(blah)
# >>> argspec
# FullArgSpec(
#    args=['s', 'l', 'd'], varargs=None, varkw=None,
#    defaults=([], {5: 5.5, 6: 6.6}), kwonlyargs=[], kwonlydefaults=None,
#    annotations={'return': typing.Dict[str, int], 's': str, 'l': typing.List[str], 'd': typing.Dict[int, float]}
# )

            # function contents
            argspec = inspect.getfullargspec(func[1])
            arg_default = {}
            args = argspec.args
            args.reverse()  # params with defaults cannot precede params without
            if argspec.defaults:
                defaults = list(argspec.defaults)
                defaults.reverse()
                for num in range(len(defaults)):
                    arg_default[args[num]] = defaults[num]
            main_args_str = ''
            args.reverse()
            for arg in args:
                main_args_str += f'_{arg}, '
                if not arg in arg_default:
                    arg_default[arg] = None
            main_ret_pye = re.sub('_ARGS_', main_args_str[:-2], main_ret_pye)

            # return type
            type_checks = ''
            annotations = argspec.annotations  # annotations is a dictionary
            assert 'return' in annotations, f"'{func[1].__name__}' has no return type"
            ret_type = annotations['return']
            if ret_type in type_map:
                main_f = re.sub('_EXCELRETURNTYPE_', type_map[ret_type], main_f)
                main_ret_pye = re.sub('_PYRETURNTYPE_', type_map[ret_type], main_ret_pye)
                main_return_s = re.sub('_RET_', 'ret', main_return_s)

                python_func  = re.sub('_FUNCTYPE_', type_map[ret_type], python_func)
                python_call  = re.sub('_FUNCTYPE_', type_map[ret_type], python_call)
            elif 'List' == ret_type._name:
                # list and dict return string 'SUCCESS' to the func cell
                # results are printed below the function
                main_f = re.sub('_EXCELRETURNTYPE_', 'object', main_f)
                main_return_s = re.sub('_RET_', 'ret', main_return_s)
                main_ret_pye = re.sub('_PYRETURNTYPE_', type_map[dict], main_ret_pye)
                list_type = ""
                pyobj_name = "pyobj"
                if ret_type.__args__[0] in to_type_map:
                    list_type = ret_type.__args__[0]
                    python_dl_return = templates.PYTHON_LIST_RETURN
                elif 'List' == ret_type.__args__[0]._name:
                    list_type = ret_type.__args__[0].__args__[0]
                    python_dl_return = templates.PYTHON_NESTED_LIST_RETURN
                    pyobj_name = "internal_pyobj"
                else:
                    raise KeyError(f'{key}.{func[0]}: {ret_type} is not a valid type for C# autogen')
                python_dl_return = re.sub('_FUNCNAME_', func[0], python_dl_return)
                if list_type == object:
                    python_dl_return = re.sub('_PYTYPE_', f"GetToTypeByValue({pyobj_name})", python_dl_return)
                else:
                    python_dl_return = re.sub('_PYTYPE_', f"{to_type_map[list_type]}({pyobj_name}.ToString())", python_dl_return)
                python_dl_return = re.sub('_PYTHONIMPORT_', key.upper(), python_dl_return)
                python_func  = re.sub('_FUNCTYPE_', type_map[dict], python_func)
                python_call  = ''
            elif 'Dict' == ret_type._name:
                main_f = re.sub('_EXCELRETURNTYPE_', 'object', main_f)
                main_return_s = re.sub('_RET_', 'ret', main_return_s)
                main_ret_pye = re.sub('_PYRETURNTYPE_', type_map[dict], main_ret_pye)

                python_func  = re.sub('_FUNCTYPE_', type_map[dict], python_func)
                python_call  = ''
                key_type = ret_type.__args__[0]
                val_type = ret_type.__args__[1]
                python_dl_return = templates.PYTHON_DICT_RETURN
                if key_type == object:
                    python_dl_return = re.sub('_KEYPYTYPE_', f"GetToTypeByValue(key)", python_dl_return)
                else:
                    python_dl_return = re.sub('_KEYPYTYPE_', f"{to_type_map[val_type]}(key.ToString())", python_dl_return)
                if val_type == object:
                    python_dl_return = re.sub('_VALPYTYPE_', f"GetToTypeByValue(pydict_ret.GetItem(key))", python_dl_return)
                else:
                    python_dl_return = re.sub('_VALPYTYPE_', f"{to_type_map[val_type]}(pydict_ret.GetItem(key).ToString())", python_dl_return)
                python_dl_return = re.sub('_FUNCNAME_', func[0], python_dl_return)
                python_dl_return = re.sub('_FUNCNAME_', func[0], python_dl_return)
                python_dl_return = re.sub('_PYTHONIMPORT_', key.upper(), python_dl_return)
            else:
                raise KeyError(f'{key}.{func[0]}: {ret_type} is not a valid type for C# autogen')

            # parameters
            main_params = ''
            python_params = ''
            pycall_params = ''
            for ea_arg in args:
                # params
                assert ea_arg in annotations, f"'{func[1].__name__}': param '{ea_arg}' has no type"
                p_type = annotations[ea_arg]
                python_dl_input = ''
                if p_type in type_map:
                    main_params += f'string '
                    python_params += f'{type_map[p_type]} '
                    # by default, call CheckEmpty with assert = true, unless there is a default value or p_type == str
                    default_value =  "" if p_type == str else ", null, true"
                    if arg_default[ea_arg] is not None:
                        if p_type == bool:
                            default_value = f', "{str(arg_default[ea_arg]).lower()}"'
                        else:
                            default_value = f', "{arg_default[ea_arg]}"'
                    type_checks += f'            {type_map[p_type]} _{ea_arg} = {to_type_map[p_type]}(CheckEmpty({ea_arg}{default_value}));\n'

                    pycall_params += ea_arg
                elif str(p_type).startswith('typing.List'):
                    if p_type.__args__[0] in type_map:
                        main_params += f'{type_map[list]} '
                        python_params += f'{type_map[list]} '
                        type_checks += f'            {type_map[list]} _{ea_arg} = ListCheckEmpty({ea_arg});\n'

                        python_dl_input = templates.PYTHON_LIST_INPUT
                        python_dl_input = re.sub('_ARGNAME_', ea_arg, python_dl_input)
                        python_dl_input = re.sub('_ARGTYPE_', type_map[p_type.__args__[0]], python_dl_input)
                        python_dl_input = re.sub('_TOTYPE_', to_type_map[p_type.__args__[0]], python_dl_input)
                        python_dl_input = re.sub('_PYTYPE_', py_type_map[p_type.__args__[0]], python_dl_input)
                        pycall_params += f'pylist_{ea_arg}'
                    elif str(p_type.__args__[0]).startswith('typing.List'):
                        main_params += f'{type_map[dict]} '
                        python_params += f'{type_map[dict]} '
                        type_checks += f'            {type_map[dict]} _{ea_arg} = DictCheckEmpty({ea_arg});\n'

                        python_dl_input = templates.PYTHON_NESTED_LIST_INPUT
                        python_dl_input = re.sub('_ARGNAME_', ea_arg, python_dl_input)
                        python_dl_input = re.sub('_ARGTYPE_', type_map[p_type.__args__[0].__args__[0]], python_dl_input)
                        python_dl_input = re.sub('_TOTYPE_', to_type_map[p_type.__args__[0].__args__[0]], python_dl_input)
                        python_dl_input = re.sub('_PYTYPE_', py_type_map[p_type.__args__[0].__args__[0]], python_dl_input)
                        pycall_params += f'pylist_{ea_arg}'
                elif str(p_type).startswith('typing.Dict'):
                    main_params += f'{type_map[dict]} '
                    python_params += f'{type_map[dict]} '
                    type_checks += f'            {type_map[dict]} _{ea_arg} = DictCheckEmpty({ea_arg});\n'

                    python_dl_input = templates.PYTHON_DICT_INPUT
                    python_dl_input = re.sub('_ARGNAME_', ea_arg, python_dl_input)
                    python_dl_input = re.sub('_KEYTYPE_', type_map[p_type.__args__[0]], python_dl_input)
                    python_dl_input = re.sub('_VALTYPE_', type_map[p_type.__args__[1]], python_dl_input)
                    python_dl_input = re.sub('_TOKEYTYPE_', to_type_map[p_type.__args__[0]], python_dl_input)
                    python_dl_input = re.sub('_TOVALTYPE_', to_type_map[p_type.__args__[1]], python_dl_input)
                    python_dl_input = re.sub('_PYTYPEVAL_', py_type_map[p_type.__args__[1]], python_dl_input)
                    pycall_params += f'pydict_{ea_arg}'
                else:
                    raise KeyError(f'{key}.{func[0]}: {p_type} is not a valid type for C# autogen')

                main_params += ea_arg
                python_params += ea_arg
                if arg_default[ea_arg] is not None:
                    if p_type == bool:
                        main_params += f' = "{str(arg_default[ea_arg]).lower()}"'
                    elif p_type in [str, int, float]:
                        main_params += f' = "{arg_default[ea_arg]}"'
                    else:
                        main_params += ' = null'
                main_params += ', '
                python_params += ', '
                pycall_params += ', '
                python_dl_inputs += python_dl_input

            main_f = re.sub('_PARAMETERS_', main_params[:-2], main_f)
            main_array = [
                f'{main_excel}\n',
                f'{main_f}\n',
                '        {\n',
                type_checks,
                f'{main_ret_pye}\n',
                f'{main_return_s}\n',
                '        }\n\n'
            ]
            for ea_line in main_array:
                main_cs += ea_line

            # python_cs
            python_func = re.sub('_PARAMETERS_', python_params[:-2], python_func)
            python_call = re.sub('_ARGS_', pycall_params[:-2], python_call)
            python_dl_return = re.sub('_PYPARAMS_', pycall_params[:-2], python_dl_return)
            if key not in python_module_list:
                python_module_list.append(key)

            python_f_body = ''
            python_f_body += python_call
            python_gil = templates.PYTHON_GIL
            python_gil = re.sub('_DLINPUTS_', python_dl_inputs, python_gil)
            python_gil = re.sub('_BODY_', python_f_body, python_gil)
            python_gil = re.sub('_DLRETURN_', python_dl_return, python_gil)
            python_cs += python_func
            python_cs += f'{python_gil}\n'
            python_mods = ''
            python_import_list = ''
            python_reload_list = ''
            for ea_module in python_module_list:
                python_mods += f"{ea_module.upper()}, "
                python_import_list += f'                {ea_module.upper()} = SCOPE.Import("quant.{ea_module}");\n'
                python_reload_list += f'                importlib.reload({ea_module.upper()});\n'
            python_mods = python_mods[:-2]
            python_import_list = python_import_list[:-1]
            python_body = templates.PYTHON_BODY
            python_body = re.sub('_PYTHONMODS_', python_mods, python_body)
            python_body = re.sub('_PYIMPORTLIST_', python_import_list, python_body)
            python_body = re.sub('_PYRELOADLIST_', python_reload_list, python_body)
            python_body = re.sub('_BODY_', python_cs, python_body)

    if gen_main:
        main_body = re.sub('_BODY_', main_cs, templates.MAIN_BODY)
        main_docstring_func = main_docstring_func[:-2]  # remove comma and newline
        main_body = re.sub('_COMMENTS_MAP_', main_docstring_func, main_body)
        write_to = ''
        if dryrun:
            write_to = 'main.cs.bak'
        else:
            write_to = 'main.cs'
        with open(str(qxlpydir / write_to), 'w') as main_cs_f:
            main_cs_f.write(main_body)

    if gen_python:
        if dryrun:
            write_to = 'python.cs.bak'
        else:
            write_to = 'python.cs'
        with open(str(qxlpydir / write_to), 'w') as python_cs_f:
            python_cs_f.write(python_body)
