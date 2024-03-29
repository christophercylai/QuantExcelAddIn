### main.cs string templates ###
MAIN_BODY = r'''using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace qxlpy
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
      <ribbon>
        <tabs>
          <tab id='qxltab' label='QXLPY'>
            <group id='qxlpy_e' label='Function Expansion'>
              <button id='expandfunc' label='Expand Function'
                onAction='expandFuncButton' size='large' screentip='Expand Function (Ctrl-INS)'
                imageMso='ConditionalFormattingColorScalesGallery' />
            </group >
            <group id='qxlpy_d' label='Function Deletion'>
              <button id='removefunc' label='Remove Function'
                onAction='removeFuncButton' size='large' screentip='Remove Function (Ctrl-DEL)'
                imageMso='RecordsDeleteRecord' />
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            dynamic xlApp = ExcelDnaUtil.Application;
        }

        public void expandFuncButton(IRibbonControl control)
        {
            AutoFill.AutoFuncFormat();
        }

        public void removeFuncButton(IRibbonControl control)
        {
            AutoFill.AutoFuncClear();
        }
    }
    // END: public class RibbonController : ExcelRibbon


    public class AutoRun : IExcelAddIn
    {
        public void AutoOpen()
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            xlApp.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(AppWbActivate);
        }

        public void AutoClose() {}

        private void AppWbActivate(Excel.Workbook Wb)
        {
            // disable auto calculate
            dynamic xlApp = ExcelDnaUtil.Application;
            xlApp.Calculation = -4135;
            xlApp.CalculateBeforeSave = false;
        }
    }

    public static class AutoFill
    {
        public static string? old_formula;

        private static bool SheetExists()
        {
            // Check if there is an active worksheet
            dynamic xlApp = ExcelDnaUtil.Application;
            var sheet = xlApp.ActiveSheet;
            if (sheet == null) {
                ExManip.WriteLog("There is no active sheet", "WARNING");
                return false;
            }
            return true;
        }

        private static string FormulaExists(int x, int y)
        {
            // Check if the cell has a formula
            dynamic xlApp = ExcelDnaUtil.Application;
            if (!xlApp.Cells(y, x).HasFormula) {
                ExManip.WriteLog("Seleted cell does not have a formula", "WARNING");
                return "";
            }

            // Get formula name
            old_formula = xlApp.Cells(y, x).Formula;
            var rgx_f = new Regex(@"[a-zA-Z][a-zA-Z0-9_]+");
            Match match_f = rgx_f.Match(old_formula);

            if (!match_f.Success) {
                ExManip.WriteLog("Formula must start with [a-zA-Z] and followed by [a-zA-Z0-9_]+", "WARNING");
                return "";
            }

            // Check whether formula is a method of ExcelFunc
            MethodInfo method_info = typeof(ExcelFunc).GetMethod(match_f.Value);
            if (method_info == null) {
                ExManip.WriteLog("The supplied formula is not a QXLPY UDF", "WARNING");
                return "";
            }

            return match_f.Value;
        }

        private static string GetComment(string func_name)
        {
            var comments_map = new Dictionary<string, string>() {
                {"qxlpyReloadPyModules",
 @"reload uxlpy python modules"
                },
_COMMENTS_MAP_
            };
            if (! comments_map.ContainsKey(func_name)) {
                string errmsg = $"Python function '{func_name}' does not have a docstring";
                ExManip.WriteLog(errmsg, "ERROR");
                throw new IndexOutOfRangeException(errmsg);
            }
            return comments_map[func_name];
        }

        public static void AutoFuncFormat()
        {
            // auto format a UDF from ExcelFunc
            dynamic xlApp = ExcelDnaUtil.Application;
            if (!SheetExists()) {
                return;
            }

            var sheet = xlApp.ActiveSheet;
            int[] ac = ExManip.GetActiveCellPos();
            int y = ac[0];
            int x = ac[1];

            string f = FormulaExists(x, y);
            if (f == "") {
                return;
            }

            MethodInfo method_info = typeof(ExcelFunc).GetMethod(f);
            ParameterInfo[] param_info = method_info.GetParameters();
            int p_len = param_info.Length + 1;

            // Cells formatting
            // Title = function name
            // backtrack is a record for undo changes in RangeEmpty()
            var backtrack = new List<dynamic>();
            backtrack.Add(xlApp.Cells(y, x));
            ExManip.RangeEmpty(xlApp.Cells(y, x + 1), backtrack);
            xlApp.Cells(y, x).Value = f;
            xlApp.Cells(y, x).Font.Bold = true;
            xlApp.Cells(y, x).Font.Color = Color.FromArgb(0, 255, 255, 255);
            xlApp.Cells(y, x).Interior.Color = Color.FromArgb(142, 0, 111, 41);
            string comment = GetComment(f);
            xlApp.Cells(y, x).AddComment(comment);
            xlApp.Cells(y, x).Comment.Shape.TextFrame.Autosize = true;
            ExManip.GetRange(y, x, y, x +1).Merge();

            // parsing the comment for a list of default values
            StringReader reader = new StringReader(comment);
            string ea_line = "";
            var param_choices = new Dictionary<string, string> { };
            while ((ea_line = reader.ReadLine()) != null) {
                if (ea_line.Contains("::")) {
                    string[] p_choice = ea_line.Split("::");
                    if (p_choice.Length == 2) {
                        string funcname = p_choice[0].Replace(" ", "");
                        string f_params = p_choice[1].Replace(" ", "");
                        param_choices.Add(funcname, f_params);
                    }
                }
            }

            // Loop through params and formula
            string new_formula = "=" + f + "(";
            int ad_row_count = 1;  // count the rows of array and dict to the right of func name
            string comma, param, def_value = "", choices = "";
            comma = ", ";
            Type param_type;
            for (int i = 1; i < p_len; i++) {
                param_type = param_info[i - 1].ParameterType;
                param = param_info[i - 1].Name;
                choices = param_choices.ContainsKey(param) ? param_choices[param] : "";
                try {
                    def_value = param_info[i - 1].HasDefaultValue ? param_info[i - 1].DefaultValue.ToString() : "";
                } catch (NullReferenceException) {
                    def_value = "";
                }
                if (param_type.Name.Contains("[]")) {
                    // array type
                    // title cell
                    ad_row_count += 1;
                    var param_cell = xlApp.Cells(y, x + ad_row_count);
                    ExManip.RangeEmpty(param_cell, backtrack);
                    param_cell.Value = param;
                    param_cell.Interior.Color = Color.FromArgb(60, 255, 255, 202);
                    param_cell.Borders.Color = Color.FromArgb(0, 0, 0, 0);
                    var array_cells = ExManip.GetRange(
                        y + 1, x + ad_row_count, y + 3, x + ad_row_count
                    );
                    // array cells
                    sheet.Columns(x + ad_row_count).ColumnWidth = 12;
                    new_formula += array_cells.Address + comma;
                    // grey out unused cell right to param name
                    xlApp.Cells(y + i, x + 1).Interior.Color = Color.FromArgb(0, 145, 145, 145);
                    def_value = "";
                    choices = "";
                } else if (param_type.Name.Contains("[,]")) {
                    // dict type
                    // title cells
                    ExManip.RangeEmpty(xlApp.Cells(y, x + ad_row_count + 1), backtrack);
                    ExManip.RangeEmpty(xlApp.Cells(y, x + ad_row_count + 2), backtrack);
                    xlApp.Cells(y, x + ad_row_count + 1).Value = param;
                    var param_cell = ExManip.GetRange(
                        y, x + ad_row_count + 1, y, x + ad_row_count + 2
                    );
                    param_cell.Merge();
                    param_cell.Interior.Color = Color.FromArgb(60, 255, 255, 202);
                    param_cell.Borders.Color = Color.FromArgb(0, 0, 0, 0);

                    // dict cells
                    var dict_cells = ExManip.GetRange(
                        y + 1, x + ad_row_count + 1,
                        y + 3, x + ad_row_count + 2
                    );
                    sheet.Columns(x + ad_row_count + 1).ColumnWidth = 12;
                    sheet.Columns(x + ad_row_count + 2).ColumnWidth = 12;
                    ad_row_count += 2;
                    new_formula += dict_cells.Address + comma;
                    // grey out unused cell right to param name
                    xlApp.Cells(y + i, x + 1).Interior.Color = Color.FromArgb(0, 145, 145, 145);
                    def_value = "";
                    choices = "";
                } else {
                    // bool, str, int, double types
                    new_formula += xlApp.Cells(y + i, x + 1).Address + comma;
                }
                ExManip.RangeEmpty(xlApp.Cells(y + i, x), backtrack);
                ExManip.RangeEmpty(xlApp.Cells(y + i, x + 1), backtrack);
                xlApp.Cells(y + i, x).Value = param;
                if (choices != "") {
                    xlApp.Cells(y + i, x + 1).Validation.Delete();
                    xlApp.Cells(y + i, x + 1).Validation.Add(
                        Excel.XlDVType.xlValidateList,
                        Excel.XlDVAlertStyle.xlValidAlertInformation,
                        Excel.XlFormatConditionOperator.xlBetween,
                        choices, Type.Missing
                    );
                    xlApp.Cells(y + i, x + 1).Validation.InCellDropdown = true;
                }
                if (def_value != "") {
                    xlApp.Cells(y + i, x + 1).Value = def_value;
                }
            }

            ExManip.RangeEmpty(xlApp.Cells(y + p_len + 1, x), backtrack);
            xlApp.Cells(y + p_len, x).Value = "return";
            dynamic param_name_range = ExManip.GetRange(y + 1, x, y + p_len, x);
            param_name_range.Interior.Color = Color.FromArgb(77, 241, 255, 205);

            dynamic nf_range = xlApp.Cells(y + p_len, x + 1);
            var rgx_param = new Regex(@", $");
            new_formula = rgx_param.Replace(new_formula, "");
            new_formula += ")";
            ExManip.RangeEmpty(nf_range, backtrack);

            sheet.Columns(x).Autofit();
            sheet.Columns(x + 1).Autofit();
            // border weight must be -4138 (just omit), 1, 2, 4
            ExManip.GetRange(y, x, y + p_len, x + 1).Borders.Color = Color.FromArgb(0, 0, 0, 0);

            // Set minimum column width
            if (sheet.Columns(x).ColumnWidth < 12) {
                sheet.Columns(x).ColumnWidth = 12;
            }
            if (sheet.Columns(x + 1).ColumnWidth < 12) {
                sheet.Columns(x + 1).ColumnWidth = 12;
            }

            // Set maximum column width
            if (sheet.Columns(x).ColumnWidth > 50) {
                sheet.Columns(x).ColumnWidth = 50;
            }
            if (sheet.Columns(x + 1).ColumnWidth > 50) {
                sheet.Columns(x + 1).ColumnWidth = 50;
            }
            nf_range.Formula2 = new_formula;
        }

        public static void AutoFuncClear()
        {
            // auto clear UDF
            if (!SheetExists()) {
                return;
            }

            int[] ac = ExManip.GetActiveCellPos();
            int y = ac[0];
            int x = ac[1];

            string f = FormulaExists(x, y);
            if (f == "") {
                return;
            }

            MethodInfo method_info = typeof(ExcelFunc).GetMethod(f);
            ParameterInfo[] param_info = method_info.GetParameters();

            // clear single cell parameters
            int p_size = param_info.Length;
            for (int i = 0; i < p_size + 2; i++) {
                for (int j = 0; j < 2; j++) {
                    ExcelFunc.ClearCell(new ExcelReference(y - i - 1, x - 1 - j));
                }
            }

            // parse out ranges of inputs from formula
            dynamic xlApp = ExcelDnaUtil.Application;
            var rgx_f = new Regex(@"[$]?[A-Z]+[$]?[0-9]+:[$]?[A-Z]+[$]?[0-9]+");
            MatchCollection pos_matches = rgx_f.Matches(old_formula);
            // in a val_ranges item, the key is [1,1] and value is [2,2] (if the range is R1C1:R2C2)
            var val_ranges = new Dictionary<int[], int[]>();
            foreach (Match m in pos_matches) {
                string ms = m.Value;
                string cell_addr = xlApp.Range(ms).Address(false, false, XlCall.xlcA1R1c1);
                string[] addrs = cell_addr.Split(':');
                int[] ac_0 = ExManip.GetCellPos(addrs[0]);
                int[] ac_1 = ExManip.GetCellPos(addrs[1]);
                val_ranges.Add(ac_0, ac_1);
            }

            // clear array and dict parameters
            int top_y = y - p_size - 2;
            int top_x = x - 1;
            foreach (var p in param_info) {
                Type t = p.ParameterType;
                if (t.Name.Contains("[]")) {
                    // array
                    top_x += 1;
                    ExcelFunc.ClearCell(new ExcelReference(top_y, top_x));
                    foreach (KeyValuePair<int[], int[]> item in val_ranges) {
                        if (item.Key[0] - 1 == top_y + 1 && item.Value[1] - 1 == top_x) {
                            ExcelFunc.ClearCell(new ExcelReference(
                                item.Key[0] - 1, item.Value[0] - 1, item.Key[1] - 1, item.Value[1] - 1
                            ));
                        }
                    }
                } else if (t.Name.Contains("[,]")) {
                    // dict
                    top_x += 1;
                    ExcelFunc.ClearCell(new ExcelReference(top_y, top_y, top_x, top_x + 1));
                    foreach (KeyValuePair<int[], int[]> item in val_ranges) {
                        if (item.Key[0] - 1 == top_y + 1 && item.Value[1]- 1 == top_x + 1) {
                            ExcelFunc.ClearCell(new ExcelReference(
                                item.Key[0] - 1, item.Value[0] - 1, item.Key[1] - 1, item.Value[1] - 1
                            ));
                        }
                    }
                    top_x += 1;
                }
            }
        }
    }
    // END: public static class AutoFill


    public static class ExManip
    {
        static PyExecutor pye = new PyExecutor();

        public static void RangeEmpty(dynamic range, List<dynamic> bt)
        {
            // if cells in range is not empty, undo all the changes by the func formatter
            bool cleanup = false;
            if (range.Value != null) {
                if (range.Value.GetType().IsArray) {
                    foreach (var v in range.Value) {
                        if (v != null) {
                            cleanup = true;
                        }
                    }
                } else { cleanup = true; }
            }

            if (cleanup) {
                foreach (dynamic ea_range in bt) {
                    ea_range.UnMerge();
                    ea_range.Clear();
                }
                bt[0].Value = AutoFill.old_formula;
                string errmsg = "Cannot overwrite non-empty cell(s): " + range.Address;
                WriteLog(errmsg, "WARNING");
                throw new ApplicationException(errmsg);
            }
            bt.Add(range);
        }

        public static dynamic GetRange(int y1, int x1, int y2, int x2)
        {
            // get the address of a range of cells
            dynamic xlApp = ExcelDnaUtil.Application;
            return xlApp.Range(xlApp.Cells(y1, x1), xlApp.Cells(y2, x2));
        }

        public static void WriteLog(string logmsg, string level)
        {
            // use python logging
            pye.qxlpyLogMessage(logmsg, level);
        }

        public static int[] GetCellPos(string cell_addr)
        {
            // Range A1 = RC, A2 = RC[1], B1 = R[1]C, B2 = R[1]C[1] ...
            // Cells A1 = 1, 1
            // Cells(Row, Column)
            var rgx_x = new Regex(@"(?<=.+C\[)[0-9]+");
            var rgx_y = new Regex(@"(?<=^R\[)[0-9]+");
            Match match_x = rgx_x.Match(cell_addr);
            Match match_y = rgx_y.Match(cell_addr);
            int x = match_x.Success ? int.Parse(match_x.Value)+1 : 1;
            int y = match_y.Success ? int.Parse(match_y.Value)+1 : 1;

            return new int[] {y, x};
        }

        public static int[] GetActiveCellPos()
        {
            // Get numeric cell address
            dynamic xlApp = ExcelDnaUtil.Application;

            // RomAbsolute=false, ColumnAbsolute=false, AddressReference
            string cell_addr = xlApp.ActiveCell.Address(false, false, XlCall.xlcA1R1c1);
            int[] ac = ExManip.GetCellPos(cell_addr);
            return ac;
        }
    }
    // END: public static class ExManip


    public static class ExcelFunc
    {
        public static void ClearCell(ExcelReference ex_ref)
        {
            ExcelAsyncUtil.QueueAsMacro(() => {
                XlCall.Excel(XlCall.xlcSelect, ex_ref);
                XlCall.Excel(XlCall.xlcClear, 1);
            });
        }

        private static object CheckEmpty(object obj, object? defval = null, bool assert = false)
        {
            string o = obj.ToString();
            if (String.IsNullOrEmpty(o) || o == "ExcelDna.Integration.ExcelEmpty" || o == "ExcelDna.Integration.ExcelMissing") {
                if (assert) {
                    string warning_msg = "Missing Argument";
                    ExManip.WriteLog(warning_msg, "WARNING");
                    throw new ArgumentNullException(warning_msg);
                }
                if (defval == null) {
                    return "";
                }
                return defval;
            }
            return obj;
        }

        private static object[] ListCheckEmpty(object[] obj)
        {
            var ret = new List<object>();
            foreach (object o in obj) {
                ret.Add(CheckEmpty(o));
            }

            // return empty list if all items are empty
            bool empty_list = true;
            foreach (object o in ret) {
                if (!String.IsNullOrEmpty(o.ToString())) {
                    empty_list = false;
                    break;
                }
            }
            if (empty_list) { ret.Clear(); }
            return ret.ToArray();
        }

        private static object[,] DictCheckEmpty(object[,] obj)
        {
            int nested_len = obj.GetLength(1);
            int len = obj.GetLength(0);
            object[,] ret = new object[len, nested_len];
            bool empty_list = true;
            for (int i = 0; i < len; i++) {
                for (int j = 0; j < nested_len; j++) {
                    ret[i, j] = CheckEmpty(obj[i, j]);
                    if (empty_list == true && !String.IsNullOrEmpty(ret[i, j].ToString())) {
                        empty_list = false;
                    }
                }
            }
            if (empty_list == true) {
                ret = new object[0, 0];
            }
            return ret;
        }

        [ExcelCommand(Name = "autoformat", ShortCut = "^{INSERT}")]
        public static void AutoFormat()
        {
            AutoFill.AutoFuncFormat();
        }

        [ExcelCommand(Name = "funcclear", ShortCut = "^{DELETE}")]
        public static void FuncClear()
        {
            AutoFill.AutoFuncClear();
        }

        static PyExecutor pye = new PyExecutor();

        [ExcelFunction(Name = "qxlpyReloadPyModules")]
        public static string qxlpyReloadPyModules()
        {
            pye.qxlpyReloadPyModules();
            return "SUCCESS";
        }

        // THE FOLLOWING FUNCTIONS ARE GENERATED BY CS_AUTOGEN //
_BODY_
    }
    // END: public static class ExcelFunc
}

'''
MAIN_EXCEL = '        [ExcelFunction(Name = "_FUNCTIONNAME_")]'
MAIN_F = '        public static _EXCELRETURNTYPE_ _FUNCTIONNAME_(_PARAMETERS_)'
MAIN_RET_PYE = '            _PYRETURNTYPE_ ret = pye._FUNCTIONNAME_(_ARGS_);'
MAIN_RETURN_S = '            return _RET_;'
MAIN_DOCSTRING = '''                {
                    "_FUNCTIONNAME_",
@"_DOCSTRING_"
                },
'''
### main.cs string templates ENDS ###


### python.cs string templates ###
PYTHON_BODY = r'''using Python.Runtime;
using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace qxlpy
{
    public class PyExecutor
    {
        private readonly PyModule SCOPE;
        private readonly dynamic SITE, _PYTHONMODS_;

        public PyExecutor()
        {
            string root = Environment.GetEnvironmentVariable("QXLPYDIR");

            string python_dll = $@"{root}\python\python311.dll";
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", python_dll);

            PythonEngine.Initialize();
            SCOPE = Py.CreateScope();

            using (Py.GIL())
            {
                SITE = SCOPE.Import("site");
                SITE.addsitedir(root);
_PYIMPORTLIST_
            }
        }

        private dynamic GetToTypeByValue (object obj)
        {
            // Return value with the desired type by type hierarchy
            // Type hierarchy: int, float, string
            string value_str = obj.ToString();
            dynamic totype;
            if (Regex.IsMatch(value_str, @"^[0-9]+$")) {
                totype = Convert.ToInt64(value_str);
            }
            else if (Regex.IsMatch(value_str, @"^[0-9]+\.[0-9]+$")) {
                totype = Convert.ToDouble(value_str);
            }
            else if (Regex.IsMatch(value_str, @"^[0-9.]+E[-+][0-9]+$")) {
                totype = float.Parse(value_str, NumberStyles.Float);
            }
            else {
                totype = value_str;
            }
            return totype;
        }

        private dynamic GetPyTypeByValue (object obj)
        {
            // Return the desired PyType
            string value_str = obj.ToString();
            dynamic pytype;
            if (Regex.IsMatch(value_str, @"^[0-9]+$")) {
                pytype = new PyInt(Convert.ToInt64(value_str));
            }
            else if (Regex.IsMatch(value_str, @"^[0-9]+\.[0-9]+$")) {
                pytype = new PyFloat(Convert.ToDouble(value_str));
            }
            else if (Regex.IsMatch(value_str, @"^[0-9.]+E[-+][0-9]+$")) {
                pytype = new PyFloat(float.Parse(value_str, NumberStyles.Float));
            }
            else {
                pytype = new PyString(value_str);
            }
            return pytype;
        }

        public void qxlpyReloadPyModules()
        {
            using (Py.GIL())
            {
                dynamic importlib = SCOPE.Import("importlib");
_PYRELOADLIST_
                return;
            }
        }

        // THE FOLLOWING FUNCTIONS WILL BE AUTOGEN //
_BODY_
    }
}
'''
PYTHON_GIL = r'''
        {
            using (Py.GIL())
            {_DLINPUTS_
_BODY_
_DLRETURN_
                return ret;
            }
        }
'''
PYTHON_FUNC = '        public _FUNCTYPE_ _FUNCTIONNAME_(_PARAMETERS_)'
PYTHON_CALL = '                _FUNCTYPE_ ret = _PYTHONIMPORT_._FUNCTIONNAME_(_ARGS_);'
PYTHON_LIST_RETURN = r'''
                PyList pylist_ret = _PYTHONIMPORT_._FUNCNAME_(_PYPARAMS_);
                long len = pylist_ret.Length();
                var ret = new object[len, 1];
                int row = 0;
                foreach (PyObject pyobj in pylist_ret) {
                    ret[row, 0] = _PYTYPE_;
                    row += 1;
                }
'''
PYTHON_DICT_RETURN = r'''
                PyDict pydict_ret = _PYTHONIMPORT_._FUNCNAME_(_PYPARAMS_);
                var ret = new object[pydict_ret.Length(), 2];
                int row = 0;
                foreach (PyObject key in pydict_ret) {
                    ret[row, 0] = _KEYPYTYPE_;
                    ret[row, 1] = _VALPYTYPE_;
                    row += 1;
                }
'''
PYTHON_NESTED_LIST_RETURN = r'''
                PyList pylist_ret = _PYTHONIMPORT_._FUNCNAME_(_PYPARAMS_);
                long row_len = pylist_ret.Length();
                long col_len = 0;
                foreach (PyObject pyobj in pylist_ret) {
                    col_len = pyobj.Length();
                    break;
                }
                var ret = new object[row_len, col_len];

                int row = 0;
                foreach (PyObject pyobj in pylist_ret) {
                    PyList pylist = PyList.AsList(pyobj);
                    int col = 0;
                    foreach (PyObject internal_pyobj in pylist) {
                        ret[row, col] = _PYTYPE_;
                        col += 1;
                    }
                    row += 1;
                }
'''
PYTHON_LIST_INPUT = r'''
                var pylist__ARGNAME_ = new PyList();
                foreach (object n in _ARGNAME_) {
                    object ea_obj = n;
                    _ARGTYPE_ obj__ARGNAME_;
                    try {
                        obj__ARGNAME_ = _TOTYPE_(ea_obj);
                    } catch (Exception e) {
                        string error_msg = $"Wrong type in array: '{Convert.ToString(n)}' is not of type '_ARGTYPE_'";
                        qxlpyLogMessage(error_msg, "ERROR");
                        throw new ArrayTypeMismatchException(error_msg);
                    }
                    pylist__ARGNAME_.Append(_PYTYPE_(obj__ARGNAME_));
                }
'''
PYTHON_DICT_INPUT = r'''
                var pydict__ARGNAME_ = new PyDict();
                for (int i = 0; i < _ARGNAME_.GetLength(0); i++) {
                    _KEYTYPE_ k__ARGNAME_;
                    _VALTYPE_ v__ARGNAME_;
                    try {
                        object objkey = _ARGNAME_[i, 0];
                        object objval = _ARGNAME_[i, 1];
                        if (Convert.ToString(objkey) == "") {
                            continue;
                        }
                        k__ARGNAME_ = _TOKEYTYPE_(objkey);
                        v__ARGNAME_ = _TOVALTYPE_(objval);
                    } catch (Exception e) {
                        string error_msg = $"Wrong type in dictionary: ";
                        error_msg += "'{Convert.ToString(objkey)}' should be '_KEYTYPE_' and ";
                        error_msg += "'{Convert.ToString(objval)}' should be '_VALTYPE_'";
                        qxlpyLogMessage(error_msg, "ERROR");
                        throw new ArrayTypeMismatchException(error_msg);
                    }
                    pydict__ARGNAME_[k__ARGNAME_] = _PYTYPEVAL_(v__ARGNAME_);
                }
'''
PYTHON_NESTED_LIST_INPUT = r'''
                var pylist__ARGNAME_ = new PyList();
                for (int i = 0; i < _ARGNAME_.GetLength(0); i++) {
                    var internal__ARGNAME_ = new PyList();
                    for (int j = 0; j < _ARGNAME_.GetLength(1); j++) {
                        object ea_obj = _ARGNAME_[i, j];
                        _ARGTYPE_ obj__ARGNAME_;
                        try {
                            obj__ARGNAME_ = _TOTYPE_(ea_obj);
                        } catch (Exception e) {
                            string error_msg = $"Wrong type in array: '{Convert.ToString(ea_obj)}' is not of type '_ARGTYPE_'";
                            qxlpyLogMessage(error_msg, "ERROR");
                            throw new ArrayTypeMismatchException(error_msg);
                        }
                        internal__ARGNAME_.Append(_PYTYPE_(obj__ARGNAME_));
                    }
                    pylist__ARGNAME_.Append(internal__ARGNAME_);
                }
'''
### python.cs string templates ENDS ###
