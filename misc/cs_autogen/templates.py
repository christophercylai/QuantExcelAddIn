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
            <group id='qxlpy' label='QXLPY'>
              <button id='expandfunc' label='Expand Function'
                onAction='expandFuncButton' size='large' screentip='Expand Function (Ctrl-INS)'
                imageMso='ConditionalFormattingColorScalesGallery' />
              <button id='removefunc' label='Remove Function'
                onAction='removeFuncButton' size='large' screentip='Remove Function (Ctrl-Shift-DEL)'
                imageMso='TableDelete' />
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
            AutoFill.AutoDataClear();
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
            int p_len = param_info.Length;

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
            xlApp.Cells(y, x).AddComment(GetComment(f));
            ExManip.GetRange(y, x, y, x +1).Merge();

            // Loop through params and formula
            string new_formula = "=" + f + "(";
            int ad_row_count = 1;  // count the rows of array and dict to the right of func name
            string comma, param, def_value;
            comma = ", ";
            Type param_type;
            for (int i = 1; i < p_len; i++) {
                param_type = param_info[i - 1].ParameterType;
                param = param_info[i - 1].Name;
                def_value = param_info[i - 1].HasDefaultValue ? param_info[i - 1].DefaultValue.ToString() : "";
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
                } else {
                    // bool, str, int, double types
                    new_formula += xlApp.Cells(y + i, x + 1).Address + comma;
                }
                ExManip.RangeEmpty(xlApp.Cells(y + i, x), backtrack);
                ExManip.RangeEmpty(xlApp.Cells(y + i, x + 1), backtrack);
                xlApp.Cells(y + i, x).Value = param;
                if (def_value != "") {
                    xlApp.Cells(y + i, x + 1).Value = def_value;
                }
            }

            ExManip.RangeEmpty(xlApp.Cells(y + p_len + 1, x), backtrack);
            xlApp.Cells(y + p_len, x).Value = "return";
            dynamic param_name_range = ExManip.GetRange(y + 1, x, y + p_len, x);
            param_name_range.Interior.Color = Color.FromArgb(77, 241, 255, 205);

            dynamic nf_range = xlApp.Cells(y + p_len, x + 1);
            new_formula += "CELL(\"address\", " + nf_range.Address.Replace("$", "") + "))";
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
            nf_range.Value = new_formula;
        }

        public static void AutoDataClear()
        {
            // auto clear data from ExcelFunc's UDF
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

            ExcelFunc.ClearData(x, y);
            ExcelFunc.ClearData(x - 1, y);
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
            int p_size = param_info.Length - 1;
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
                WriteLog(errmsg, "ERROR");
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
            PyExecutor pye = new();
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

        public static void ClearData(int x, int y) {
            // clean up old data
            dynamic xlApp = ExcelDnaUtil.Application;
            bool data_formatted = true;
            int cell_count = 0;
            while (data_formatted) {
                dynamic addr = xlApp.Cells(y + cell_count + 1, x);
                if (addr.Font.Name == "Courier New" &&
                    addr.Font.Size == 9 &&
                    addr.Font.Italic)
                {
                    ClearCell(new ExcelReference(y + cell_count, x - 1));
                } else {
                    data_formatted = false;
                }
                cell_count += 1;
            }
        }

        private static void CheckEmpty(object obj)
        {
            string o = obj.ToString();
            if (String.IsNullOrEmpty(o) || o == "ExcelDna.Integration.ExcelEmpty" || o == "ExcelDna.Integration.ExcelMissing") {
                throw new ArgumentNullException("Missing Arguments");
            }
        }

        private static void ListCheckEmpty(object[] obj)
        {
            foreach (object o in obj) {
                CheckEmpty(o);
            }
        }

        private static void DictCheckEmpty(object[,] obj)
        {
            foreach (object o in obj) {
                CheckEmpty(o);
            }
        }

        private static void FillDataCell(ExcelReference ex_ref, object obj)
        {
            ExcelAsyncUtil.QueueAsMacro(() => {
                // select ex_ref as active
                XlCall.Excel(XlCall.xlcSelect, ex_ref);
                XlCall.Excel(XlCall.xlcFormatFont, "Courier New", 9, false, true);
                XlCall.Excel(XlCall.xlcPatterns, 1, 35, 1);
                XlCall.Excel(XlCall.xlcBorder, 1);
                ex_ref.SetValue(obj);
            });
        }

        [ExcelCommand(Name = "autoformat", ShortCut = "^{INSERT}")]
        public static void AutoFormat()
        {
            AutoFill.AutoFuncFormat();
        }

        [ExcelCommand(Name = "dataclear", ShortCut = "^{DELETE}")]
        public static void DataClear()
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic r = xlApp.ActiveCell;
            AutoFill.AutoDataClear();
            // re-focus on the formula cell
            r.Activate();
        }

        [ExcelCommand(Name = "allclear", ShortCut = "^+{DELETE}")]
        public static void AllClear()
        {
            AutoFill.AutoDataClear();
            AutoFill.AutoFuncClear();
        }

        [ExcelCommand(Name = "funcclear", ShortCut = "+{DELETE}")]
        public static void FuncClear()
        {
            AutoFill.AutoFuncClear();
        }

        // THE FOLLOWING FUNCTIONS ARE GENERATED BY CS_AUTOGEN //
_BODY_
    }
    // END: public static class ExcelFunc
}

'''
MAIN_EXCEL = '        [ExcelFunction(Name = "_FUNCTION_NAME_")]'
MAIN_F = '        public static _EXCEL_RETURN_TYPE_ _FUNCTION_NAME_(_PARAMETERS_)'
MAIN_RET_PYE = '            _PY_RETURN_TYPE_ ret = pye._FUNCTION_NAME_(_ARGS_);'
MAIN_RETURN_S = '            return _RET_;'
MAIN_DOCSTRING = '''                {
                    "_FUNCTION_NAME_",
@"_DOCSTRING_"
                },
'''
MAIN_LIST = r'''
            int len = ret.Length;
            if (len == 0) { return "N/A"; }

            dynamic xlApp = ExcelDnaUtil.Application;
            string cell_addr = xlApp.Range(func_pos).Address(false, false, XlCall.xlcA1R1c1);
            int[] ac = ExManip.GetCellPos(cell_addr);
            int y = ac[0];
            int x = ac[1];

            // check empty cells
            for (int i = 0; i < len; i++) {
                var x_ref = new ExcelReference(y + i, x - 1);
                string cell_location = x_ref.GetValue().ToString();
                if (cell_location != "ExcelDna.Integration.ExcelEmpty") {
                    string addr = xlApp.Cells(y + i + 1, x).Address;
                    string errmsg = $@"Cannot overwrite non-empty cell(cell_location): {addr}";
                    pye.qxlpyLogMessage(errmsg, "WARNING");
                    return errmsg;
                }
            }

            // fill values
            for (int i = 0; i < len; i++) {
                var ex_ref = new ExcelReference(y + i, x - 1);
                FillDataCell(ex_ref, ret[i]);
            }
'''
MAIN_DICT = r'''
            object[][] kv_pair = {
                ret[0].ToArray(),
                ret[1].ToArray()
            };
            int len = kv_pair[0].Length;
            if (len == 0) { return "N/A"; }

            dynamic xlApp = ExcelDnaUtil.Application;
            string cell_addr = xlApp.Range(func_pos).Address(false, false, XlCall.xlcA1R1c1);
            int[] ac = ExManip.GetCellPos(cell_addr);
            int y = ac[0];
            int x = ac[1];

            // check empty cells
            for (int i = 0; i < len; i++) {
                for (int j = 0; j < 2; j++) {
                    var x_ref = new ExcelReference(y + i, x - 2 + j);
                    string cell_location = x_ref.GetValue().ToString();
                    if (cell_location != "ExcelDna.Integration.ExcelEmpty") {
                        string errmsg = "Cannot overwrite non-empty cell(s): " + xlApp.Cells(y + i + 1, x - 1 + j).Address;
                        pye.qxlpyLogMessage(errmsg, "WARNING");
                        return errmsg;
                    }
                }
            }

            // fill values
            for (int i = 0; i < len; i++) {
                for (int j = 0; j < 2; j++) {
                    var ex_ref = new ExcelReference(y + i, x - 2 + j);
                    FillDataCell(ex_ref, kv_pair[j][i]);
                }
            }
'''
### main.cs string templates ENDS ###


### python.cs string templates ###
PYTHON_BODY = r'''
using Python.Runtime;


namespace qxlpy
{
    public class PyExecutor
    {
        private readonly PyModule SCOPE;

        public PyExecutor()
        {
            string root = Environment.GetEnvironmentVariable("QXLPYDIR");

            string python_dll = $@"{root}\python\python311.dll";
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", python_dll);

            PythonEngine.Initialize();
            SCOPE = Py.CreateScope();

            using (Py.GIL())
            {
                dynamic site = SCOPE.Import("site");
                site.addsitedir(root);
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
            {_DL_INPUTS_
_BODY_
_DL_RETURN_
                return ret;
            }
        }
'''
PYTHON_FUNC = '        public _FUNC_TYPE_ _FUNCTION_NAME_(_PARAMETERS_)'
PYTHON_IPT = '                dynamic imp = SCOPE.Import("quant._MODULE_NAME_");'
PYTHON_CALL = '                _FUNC_TYPE_ ret = imp._FUNCTION_NAME_(_ARGS_);'
PYTHON_LIST_RETURN = r'''
                var ret_list = new List<_LIST_TYPE_>();
                PyList pylist_ret = imp._FUNC_NAME_(_PY_PARAMS_);
                foreach (PyObject pyobj in pylist_ret) {
                    ret_list.Add(Convert._TO_TYPE_(pyobj));
                }
                object[] ret = ret_list.ToArray();
'''
PYTHON_DICT_RETURN = r'''
                var ret = new List<List<object>>();
                var keys = new List<object>();
                var values = new List<object>();
                PyDict pydict_ret = imp._FUNC_NAME_(_PY_PARAMS_);
                foreach (PyObject key in pydict_ret) {
                    keys.Add(Convert._TO_KEY_TYPE_(key));
                    values.Add(Convert._TO_VAL_TYPE_(pydict_ret.GetItem(key)));
                }
                ret.Add(keys);
                ret.Add(values);
'''
PYTHON_LIST_INPUT = r'''
                var pylist__ARG_NAME_ = new PyList();
                foreach (object n in _ARG_NAME_) {
                    _ARG_TYPE_ obj__ARG_NAME_;
                    try {
                        obj__ARG_NAME_ = Convert._TO_TYPE_(n);
                    } catch (Exception e) {
                        string error_msg = $"Wrong type in array: '{Convert.ToString(n)}' is not of type '_ARG_TYPE_'";
                        qxlpyLogMessage(error_msg, "ERROR");
                        throw new ArrayTypeMismatchException(error_msg);
                    }
                    pylist__ARG_NAME_.Append(new _PY_TYPE_(obj__ARG_NAME_));
                }
'''
PYTHON_DICT_INPUT = r'''
                var pydict__ARG_NAME_ = new PyDict();
                for (int i = 0; i < _ARG_NAME_.GetLength(0); i++) {
                    _KEY_TYPE_ k__ARG_NAME_;
                    _VAL_TYPE_ v__ARG_NAME_;
                    try {
                        k__ARG_NAME_ = Convert._TO_KEYTYPE_(_ARG_NAME_[i, 0]);
                        v__ARG_NAME_ = Convert._TO_VALTYPE_(_ARG_NAME_[i, 1]);
                    } catch (Exception e) {
                        string error_msg = $"Wrong type in dictionary: ";
                        error_msg += "'{Convert.ToString(k__ARG_NAME_)}' should be '_KEY_TYPE_' and ";
                        error_msg += "'{Convert.ToString(v__ARG_NAME_)}' should be '_VAL_TYPE_'";
                        qxlpyLogMessage(error_msg, "ERROR");
                        throw new ArrayTypeMismatchException(error_msg);
                    }
                    pydict__ARG_NAME_[k__ARG_NAME_.ToString()] = new _PY_TYPE_VAL_(v__ARG_NAME_);
                }
'''
### python.cs string templates ENDS ###
