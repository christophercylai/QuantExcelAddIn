using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Reflection;


namespace qxlpy

{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            // note: onLoad option can be added after customUI to run a method when the ribbon loads
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
      <ribbon>
        <tabs>
          <tab id='qxltab' label='QXLPY'>
            <group id='qxlpy' label='QXLPY'>
              <button id='expandfunc' label='Expand Function'
                onAction='OnButtonPressed' size='large'
                imageMso='ConditionalFormattingColorScalesGallery' />
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            AutoFill.AutoFuncFormat();
        }
    }
    // END: public class RibbonController : ExcelRibbon


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
            var rgx_f = new Regex(@"[a-zA-Z][a-zA-Z0-9]+");
            Match match_f = rgx_f.Match(old_formula);

            if (!match_f.Success) {
                ExManip.WriteLog("Formula must start with [a-zA-Z] and followed by [a-zA-Z0-9]+", "WARNING");
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

            // Check whether formula is a method of ExcelFunc
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
            xlApp.Cells(y, x).Interior.Color = Color.FromArgb(142, 0, 111, 41);
            ExManip.GetRange(y, x, y, x +1).Merge();

            // Loop through params and formula
            string new_formula = "=" + f + "(";
            int ad_row_count = 1;  // count the rows of array and dict to the right of func name
            string comma, param, def_value;
            Type param_type;
            for (int i = 1; i <= p_len; i++) {
                param_type = param_info[i - 1].ParameterType;
                comma = i == p_len ? "" : ", ";
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
                    ExManip.RangeEmpty(array_cells, backtrack);
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
                    ExManip.RangeEmpty(dict_cells, backtrack);
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
            xlApp.Cells(y + p_len + 1, x).Value = "return";
            dynamic param_name_range = ExManip.GetRange(y + 1, x, y + p_len + 1, x);
            param_name_range.Interior.Color = Color.FromArgb(77, 241, 255, 205);
            new_formula += ")";
            ExManip.RangeEmpty(xlApp.Cells(y + p_len + 1, x + 1), backtrack);
            xlApp.Cells(y + p_len + 1, x + 1).Value = new_formula;

            sheet.Columns(x).Autofit();
            sheet.Columns(x + 1).Autofit();
            // border weight must be -4138 (just omit), 1, 2, 4
            ExManip.GetRange(y, x, y + p_len + 1, x + 1).Borders.Color = Color.FromArgb(0, 0, 0, 0);

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
        }

        public static void AutoFuncClear() {
            // auto format a UDF from ExcelFunc
            dynamic xlApp = ExcelDnaUtil.Application;
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

        }
    }
    // END: public static clase AutoFill


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
                throw new ApplicationException("Cannot overwrite non-empty cell(s): " + range.Address);

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
            pye.LogMessage(logmsg, level);
        }

        public static int[] GetActiveCellPos()
        {
            // Get numeric cell address
            dynamic xlApp = ExcelDnaUtil.Application;

            // RomAbsolute=false, ColumnAbsolute=false, AddressReference
            string cell_addr = xlApp.ActiveCell.Address(false, false, XlCall.xlcA1R1c1);
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
    }
    // END: public static class ExManip


    public static class ExcelFunc
    {
        private static void CheckType(object obj)
        {
            if (obj.ToString() == "") {
                throw new ArgumentNullException("Missing Arguments");
            }
        }

        private static void ListCheckType(object[] obj)
        {
            foreach (object o in obj) {
                CheckType(o);
            }
        }

        private static void DictCheckType(object[,] obj)
        {
            foreach (object o in obj) {
                CheckType(o);
            }
        }

        [ExcelCommand(Name = "autoformat", ShortCut = "^{INSERT}")]
        public static void AutoFormat()
        {
            AutoFill.AutoFuncFormat();
        }

        [ExcelCommand(Name = "autoclear", ShortCut = "^{DELETE}")]
        public static void AutoClear()
        {
            AutoFill.AutoFuncClear();
        }

        [ExcelFunction(Name = "QxlpyGetPath")]
        public static string QxlpyGetPath()
        {
            PyExecutor pye = new();
            string path = pye.GetPath();
            return path;
        }

        // THE FOLLOWING FUNCTIONS WILL BE AUTOGEN //

        [ExcelFunction(Name = "QxlpyLogMessage")]
        public static string QxlpyLogMessage(string logmsg, string level = "INFO")
        {
            CheckType(logmsg);
            CheckType(level);
            PyExecutor pye = new();
            string ret = pye.LogMessage(logmsg, level);
            return ret;
        }

        [ExcelFunction(Name = "QxlpyGetCalculate")]
        public static string QxlpyGetCalculate(object[] objlist)
        {
            ListCheckType(objlist);
            PyExecutor pye = new();
            string ret = pye.GetCalculate(objlist);
            return ret;
        }

        [ExcelFunction(Name = "QxlpyCalculateAddNum")]
        public static double QxlpyCalculateAddNum(string addr)
        {
            CheckType(addr);
            PyExecutor pye = new();
            double ret = pye.CalculateAddNum(addr);
            return ret;
        }

        [ExcelFunction(Name = "QxlpyStoreStrDict")]
        public static string QxlpyStoreStrDict(object[,] objdict)
        {
            DictCheckType(objdict);
            PyExecutor pye = new();
            string ret = pye.StoreStrDict(objdict);
            return ret;
        }

        [ExcelFunction(Name = "QxlpyListGlobalObjects")]
        public static string QxlpyListGlobalObjects()
        {
            PyExecutor pye = new();
            object[] ret = pye.ListGlobalObjects().ToArray();
            int len = ret.Length;
            if (len == 0) { return "N/A"; }

            int[] ac = ExManip.GetActiveCellPos();
            int y = ac[0];
            int x = ac[1];

            dynamic xlApp = ExcelDnaUtil.Application;
            for (int i = 0; i < len; i++) {
                var ex_ref = new ExcelReference(y + i, x - 1);
                string s = ex_ref.GetValue().ToString();
                if (s != "ExcelDna.Integration.ExcelEmpty") {
                    string errmsg = "Cannot overwrite non-empty cell(s): " + xlApp.Cells(y + i + 1, x).Address;
                    pye.LogMessage(errmsg, "WARNING");
                    return errmsg;
                }
                s = ret[i].ToString();
                ExcelAsyncUtil.QueueAsMacro(() => { ex_ref.SetValue(s); });
            }
            return "SUCCESS";
        }

        [ExcelFunction(Name = "QxlpyGetStrDict")]
        public static string QxlpyGetStrDict(string obj_name)
        {
            CheckType(obj_name);
            PyExecutor pye = new();
            Dictionary<string, List<string>> ret = pye.GetStrDict(obj_name);
            string[][] kv_pair = {
                ret["keys"].ToArray(),
                ret["values"].ToArray()
            };
            int len = kv_pair[0].Length;
            if (len == 0) { return "N/A"; }

            int[] ac = ExManip.GetActiveCellPos();
            int y = ac[0];
            int x = ac[1];

            dynamic xlApp = ExcelDnaUtil.Application;
            for (int i = 0; i < len; i++) {
                for (int j = 0; j < 2; j++) {
                    var ex_ref = new ExcelReference(y + i, x - 2 + j);
                    string s = ex_ref.GetValue().ToString();
                    if (s != "ExcelDna.Integration.ExcelEmpty") {
                        string errmsg = "Cannot overwrite non-empty cell(s): " + xlApp.Cells(y + i + 1, x - 1 + j).Address;
                        pye.LogMessage(errmsg, "WARNING");
                        return errmsg;
                    }
                    s = kv_pair[j][i].ToString();
                    ExcelAsyncUtil.QueueAsMacro(() => { ex_ref.SetValue(s); });
                }
            }
            return "SUCCESS";
        }

        [ExcelFunction(Name = "QxlpyObjectExists")]
        public static bool QxlpyObjectExists(string obj_name)
        {
            CheckType(obj_name);
            PyExecutor pye = new();
            bool ret = pye.ObjectExists(obj_name);
            return ret;
        }
    }
    // END: public static class ExcelFunc
}
