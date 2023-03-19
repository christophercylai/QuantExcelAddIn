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
            dynamic xlApp = ExcelDnaUtil.Application;
            // Check if there is an active worksheet
            if (xlApp.ActiveSheet == null) { return; }

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

            // Check if the cell has a formula
            if (!xlApp.Cells(y, x).HasFormula) { return; }

            // Get formula name
            string cell_formula = xlApp.Cells(y, x).Formula;
            var rgx_f = new Regex(@"[a-zA-Z][a-zA-Z0-9]+");
            Match match_f = rgx_f.Match(cell_formula);

            if (!match_f.Success) {
                return;
            }
            string f = match_f.Value;

            // Check whether formula is a method of ExcelFunc
            MethodInfo method_info = typeof(ExcelFunc).GetMethod(f);
            if (method_info == null) { return; }
            ParameterInfo[] param_info = method_info.GetParameters();
            int p_len = param_info.Length;

            // Cells formatting
            // Title = function name
            xlApp.Cells(y, x).Value = f;
            xlApp.Cells(y, x).Interior.Color = Color.FromArgb(142, 0, 111, 41);
            xlApp.Range(xlApp.Cells(y, x), xlApp.Cells(y, x + 1)).Merge();

            // Loop through params and formula
            string new_formula = "=" + f + "(";
            int ad_row_count = 1;  // count the rows of array and dict to the right of func name
            string comma, param;
            Type param_type;
            for (int i = 1; i <= p_len; i++) {
                param_type = param_info[i - 1].ParameterType;
                comma = i == p_len ? "" : ", ";
                param = param_info[i - 1].Name;
                if (param_type.Name.Contains("[]")) {
                    // array type
                    ad_row_count += 1;
                    var param_cell = xlApp.Cells(y, x + ad_row_count);
                    param_cell.Value = param;
                    param_cell.Interior.Color = Color.FromArgb(60, 255, 255, 202);
                    param_cell.Borders.Color = Color.FromArgb(0, 0, 0, 0);
                    var array_cells = xlApp.Range(
                        xlApp.Cells(y + 1, x + ad_row_count),
                        xlApp.Cells(y + 3, x + ad_row_count)
                    );
                    new_formula += array_cells.Address + comma;
                    xlApp.Cells(y + i, x + 1).Interior.Color = Color.FromArgb(0, 145, 145, 145);
                } else {
                    // str, int, double types
                    new_formula += xlApp.Cells(y + i, x + 1).Address + comma;
                }
                xlApp.Cells(y + i, x).Value = param;
            }
            xlApp.Cells(y + p_len + 1, x).Value = "return";
            xlApp.Range(xlApp.Cells(y + 1, x), xlApp.Cells(y + p_len + 1, x)).Interior.Color = Color.FromArgb(77, 241, 255, 205);
            new_formula += ")";
            xlApp.Cells(y + p_len + 1, x + 1).Value = new_formula;

            xlApp.Range(xlApp.Cells(y, x), xlApp.Cells(y + p_len + 1, x + 1)).Columns.Autofit();
            // border weight must be -4138 (just omit), 1, 2, 4
            xlApp.Range(xlApp.Cells(y, x), xlApp.Cells(y + p_len + 1, x + 1)).Borders.Color = Color.FromArgb(0, 0, 0, 0);

            // Set minimum column width
            if (xlApp.Cells(y, x).ColumnWidth < 12) {
                xlApp.Cells(y, x).ColumnWidth = 12;
            }
            if (xlApp.Cells(y, x+1).ColumnWidth < 12) {
                xlApp.Cells(y, x+1).ColumnWidth = 12;
            }

            // Set maximum column width
            if (xlApp.Cells(y, x).ColumnWidth > 50) {
                xlApp.Cells(y, x).ColumnWidth = 50;
            }
            if (xlApp.Cells(y, x+1).ColumnWidth > 50) {
                xlApp.Cells(y, x+1).ColumnWidth = 50;
            }
        }
    }

    public static class ExcelFunc
    {
        [ExcelFunction(Name = "QxlpyGetPath")]
        public static string QxlpyGetPath()
        {
            PyExecutor pye = new();
            string path = pye.GetPath();
            return path;
        }

        [ExcelFunction(Name = "QxlpyLogMessage")]
        public static string QxlpyLogMessage(string log_msg, string level)
        {
            if (log_msg == "" || level == "") {
                throw new ArgumentNullException("Missing Arguments");
            }
            PyExecutor pye = new();
            pye.LogMessage(log_msg, level);
            string ret = "'" + log_msg + "'" + " is written on Logs/qxlcs.log";
            return ret;
        }

        // THE FOLLOWING FUNCTIONS WILL BE AUTOGEN //

        [ExcelFunction(Name = "QxlpyGetCalculate")]
        public static string QxlpyGetCalculate(object[] numlist, object[] strlist)
        {
            if (numlist[0].ToString() == "") {
                throw new ArgumentNullException("Missing Arguments");
            }
            PyExecutor pye = new();
            string calc = pye.GetCalculate(numlist);
            return calc;
        }

        [ExcelFunction(Name = "QxlpyCalculateAddNum")]
        public static double QxlpyCalculateAddNum(string calc_id)
        {
            if (calc_id == "") {
                throw new ArgumentNullException("Missing Arguments");
            }
            PyExecutor pye = new();
            double result = pye.CalculateAddNum(calc_id);
            return result;
        }
    }
}
