using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;


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
            // RomAbsolute=false, ColumnAbsolute=false, AddressReference
            string cell_addr = xlApp.ActiveCell.Address(false, false, XlCall.xlcA1R1c1);
            var rgx_x = new Regex(@"(?<=^R\[)[0-9]+");
            var rgx_y = new Regex(@"(?<=.+C\[)[0-9]+");
            Match match_x = rgx_x.Match(cell_addr);
            Match match_y = rgx_y.Match(cell_addr);
            // Range A1 = RC, A2 = RC[1], B1 = R[1]C, B2 = R[1]C[1] ...
            // Cell A1 = 1, 1
            int x = match_x.Success ? int.Parse(match_x.Value)+1 : 1;
            int y = match_y.Success ? int.Parse(match_y.Value)+1 : 1;
            var cell_value = xlApp.Cells(x, y).Value;
            xlApp.Cells(1, 1).Value = cell_value;
        }
    }

    public static class ExcelFunc
    {
        [ExcelFunction(Name = "ReturnPath")]
        public static string ReturnPath()
        {
            PyExecutor pye = new();
            string path = pye.GetPath();
            return path;
        }

        [ExcelFunction(Name = "HelloUser")]
        public static string HelloUser(string name, int age)
        {
            PyExecutor pye = new();
            string hw = pye.HelloUser(name, age);
            return hw;
        }

        [ExcelFunction(Name = "GetCalculate")]
        public static string GetCalculate()
        {
            PyExecutor pye = new();
            double[] numlist = {3, 4, 5};
            string calc = pye.Calculate(numlist);
            return calc;
        }

        [ExcelFunction(Name = "CalculateAdd")]
        public static double CalculateAdd(string calc_id)
        {
            PyExecutor pye = new();
            double result = pye.AddNumbers(calc_id);
            return result;
        }

        [ExcelFunction(Name = "LogMessage")]
        public static string LogMessage(string log_msg, string level)
        {
            PyExecutor pye = new();
            pye.PrintLog(log_msg, level);
            string ret = "'" + log_msg + "'" + " is written on Logs/qxlpy.log";
            return ret;
        }
    }
}
