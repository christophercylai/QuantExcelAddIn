using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;


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
            <group id='qxlgroup' label='qxlgroup'>
              <button id='expandfunc' label='Expand Function' onAction='OnButtonPressed'/>
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
            xlApp.ActiveCell.Value = xlApp.ActiveCell.Address(false, false, XlCall.xlcA1R1c1);
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
