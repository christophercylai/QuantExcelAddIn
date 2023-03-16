using ExcelDna.Integration;


namespace qxlpy
{
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
    }
}
