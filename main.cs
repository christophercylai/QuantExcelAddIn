using ExcelDna.Integration;


namespace qxlpy
{
    public static class ExcelFunc
    {
        [ExcelFunction(Name = "ReturnPath")]
        public static string ReturnPath()
        {
            PyExecutor pye = new();
            string ret = pye.GetPath();
            return ret;
        }
    }
}
