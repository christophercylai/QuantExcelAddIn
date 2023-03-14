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
    }
}
