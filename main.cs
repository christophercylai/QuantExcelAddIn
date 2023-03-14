using ExcelDna.Integration;
using System.Runtime.InteropServices;


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
            dynamic calc = pye.Calculate(numlist);

            // use GCHandleType.Normal because unmanaged obj cannot be Pinned
            // https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.gchandletype
            GCHandle handle = GCHandle.Alloc(calc, GCHandleType.Normal);

            IntPtr pointer = GCHandle.ToIntPtr(handle);
            string pointerDisplay = pointer.ToString();
            handle.Free();
            return pointerDisplay;
        }
    }
}
