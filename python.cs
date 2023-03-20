using Python.Runtime;


namespace qxlpy
{
    public class PyExecutor
    {
        private readonly PyModule SCOPE;

        public PyExecutor()
        {
            DirectoryInfo di = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent;
            // TODO: remove this once we automate Python setup
            if (!di.Exists) {
                throw new DirectoryNotFoundException(
                    "Cannot find the parent directory to QuantExcelAddIn, where Python should reside."
                );
            }
            string root = di.FullName;

            string python_dll = $@"{root}\..\python37\python37.dll";
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", python_dll);

            PythonEngine.Initialize();
            SCOPE = Py.CreateScope();

            using (Py.GIL())
            {
                dynamic site = SCOPE.Import("site");
                site.addsitedir(root);
            }
        }

        public string GetPath()
        {
            using (Py.GIL())
            {
                dynamic os = SCOPE.Import("os");
                string path_env = os.getenv("PATH");
                return path_env;
            }
        }

        // THE FOLLOWING FUNCTIONS WILL BE AUTOGEN //

        public string LogMessage(string logmsg, string level)
        {
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.cslog");
                string ret = imp.LogMessage(logmsg, level);
                return ret;
            }
        }

        public string GetCalculate(object[] numlst)
        {
            // returns the address of the Calculate py obj
            using (Py.GIL())
            {
                PyList pylist = new PyList();
                PyFloat pyf;
                double num;
                bool parse_ok;
                foreach (object n in numlst) {
                    parse_ok = Double.TryParse(n.ToString(), out num);
                    if (!parse_ok) { throw new ArrayTypeMismatchException("Wrong type in array"); }
                    pyf = new PyFloat(num);
                    pylist.Append(pyf);
                }
                dynamic imp = SCOPE.Import("quant.calculate");
                string ret = imp.GetCalculate(numlst);
                return ret;
            }
        }

        public double CalculateAddNum(string addr)
        {
            // this func thats the address returned from Calculate
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.calculate");
                double ret = imp.CalculateAddNum(addr);
                return ret;
            }
        }
    }
}
