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

        public void LogMessage(string logmsg, string level)
        {
            using (Py.GIL())
            {
                dynamic quant = SCOPE.Import("quant");
                level = string.IsNullOrEmpty(level) ? "INFO" : level;
                var loglevels = new Dictionary<string, dynamic> {
                    {"DEBUG", quant.cs_logger.debug},
                    {"INFO", quant.cs_logger.info},
                    {"WARNING", quant.cs_logger.warning},
                    {"ERROR", quant.cs_logger.error},
                    {"CRITICAL", quant.cs_logger.critical}
                };

                if (!loglevels.ContainsKey(level)) {
                    level = "INFO";
                }
                dynamic logger = loglevels[level];
                logger(logmsg);
            }
        }

        // THE FOLLOWING FUNCTIONS WILL BE AUTOGEN //

        public string GetCalculate(object[] numlist)
        {
            // returns the address of the Calculate py obj
            using (Py.GIL())
            {
                PyList pylist = new PyList();
                PyFloat pyf;
                double num;
                bool parse_ok;
                foreach (object n in numlist) {
                    parse_ok = Double.TryParse(n.ToString(), out num);
                    if (!parse_ok) { throw new ArrayTypeMismatchException("Wrong type in array"); }
                    pyf = new PyFloat(num);
                    pylist.Append(pyf);
                }
                dynamic calc = SCOPE.Import("quant.calculate");
                string addr = calc.GetCalculate(numlist);

                return addr;
            }
        }

        public double CalculateAddNum(string calc_id)
        {
            // this func thats the address returned from Calculate
            using (Py.GIL())
            {
                dynamic calc = SCOPE.Import("quant.calculate");
                double result = calc.CalculateAddNum(calc_id);
                return result;
            }
        }
    }
}
