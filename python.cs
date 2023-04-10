using Python.Runtime;


namespace qxlpy
{
    public class PyExecutor
    {
        private readonly PyModule SCOPE;

        public PyExecutor()
        {
            string root = Environment.GetEnvironmentVariable("QXLPYDIR");

            string python_dll = $@"{root}\python\python311.dll";
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", python_dll);

            PythonEngine.Initialize();
            SCOPE = Py.CreateScope();

            using (Py.GIL())
            {
                dynamic site = SCOPE.Import("site");
                site.addsitedir(root);
            }
        }

        // THE FOLLOWING FUNCTIONS WILL BE AUTOGEN //
        public string qxlpyLogMessage(string logmsg, string level)
        {
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.cslog");
                string ret = imp.qxlpyLogMessage(logmsg, level);
                return ret;
            }
        }

        public string qxlpyGetCalculate(object[] dub_list)
        {
            // returns the address of the Calculate py obj
            using (Py.GIL())
            {
                var pylist = new PyList();
                double dub;
                bool parse_ok;
                foreach (object n in dub_list) {
                    parse_ok = Double.TryParse(n.ToString(), out dub);
                    if (!parse_ok) { throw new ArrayTypeMismatchException("Wrong type in array"); }
                    pylist.Append(new PyFloat(dub));
                }

                dynamic imp = SCOPE.Import("quant.calculate");
                string ret = imp.qxlpyGetCalculate(pylist);
                return ret;
            }
        }

        public double qxlpyCalculateAddNum(string addr)
        {
            // this func takes the address returned from Calculate
            // and make add computation
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.calculate");
                double ret = imp.qxlpyCalculateAddNum(addr);
                return ret;
            }
        }

        public string qxlpyStoreStrDict(object[,] objdict)
        {
            // returns the address of the Calculate py obj
            // <key: str, value: str>
            using (Py.GIL())
            {
                var pydict = new PyDict();
                string k, v;
                bool parse_ok;
                int dict_len = objdict.GetLength(0);
                string empty = "ExcelDna.Integration.ExcelEmpty";
                for (int i = 0; i < dict_len; i++) {
                    k = objdict[i, 0].ToString();
                    v = objdict[i, 1].ToString();
                    parse_ok = k != empty && v != empty ;
                    if (!parse_ok) { throw new ArrayTypeMismatchException("There is an empty string"); }
                    pydict[k] = new PyString(v);
                }

                dynamic imp = SCOPE.Import("quant.objects");
                string ret = imp.qxlpyStoreStrDict(pydict);
                return ret;
            }
        }

        public object[] qxlpyListGlobalObjects()
        {
            // returns a list of stored objects
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.objects");
                var ret_list = new List<string>();

                PyList pylist = imp.qxlpyListGlobalObjects();
                foreach (PyObject pyobj in pylist) {
                    ret_list.Add(pyobj.ToString());
                }
                object[] ret = ret_list.ToArray();
                return ret;
            }
        }

        public Dictionary<string, List<object>> qxlpyGetStrDict(string obj_name)
        {
            // returns a dictionary object
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.objects");
                var ret = new Dictionary<string, List<object>>();

                var keys = new List<object>();
                var values = new List<object>();
                PyDict pydict = imp.qxlpyGetStrDict(obj_name);
                foreach (PyObject key in pydict) {
                    keys.Add(Convert.ToString(key));
                    values.Add(Convert.ToString(pydict.GetItem(key)));
                }
                ret["keys"] = keys;
                ret["values"] = values;
                return ret;
            }
        }

        public bool qxlpyObjectExists (string obj_name)
        {
            // check the existence of an obj
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.objects");
                bool ret = imp.qxlpyObjectExists(obj_name);
                return ret;
            }
        }
    }
}
