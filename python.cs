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
                var pylist_dub_list = new PyList();
                foreach (object n in dub_list) {
                    double obj;
                    try {
                        obj = Convert.ToDouble(n);
                    } catch (Exception e) {
                        string error_msg = $"Wrong type in array: '{Convert.ToString(n)}' is not of type 'double'";
                        qxlpyLogMessage(error_msg, "ERROR");
                        throw new ArrayTypeMismatchException(error_msg);
                    }
                    pylist_dub_list.Append(new PyFloat(obj));
                }

                dynamic imp = SCOPE.Import("quant.calculate");
                string ret = imp.qxlpyGetCalculate(pylist_dub_list);
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
                for (int i = 0; i < objdict.GetLength(0); i++) {
                    try {
                        k = Convert.ToString(objdict[i, 0]);
                        v = Convert.ToString(objdict[i, 1]);
                    } catch (Exception e) {
                        string error_msg = $"Wrong type in dictionary: ";
                        error_msg += "'{Convert.ToString(k)}' should be 'string' and ";
                        error_msg += "'{Convert.ToString(v)}' should be 'string'";
                        qxlpyLogMessage(error_msg, "ERROR");
                        throw new ArrayTypeMismatchException(error_msg);
                    }
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
                    ret_list.Add(Convert.ToString(pyobj));
                }
                object[] ret = ret_list.ToArray();

                return ret;
            }
        }

        public List<List<object>> qxlpyGetStrDict(string obj_name)
        {
            // returns a dictionary object
            using (Py.GIL())
            {
                dynamic imp = SCOPE.Import("quant.objects");

                var ret = new List<List<object>>();
                var keys = new List<object>();
                var values = new List<object>();
                PyDict pydict = imp.qxlpyGetStrDict(obj_name);
                foreach (PyObject key in pydict) {
                    keys.Add(Convert.ToString(key));
                    values.Add(Convert.ToString(pydict.GetItem(key)));
                }
                ret.Add(keys);
                ret.Add(values);

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
