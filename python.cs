//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using Python.Runtime;


namespace qxlpy

{
    public class PyExecutor
    {
        private readonly PyModule SCOPE;

        public PyExecutor()
        {
            string root = Environment.GetEnvironmentVariable("qxlpyRoot") ?? @".";
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

        public string HelloUser(string name, int age)
        {
            using (Py.GIL())
            {
                dynamic hw = SCOPE.Import("quant.hello");
                dynamic hw_obj = hw.Hello(name, age);
                string hw_str = hw_obj.say_hello();

                return hw_str;
            }
        }

        public dynamic Calculate(double[] numlist)
        {
            using (Py.GIL())
            {
                PyList pylist = new PyList();
                PyFloat pyf;
                foreach (double n in numlist) {
                    pyf = new PyFloat(n);
                    pylist.Append(pyf);
                }
                dynamic calc = SCOPE.Import("quant.calc");
                dynamic py_obj = calc.calculate.Calculate(numlist);
                dynamic quant = SCOPE.Import("quant");

                string pyobj_id = quant.STORE_OBJ(py_obj);

                return pyobj_id;
            }
        }

        public double AddNumbers(string calc_id)
        {
            using (Py.GIL())
            {
                dynamic quant = SCOPE.Import("quant");
                dynamic calc_obj = quant.GET_OBJ(calc_id);
                double result = calc_obj.add();
                return result;
            }
        }
    }
}
