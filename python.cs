using Python.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace qxlpy

{
    public class PyExecutor
    {
        private readonly PyModule scope;
        string root = Environment.GetEnvironmentVariable("qxlpyRoot") ?? @".";

        public PyExecutor()
        {
            string python_dll = $@"{root}\..\python37\python37.dll";
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", python_dll);

            PythonEngine.Initialize();
            scope = Py.CreateScope();
        }

        public string GetPath()
        {
            using (Py.GIL())
            {
                dynamic os = scope.Import("os");
                string path_env = os.getenv("PATH");
                return path_env;
            }
        }

        public string HelloUser(string name = "noname", int age = 0)
        {
            using (Py.GIL())
            {
                dynamic site = scope.Import("site");
                site.addsitedir(root);

                dynamic hw = scope.Import("quant.hello");
                dynamic hw_obj = hw.Hello(name, age);
                string hw_str = hw_obj.say_hello();

                return hw_str;
            }
        }
    }
}
