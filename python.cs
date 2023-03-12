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
        public PyExecutor()
        {
            string python_dll = @"C:\Users\christophercylai\Python37\python37.dll";
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", python_dll);
            PythonEngine.Initialize();
            scope = Py.CreateScope();
        }

        public string GetPath()
        {
            using (Py.GIL())
            {
                scope.Exec("print('HelloWorld!')");
                dynamic os = scope.Import("os");
                string path_env = os.getenv("PATH");
                return path_env;
            }
        }
    }
}
