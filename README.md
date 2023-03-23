# Excel AddIn for Quantitative Solver
* This AddIn is written with C#, while utilizing Excel-DNA for creating Excel User Defined Functions (UDFs), as well as Python.NET to allow Python functionalities to be called by C#.

## Software Requirements
* Visual Studio 2022
* dotNet 6.0
* Python 3.7.9
* Python.Net 3.0.1
* Pytest 7.2.2
* Excel-DNA 1.6.0

## Python and Visual Studio 2022 Setup
Note: all of these steps will be automated.
### Python 3.7.9 Setup
* [CLICK HERE to Download Python-3.7.9](https://www.python.org/ftp/python/3.7.9/python-3.7.9-embed-amd64.zip)
* Create a directory: `%USERPROFILE%\github`.
* Git clone this repository inside `github`.
* Extract the zip in `%USERPROFILE%\github\python37`.
* Edit `github\python37\python37._pth` and uncomment `import site`.
* [CLICK HERE to Download pip](https://bootstrap.pypa.io/pip/pip.pyz)
* Move `pip.pyz` to `github\python37`.
* Open up a Powershell and run these commands:
```
cd $env:USERPROFILE\github\python37
.\python.exe pip.pyz install pip
rm pip.pyz
.\python.exe -m pip install virtualenv
.\python.exe -m virtualenv venv ../.venv  # create virtualenv in the github directory
cd ..
./.venv/Scripts/activate
pip install pythonnet==3.0.1  # all pip packages for this project should be installed under this virtualenv
pip install pytest==7.2.2
```
### Visual Studio (VS) 2022 Setup
* When installing VS 2022, please install with these 2 modules:
    * .NET desktop development
    * Office/Sharepoint development
* Open `Developer Command for VS 2022` and start the QuantExcelAddIn project.
```
cd %USERPROFILE%\github\QuantExcelAddIn
devenv qxlpy.sln
```
* Inside VS 2022, open up the `Package Manager Console` and run:
```
NuGet\Install-Package ExcelDna.AddIn -Version 1.6.0
```

## Reference
* [Python.NET Reference](https://pythonnet.github.io/pythonnet/reference.html#)
* [Excel-DNA Quickstart Tutorial](https://colinlegg.wordpress.com/2016/09/07/my-first-c-net-udf-using-excel-dna-and-visual-studio/)
* [Excel Object model)](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model)
