# Excel AddIn for Quantitative Solver
* This AddIn is written with C#, while utilizing Excel-DNA for creating Excel User Defined Functions (UDFs), as well as Python.NET to allow Python functionalities to be called by C#.

## Software Requirements
* Visual Studio 2022
* Excel 2019
* dotNet 6.0
* Python 3.11.3
* Python.Net 3.0.1
* Excel-DNA 1.6.0

## Python and Visual Studio 2022 Setup
### Visual Studio (VS) 2022 Setup
* When installing VS 2022, please install with these 2 modules:
    * .NET desktop development
    * Office/Sharepoint development
* Open `Visual Studio 2022 Developer PowerShell` and start the QuantExcelAddIn project directory.
```
devenv qxlpy.sln
```
* Inside VS 2022, open up the `Package Manager Console` and run:
```
NuGet\Install-Package ExcelDna.AddIn -Version 1.6.0
```
### Python 3.11.3 Setup
* Open `Visual Studio 2022 Developer PowerShell` and go to the QuantExcelAddIn project directory.
```
.\setup.ps1 -bitness 64  # if Excel is 32 bit, bitness is 32
```

## Python Linting and Testing
* Open `Visual Studio 2022 Developer PowerShell` and go to the QuantExcelAddIn project directory.
```
```

## Reference
* [Python.NET Reference](https://pythonnet.github.io/pythonnet/reference.html#)
* [Excel-DNA Quickstart Tutorial](https://colinlegg.wordpress.com/2016/09/07/my-first-c-net-udf-using-excel-dna-and-visual-studio/)
* [Excel Object model)](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model)
