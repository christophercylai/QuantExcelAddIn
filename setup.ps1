# parameters
param(
    [Parameter(
        Mandatory=$false, Position=0, ValueFromPipeline=$false,
        HelpMessage = 'Bitness of Excel 2016'
    )]
    [ValidateSet(32, 64)]
    [int]$bitness=64,

    [Parameter(
        Mandatory=$false, Position=1, ValueFromPipeline=$false,
        HelpMessage = 'Run pylint on the quant Python package'
    )]
    [ValidateSet($true, $false)]
    [bool]$pylint=$false,

    [Parameter(
        Mandatory=$false, Position=2, ValueFromPipeline=$false,
        HelpMessage = 'Run pytest on the quant Python package'
    )]
    [ValidateSet($true, $false)]
    [bool]$pytest=$false,

    [Parameter(
        Mandatory=$false, Position=3, ValueFromPipeline=$false,
        HelpMessage = 'Logging level'
    )]
    [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL')]
    [string]$loglvl='INFO',

    [Parameter(
        Mandatory=$false, Position=4, ValueFromPipeline=$false,
        HelpMessage = 'Build Type'
    )]
    [ValidateSet('Release', 'Debug')]
    [string]$build_type='Release',

    [Parameter(
        Mandatory=$false, Position=5, ValueFromPipeline=$false,
        HelpMessage = 'Build Qxlpy'
    )]
    [ValidateSet($true, $false)]
    [bool]$build_qxlpy=$false,

    [Parameter(
        Mandatory=$false, Position=6, ValueFromPipeline=$false,
        HelpMessage = 'Run Qxlpy Excel AddIn'
    )]
    [ValidateSet($true, $false)]
    [bool]$run_qxlpy=$false
)

# paths setup
$root = $pwd.Path
$env:QXLPYDIR = $root
$env:QXLPYLOGLEVEL = $loglvl
$excel_path="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
if ($bitness -eq 32) {
    $excel_path="C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
}
if (!(Test-Path $excel_path)) {
    echo "WARNING: $excel_path DOES NOT EXISTS"
}

# create the "Logs" directory
if (!(Test-Path $root\Logs)) { mkdir $root\Logs }

# install Python 3.11.3 embeddable package
if (!(Test-Path $root\python)) {
    echo ""
    echo "### Setting Up Python ###"
    if ($bitness -eq 32) {
        wget -Uri https://www.python.org/ftp/python/3.11.3/python-3.11.3-embed-win32.zip -Outfile $root\python310.zip
    } elseif ($bitness -eq 64) {
        wget -Uri https://www.python.org/ftp/python/3.11.3/python-3.11.3-embed-amd64.zip -Outfile $root\python310.zip
    }
    Expand-Archive $root\python310.zip -DestinationPath $root\python
    "import site" | Add-Content -Path $root\python\python311._pth
    cp $root\misc\pip.pyz $root\python
    rm $root\python310.zip

    # install pip
    cd $root\python
   .\python.exe .\pip.pyz install pip==23.0.1
    rm .\pip.pyz

    # pip install the needed Python modules
    .\python.exe -m pip install -r $root\requirements.txt
    $status = $?
    if (! $status) {
        echo "### Python Setup [Failed] ###"
        echo ""
        cd $root
        exit(1)
    }
    cd $root
    echo "### Python Setup [OK] ###"
}

# run pylint
$exitcode = 0
if ($pylint) {
    echo ""
    echo "### Running Pylint ###"
    .\python\python.exe -m pylint quant
    $status = $?
    if (!$status) {
        $exitcode += 1
        echo "### Pylint [Failed] ###"
    } else {
        echo "### Pylint [OK] ###"
    }
}

# run pytest
if ($pytest) {
    echo ""
    echo "### Running Pytest ###"
    .\python\python.exe -m pytest quant
    $status = $?
    if (!$status) {
        $exitcode += 1
        echo "### Pytest [Failed] ###"
    } else {
        echo "### Pytest [OK] ###"
    }
}
if ($exitcode -ne 0) {
    echo "### Pylint and/or Pytest [Failed] ###"
    echo ""
    exit(1)
}

# autogen C# Excel AddIn code
echo ""
echo "### C# Excel AddIn Autogen ###"
cd $root\python
.\python.exe $root\cs_autogen.py
$status = $?
if (! $status) {
    echo "### Autogen [Failed] ###"
    echo ""
    cd $root
    exit(1)
} else {
    cd $root
    echo "### Autogen [OK] ###"
}

# build Qxlpy
if ($build_qxlpy) {
    echo ""
    echo "### Building Qxlpy ###"
    devenv qxlpy.sln /Build $build_type
    $status = $?
    if (! $status) {
        echo "### Qxlpy Building [Failed] ###"
        echo ""
        exit(1)
    }
    echo "### Qxlpy Building [OK] ###"
}

# run Qxlpy Excel AddIn
if ($run_qxlpy) {
    $bit = "64"
    if ($bitness -eq 32) { $bit = "" }
    Start-Process -FilePath $excel_path -ArgumentList "${root}\bin\${build_type}\net6.0-windows\qxlpy-AddIn${bit}.xll"
}
