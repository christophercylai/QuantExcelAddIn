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
    [string]$loglvl='INFO'
)

$root = $pwd.Path
$env:QXLPYDIR = $root
$env:QXLPYLOGLEVEL = $loglvl


if (!(Test-Path $root\Logs)) { mkdir $root\Logs }

# install Python 3.7.9 embeddable package
if (!(Test-Path $root\python)) {
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
    cd $root
}

# run pylint
if ((Test-Path $root\python\python.exe)) {
    if ($pylint) {
        .\python\python.exe -m pylint quant
    }
}

# run pytest
if ((Test-Path $root\python\python.exe)) {
    if ($pytest) {
        .\python\python.exe -m pytest quant
    }
}
