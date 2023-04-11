import os
import site
import argparse
from pathlib import Path

qxlpydir = Path(os.getenv('QXLPYDIR'))
miscdir = qxlpydir / 'misc'
site.addsitedir(str(miscdir))
from cs_autogen import autogen


def get_args():
    """
    get arguments from command line
    """
    parser = argparse.ArgumentParser(
        prog = 'Autogen C# Excel AddIn',
        description = 'Generate main.cs and python.cs from the qxlpy functions'
    )
    parser.add_argument(
        '-d', '--dry_run', type = bool, default = False,
        help = 'If dry_run is True, then write the resulting files with the ".bak" extension'
    )
    parser.add_argument(
        '-m', '--gen_main', type = bool, default = True,
        help = 'Generate "main.cs" if "gen_main" is True'
    )
    parser.add_argument(
        '-p', '--gen_python', type = bool, default = True,
        help = 'Generate "python.cs" if "gen_python" is True'
    )
    return parser.parse_args()


if __name__ == '__main__':
    ARGS = get_args()
    autogen.autogen(
        gen_main=ARGS.gen_main, gen_python=ARGS.gen_python, dryrun=ARGS.dry_run
    )
