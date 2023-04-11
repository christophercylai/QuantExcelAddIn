import os
import site
from pathlib import Path


qxlpydir = Path(os.getenv('QXLPYDIR'))
miscdir = qxlpydir / 'misc'

site.addsitedir(str(miscdir))

from cs_autogen import autogen
autogen.autogen(gen_main=True, gen_python=True, dryrun=False)
