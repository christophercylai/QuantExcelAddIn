import os
import site
from pathlib import Path


qxlpydir = Path(os.getenv('QXLPYDIR'))
miscdir = qxlpydir / 'misc'

site.addsitedir(str(miscdir))

from cs_autogen import autogen
