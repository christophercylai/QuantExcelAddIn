from pathlib import Path
import sys

from .qxlpy_log import py_logger


# add quant module to be searchable python
# i.e., give all modules access to this __init__.py
sys.path.append(Path(__path__[0]))
