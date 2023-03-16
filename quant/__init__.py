from pathlib import Path
import sys

# This line puts 'quant' in Python's import search path
sys.path.append(sys.modules['quant'].__file__)

### ========== Global Variables ========== ###

# objects can be stored by Python and then retrived by C# later with GLOBAL_OBJS
GLOBAL_OBJS = {}
