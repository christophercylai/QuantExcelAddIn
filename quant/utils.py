"""
Qxlpy Utilities
"""
from typing import List, Dict

# pylint: disable=invalid-name


def qxlpyCSharpAutogenTest(
    multiplier: float, num_list: List[int], mix_dict: Dict[str, float])-> Dict[str, float]:
    """
    Test CSharp autoget script
    """
    for ea_int in num_list:
        mix_dict[str(ea_int)] = ea_int * multiplier

    return mix_dict
