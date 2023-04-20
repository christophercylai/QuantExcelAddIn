"""
Qxlpy Utilities
"""
from typing import List, Dict
import pandas as pd

from quant import py_logger
from . import global_obj

pd.options.plotting.backend = "plotly"

# pylint: disable=invalid-name


def qxlpyCSharpAutogenTest(
        num_list: List[int], mix_dict: Dict[str, float],
        multiplier: float = 3.14, overwrite: bool = True
    )-> Dict[str, float]:
    """
    Test CSharp autoget script
    """
    if overwrite:
        for ea_int in num_list:
            mix_dict[str(ea_int)] = ea_int * multiplier
    else:
        for ea_int in num_list:
            if not str(ea_int) in mix_dict:
                mix_dict[str(ea_int)] = ea_int * multiplier

    return mix_dict

def qxlpyPlotData(
        cached_arrays: List[str], periods: int,
        labels: List[str], startdate: str = '2022/12/01',
        title: str = "Data Points with Plotly",
        show_plot: bool = False
    )-> str:
    """
    Plot cached python list
    """
    cached_objs = global_obj.list_objs()
    plot_obj = {}
    for pos in range(len(cached_arrays)):  # pylint: disable=consider-using-enumerate
        handle = cached_arrays[pos]
        # check obj exists
        if not handle in cached_objs:
            errmsg = f'"{handle}" does not exist in cache'
            py_logger.error(errmsg)
            raise KeyError(errmsg)

        plot_obj[labels[pos]] = global_obj.get_obj(handle)

        # check obj type
        if not isinstance(plot_obj[labels[pos]], list):
            errmsg = f'"{handle}" is not of type "list"'
            py_logger.error(errmsg)
            raise TypeError(errmsg)
        # check obj length
        if len(plot_obj[labels[pos]]) != periods:
            errmsg = f'Length of {labels[pos]} is not the same as periods: {periods}'
            py_logger.error(errmsg)

    df = pd.DataFrame(
        plot_obj,
        index = pd.date_range(start=startdate, freq="M", periods=periods)
    )

    data_frame_handle = global_obj.store_obj(df)

    # plotting
    plotly = df.plot(
        title = title
    )
    if show_plot:
        plotly.show()

    return data_frame_handle
