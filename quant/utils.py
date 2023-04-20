"""
Qxlpy Utilities
"""
from typing import List, Dict
import pandas as pd
import plotly.express as px

from quant import py_logger
from . import global_obj

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

def qxlpyPlotDataFrame(
        dframe_obj: str,
        plot_type: str = "line",
        title: str = "Data Points with Plotly",
        dry_run: bool = False
    )-> str:
    """
    Plot cached Pandas DataFrame obj
    startdate has to be in string, such as: '2022/12/01'
    """
    plot_obj = {
        "line": px.line,
        "bar": px.bar,
        "scatter": px.scatter
    }
    if plot_type not in plot_obj:
        errmsg = f'Plot type "{plot_type}" is not supported. Plot types: {list(plot_obj.keys())}'
        py_logger.error(errmsg)
        raise KeyError(errmsg)
    fig = plot_obj[plot_type](global_obj.get_obj(dframe_obj), title=title)
    if not dry_run:
        fig.show()
    return 'SUCCESS'

def qxlpyCreatePlotDataFrame(
        cached_arrays: List[str], labels: List[str],
        periods: int, startdate: str = '2022/12/01'
    )-> str:
    """
    Take a list of cached List[float] objs and create a DataFrame
    startdate must be in string format that looks like 2022/12/01
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
    return global_obj.store_obj(df)
