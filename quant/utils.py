"""
Qxlpy Utilities
"""
from typing import List
from datetime import datetime
import pandas as pd
import plotly.express as px

from quant import py_logger
from . import global_obj

# pylint: disable=invalid-name


def qxlpyPlotDataFrame(
        dframe_obj: str,
        plot_type: str = "line",
        title: str = "Data Chart with Plotly",
        dry_run: bool = False
    )-> str:
    """
    Plot cached Pandas DataFrame obj
    plot_type:: line, bar, scatter
    dry_run:: TRUE, FALSE
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
        df_prefix: str, startdate: int = 20221201,
    )-> str:
    """
    Take a list of cached List[float] objs and create a DataFrame
    """
    s_date = datetime.strptime(str(startdate), "%Y%m%d")
    startdate = s_date.strftime("%Y/%m/%d")
    cached_objs = global_obj.list_objs()
    plot_obj = {}
    periods = None
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
        if not periods:
            periods = len(plot_obj[labels[pos]])
        if len(plot_obj[labels[pos]]) != periods:
            errmsg = f'{cached_arrays} cannot be jagged: all elements must have the same length'
            py_logger.error(errmsg)

    df = pd.DataFrame(
        plot_obj,
        index = pd.date_range(start=startdate, freq="M", periods=periods)
    )
    return global_obj.store_obj(df, df_prefix)
