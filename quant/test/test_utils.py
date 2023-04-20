# pylint: disable=missing-module-docstring
from .. import utils
from .. import objects

def test_plot_data():
    # pylint: disable=missing-function-docstring
    tdbank = objects.qxlpyStoreDoubleList([1, 2, 3])
    bmo = objects.qxlpyStoreDoubleList([3, 2, 1])
    boa = objects.qxlpyStoreDoubleList([1, 3, 2])
    dframe = utils.qxlpyCreatePlotDataFrame(
        [tdbank, bmo, boa],
        ['td', 'bmo', 'boa'], 3
    )
    utils.qxlpyPlotDataFrame(dframe, "line", "dummy", True)
