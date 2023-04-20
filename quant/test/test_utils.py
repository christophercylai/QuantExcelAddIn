# pylint: disable=missing-module-docstring
from .. import utils
from .. import objects

def test_plot_data():
    # pylint: disable=missing-function-docstring
    tdbank = objects.qxlpyStoreDoubleList([1, 2, 3])
    bmo = objects.qxlpyStoreDoubleList([3, 2, 1])
    boa = objects.qxlpyStoreDoubleList([1, 3, 2])
    utils.qxlpyPlotData(
        [tdbank, bmo, boa], 3,
        ['td', 'bmo', 'boa'],
        show_plot = False
    )
