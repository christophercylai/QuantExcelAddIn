# pylint: disable=missing-module-docstring
from .. import cslog


def test_cslog():
    # pylint: disable=missing-function-docstring
    logmsg = "test logging"
    loglvl = "DUMMY"
    expected = f"'{logmsg}' is written on Logs/qxlcs.log as INFO"
    actual = cslog.LogMessage(logmsg, loglvl)
    assert expected == actual
