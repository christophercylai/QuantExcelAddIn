from .. import cslog


def test_cslog():
    logmsg = "test logging"
    loglvl = "DUMMY"
    expected = f"'{logmsg}' is written on Logs/qxlcs.log as INFO"
    actual = cslog.LogMessage(logmsg, loglvl)
    assert expected == actual
