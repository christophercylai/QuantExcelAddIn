"""
Logger for CSharp
"""
from .global_log import cs_logger

# pylint: disable=invalid-name


def qxlpyLogMessage(logmsg: str, level: str = "INFO") -> str:
    """
    Log message and return a message to Excel
    level:: INFO, DEBUG, WARNING, ERROR, CRITICAL
    """
    loglevels = {
        "DEBUG": cs_logger.debug,
        "INFO": cs_logger.info,
        "WARNING": cs_logger.warning,
        "ERROR": cs_logger.error,
        "CRITICAL": cs_logger.critical
    }
    level = level if level in loglevels else "INFO"
    loglevels[level](logmsg)
    ret = f"'{logmsg}' is written on Logs/qxlcs.log as {level}"
    return ret
