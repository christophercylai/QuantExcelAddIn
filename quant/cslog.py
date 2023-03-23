"""
Logger for CSharp
"""
from .qxlpy_log import cs_logger


def LogMessage(logmsg: str, level: str = "INFO") -> str:
    """
    Log message and return a message to Excel
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
