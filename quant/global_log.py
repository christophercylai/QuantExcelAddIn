"""
qxlpy logging configuration
"""
import logging
from logging import handlers
from os import getenv
from pathlib import Path


py_logger = logging.getLogger('qxlpy_log')
cs_logger = logging.getLogger('csharp_log')


def set_log_level(logger):
    """
    set logging level for a logger
    default is DEBUG
    """
    log_lvl = getenv('QXLPYLOGLEVEL', 'INFO')
    loglvldict = {
        'CRITICAL': logging.CRITICAL,
        'ERROR': logging.ERROR,
        'WARNING': logging.WARNING,
        'INFO': logging.INFO,
        'DEBUG': logging.DEBUG,
    }
    logger.setLevel(loglvldict[log_lvl])

def add_rotate_handler(logger, log_format: str, log_name: str, mbyte: int = 20, backup: int = 9):
    """
    log handler that rotates a limited number of log files,
    so that they won't get too big
    """
    # create log dir
    logdir = Path(__file__).parents[1]
    logdir = Path(logdir, 'Logs')

    # all logs go to RotatingFileHandler
    formatter = logging.Formatter(log_format)
    rotate_handler = handlers.RotatingFileHandler(
        filename=str(Path(logdir, log_name)),
        mode='a',
        maxBytes=1024*1024*mbyte,
        backupCount=backup,
        encoding='ascii'
    )  # pylint: disable=line-too-long
    rotate_handler.setFormatter(formatter)
    logger.addHandler(rotate_handler)


set_log_level(py_logger)
add_rotate_handler(
    py_logger,
    '%(asctime)s - [%(levelname)s] %(filename)s at line %(lineno)d: %(message)s', 'qxlpy.log'
)
set_log_level(cs_logger)
add_rotate_handler(cs_logger, '%(asctime)s - [%(levelname)s]: %(message)s', 'qxlcs.log')
