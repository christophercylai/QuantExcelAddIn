import logging
from logging import handlers
from os import getenv
from pathlib import Path


logger = logging.getLogger('qxlpy_log')


def set_log_level():
    # set level based on QXLPYLOGLEVEL, default is DEBUG
    log_lvl = getenv('QXLPYLOGLEVEL', 'DEBUG')
    loglvldict = {
        'CRITICAL': logging.CRITICAL,
        'ERROR': logging.ERROR,
        'WARNING': logging.WARNING,
        'INFO': logging.INFO,
        'DEBUG': logging.DEBUG,
    }
    logger.setLevel(loglvldict[log_lvl])

def add_rotate_handler(mb=20, backup=9):
    # create log dir
    logdir = Path(__file__).parents[1]
    logdir = Path(logdir, 'Logs')

    # all logs go to RotatingFileHandler
    formatter = logging.Formatter('%(asctime)s - [%(levelname)s] %(filename)s at line %(lineno)d: %(message)s')
    rotate_handler = handlers.RotatingFileHandler(
        filename=str(Path(logdir, 'qxlpy.log')),
        mode='a',
        maxBytes=1024*1024*mb,
        backupCount=backup,
        encoding='ascii'
    )
    rotate_handler.setFormatter(formatter)
    logger.addHandler(rotate_handler)


set_log_level()
add_rotate_handler()
