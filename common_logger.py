"""Creates and manages a singleton logger instance.

    Attributes:
        _logger (Logger): Holds the instance of the shared logger

.. _Google Python Style Guide:
   http://google.github.io/styleguide/pyguide.html

"""

import logging
from logging import config as logging_config
from logging import Logger

import config

__logger = None

def get_logger() -> Logger:
    """Returns the existing logger or creates a new one if the first time

    Uses the values specified in the config module to determine the log
    filename, log level (based on verbosity), and whether a console handler
    should be included based on the noisy flag.
    The method will only create a singleton logger that is requested by and
    shared across modules.

    Attributes:
        verbosity (int): Used as subscript to deterimine log level.
            Source is the verbosity attribute from the config object.
        log_path (str): Path and file name for the log file that is populated.
            Source is the log_filename attribute from the config object.
        noisy (int): Determines whether or not the log statements are echoed
            to the console.  Source is the noisy attribute from config object.

    Args:

    Returns:
        Logger: __logger
    """
    global __logger# pylint: disable=global-statement
    verbosity = config.verbosity
    log_path = config.log_filename
    noisy = config.noisy
    if __logger is None:
        name = 'default'
        log_levels = ['ERROR', 'WARNING', 'INFO', 'DEBUG']
        level = log_levels[verbosity]
        cLevel = log_levels[verbosity] if noisy else 'ERROR'
        handlers = ['file', 'console']
        logging_config.dictConfig({
            'version': 1,
            'formatters': {
                'default': {
                    'format': '%(asctime)s - %(levelname)s - %(message)s',
                    'datefmt': '%Y-%m-%d %H:%M:%S'}
            },
            'handlers': {
                'console': {
                    'level': cLevel,
                    'class': 'logging.StreamHandler',
                    'formatter': 'default',
                    'stream': 'ext://sys.stdout'
                },
                'file': {
                    'level': level,
                    'class': 'logging.handlers.RotatingFileHandler',
                    'formatter': 'default',
                    'filename': log_path,
                    'maxBytes': (10*1024*1024),
                    'backupCount': 3,
                    'mode': 'a'
                }
            },
            'loggers': {
                'default': {
                    'level': level,
                    'handlers': handlers
                }
            },
            'disable_existing_loggers': False
        })
        __logger = logging.getLogger(name)
    return __logger

def main():
    """ Only needed by convention """
    pass

if __name__ == '__main__':
    main()
