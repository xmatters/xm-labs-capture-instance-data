# encoding: utf-8
"""Extracts and captures all instance-specific data to the file system
    
    This is a command line utility that takes iteratively
	goes through all data in an instance and stores it in the local file
	system as JSON payload files that may be loaded into another instance
	later using the corresponding restore-instance-data.py utility.
    
    Example:
    Arguments are described via the -H command
    Here are some examples::
    
    $ python3 capture-instance-data.py -vv -c -d defaults.json all
    $ python3 capture-instance-data.py -vvv -c -d myco.defaults.json sites

    .. _Google Python Style Guide:
    http://google.github.io/styleguide/pyguide.html
    
    """

import os
import sys

import config
import cli

__all__ = []
__version__ = config.VERSION
__date__ = config.DATE
__updated__ = config.UPDATED

def main(argv=None):
    """ Begins the New Properties process """
    
    args = cli.process_command_line(argv, __doc__)
    args.func(args)
    return 0

if __name__ == "__main__":
    if config.DEBUG:
        sys.argv.append("-h")
        sys.argv.append("-v")
    if config.TESTRUN:
        import doctest
        doctest.testmod()
    if config.PROFILE:
        import cProfile
        import pstats
        profile_filename = 'capture-instance-data_profile.txt'
        cProfile.run('main()', profile_filename)
        statsfile = open("profile_stats.txt", "wb")
        p = pstats.Stats(profile_filename, stream=statsfile)
        stats = p.strip_dirs().sort_stats('cumulative')
        stats.print_stats()
        statsfile.close()
        sys.exit(0)
    try:
        sys.exit(main())
    
    except SystemExit as e:
        if e.code != 0:
            raise
        else:
            os._exit(0)

