# -*- coding: utf-8 -*-
"""Holds variables shared between modules

.. _Google Python Style Guide:
   http://google.github.io/styleguide/pyguide.html

"""

import os
import sys
import time

# Used by command line processor
VERSION = 1.0
DATE = '2018-12-13'
UPDATED = '2018-12-13'
AUTHOR = 'jolin@xmatters.com'
COPYRIGHT = '2018 xmatters, Inc. All rights reserved.'
LICENSE = 'Apache License 2.0'
LICENSE_REF = 'http://www.apache.org/licenses/LICENSE-2.0'

# Global Constants
DEBUG = 0
TESTRUN = 0
PROFILE = 0

""" Global Variables
    Defaults are set from configuration file via processArgs()
"""
program_name = os.path.basename(sys.argv[0])
time_str = time.strftime("-%Y%m%d-%H%M")
page_size = 1000
xmod_url = None
out_directory = None
properties_filename = None
log_filename = None
dir_sep = "/"
basic_auth = None
verbosity = 0
noisy = False
non_prod = None
instance_type = None
base_name = None
sites_filename = None
users_filename = None
devices_filename = None
groups_filename = None

# Error codes
ERR_CLI_EXCEPTION = -1
ERR_CLI_MISSING_DEFAULTS_CODE = -2
ERR_CLI_MISSING_DEFAULTS_MSG = "Missing defaults file: "
ERR_CLI_MISSING_XMOD_URL_CODE = -3
ERR_CLI_MISSING_XMOD_URL_MSG = ("xmatters URL was not specified on the command"
                                " line or via defaults")
ERR_CLI_MISSING_USER_CODE = -4
ERR_CLI_MISSING_USER_MSG = ("xmatters User was not specified on the command "
                            "line or via defaults")
ERR_CLI_MISSING_PASSWORD_CODE = -5
ERR_CLI_MISSING_PASSWORD_MSG = ("xmatters Password was not specified on the "
                                "command line or via defaults")
ERR_CLI_MISSING_OUTPUT_DIR_CODE = -6
ERR_CLI_MISSING_OUTPUT_DIR_MSG = ("Output directory was not specified on the "
                                  "command line or via defaults")
ERR_CLI_MISSING_PROPERTIES_FILENAME_CODE = -7
ERR_CLI_MISSING_PROPERTIES_FILENAME_MSG = ("Properties information filename was not "
                                           "specified on the command line or via defaults")
ERR_CLI_MISSING_COMMAND_CODE = -8
ERR_CLI_MISSING_COMMAND_MSG = ("A command was not specified.  Must specify 'sites', 'users', "
                               "'devices', 'groups', or 'all'")
ERR_CLI_MISSING_BASENAME_CODE = -9
ERR_CLI_MISSING_BASENAME_MSG = ("Base output file name was not specified "
                                "on the command line or via defaults")
ERR_REQUEST_EXCEPTION_CODE = -10
ERR_REQUEST_EXCEPTION_MSG = ("Request Exception while trying to GET %s\n"
                             "Exception: %s")
ERR_REQUEST_NEXT_EXCEPTION_CODE = -11
ERR_INITIAL_REQUEST_FAILED_CODE = -12
ERR_INITIAL_REQUEST_FAILED_MSG = ("Error %d on initial request to %s.\nPlease "
                                  "verify instance address, user, and password")

def main():
    """ To pass conventions, in case we need to execute main """
    pass

if __name__ == '__main__':
    main()
