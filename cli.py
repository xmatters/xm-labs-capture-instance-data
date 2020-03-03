"""Command line and argument processor

.. _Following Google Python Style Guide:
   http://google.github.io/styleguide/pyguide.htm
"""

import sys
import json
import argparse
import getpass
from datetime import datetime

from requests import auth

import config
import common_logger
import processor


def process_sites(args):
    """Called when command line specifies Sites"""
    common_logger.get_logger().debug('Processing Sites only')
    processor.process(['sites'])
    return

def process_users(args):
    """Called when command line specifies Users"""
    common_logger.get_logger().debug('Processing Users only')
    processor.process(['users'])
    return

def process_devices(args):
    """Called when command line specifies Devices"""
    common_logger.get_logger().debug('Processing Devices only')
    processor.process(['devices'])
    return

def process_groups(args):
    """Called when command line specifies Groups"""
    common_logger.get_logger().debug('Processing Groups only')
    processor.process(['groups'])
    return

def process_all(args):
    """Called when command line specifies all operations"""
    common_logger.get_logger().debug('Processing Sites, Users, Devices, and Groups')
    processor.process(['sites','users','devices','groups'])
    return

class _CLIError(Exception):
    """Generic exception to raise and log different fatal errors."""
    def __init__(self, msg, rc=config.ERR_CLI_EXCEPTION):
        super(_CLIError).__init__(type(self))
        self.result_code = rc
        self.msg = "E: %s" % msg

    def __str__(self):
        return self.msg

    def __unicode__(self):
        return self.msg

class __Password(argparse.Action):
    """Container to get and/or hold incoming password"""
    def __call__(self, parser, namespace, values, option_string): # pylint: disable=signature-differs
        if values is None:
            values = getpass.getpass()
        setattr(namespace, self.dest, values)

def process_command_line(argv=None, prog_doc=''): # pylint: disable=too-many-branches,too-many-statements
    """Evaluates and responds to passed in command line arguments"""
    llogger = None

    if argv is None:
        argv = sys.argv
    else:
        sys.argv.extend(argv)

    program_version_message = '%%(prog)s %s (%s)' % (
        "v%s" % config.VERSION, config.UPDATED)
    program_license = """%s

  Created by %s on %s.
  Copyright %s

  Licensed under the %s
  %s

  Distributed on an "AS IS" basis without warranties
  or conditions of any kind, either express or implied.

USAGE
""" % (prog_doc.split("\n")[1], config.AUTHOR, config.DATE,
       config.COPYRIGHT, config.LICENSE, config.LICENSE_REF)

    try:
        # Setup argument parser
        parser = argparse.ArgumentParser(
            description=program_license,
            formatter_class=argparse.RawDescriptionHelpFormatter)
        subparsers = parser.add_subparsers(dest='command_name')
        # Add common arguments
        parser.add_argument("-b", "--basename", dest="base_name",
                            default=None,
                            help=(
                                "If not specified in the defaults file, use "
                                "-b to specify the base name of the output "
                                "file. The names will have a timestamp and "
                                ".json appended to the end."))
        parser.add_argument("-c", "--console", dest="noisy",
                            action='store_true',
                            help=(
                                "If specified, will echo all log output to "
                                "the console at the requested verbosity based "
                                "on the -v option"))
        parser.add_argument("-d", "--defaults", dest="defaults_filename",
                            default="defaults.json",
                            help=(
                                "Specifes the name of the file containing "
                                "default settings [default: %(default)s]"))
        parser.add_argument("-i", "--itype", dest="instance_type",
                            default=None,
                            choices=['np', 'prod'],
                            help=(
                                  "Specifies whether we are updating the "
                                  "Production (prod) or Non-Production "
                                  "(np) instance. "
                                  "[default: %(default)s]"))
        parser.add_argument("-l", "--lfile", dest="log_filename",
                            default=None,
                            help=(
                                "If not specified in the defaults file, use "
                                "-l to specify the base name of the log file. "
                                "The name will have a timestamp and .log "
                                "appended to the end."))
        parser.add_argument("-o", "--odir", dest="out_directory",
                            default=None,
                            help=(
                                  "If not specified in the defaults file, use -o"
                                  " to specify the file system location where "
                                  "the output files will be written."))
        parser.add_argument('-p', action=__Password, nargs='?',
                            dest='password', default=None,
                            help=(
                                  "If not specified in the defaults file, use -p"
                                  " to specify a password either on the command"
                                  " line, or be prompted"))
        parser.add_argument("-u", "--user", dest="user",
                            default=None,
                            help=("If not specified in the defaults file, use "
                                  "-u to specify the xmatters user id that has"
                                  " permissions to get Event and Notification "
                                  "data."))
        parser.add_argument("-V", "--version",
                          action='version', version=program_version_message)
        parser.add_argument("-v", dest="verbose",
                          action="count", default=0,
                          help=(
                                "set verbosity level.  Each occurrence of v "
                                "increases the logging level.  By default it "
                                "is ERRORs only, a single v (-v) means add "
                                "WARNING logging, a double v (-vv) means add "
                                "INFO logging, and a tripple v (-vvv) means "
                                "add DEBUG logging [default: %(default)s]"))
        parser.add_argument("-x", "--xmodurl", dest="xmod_url",
                            default=None,
                            help=("If not specified in the defaults file, use "
                                  "-i to specify the base URL of your xmatters"
                                  " instance.  For example, 'https://myco.host"
                                  "ed.xmatters.com' without quotes."))
        #Add in event command parsers
        sites_parser = subparsers.add_parser(
            'sites', description=("Only captures 'Sites'"),
            help=("Use this command in order to only read and save Sites."))
        sites_parser.set_defaults(func=process_sites)
        users_parser = subparsers.add_parser(
            'users', description=("Only captures 'Users'"),
            help=("Use this command in order to only read and save Uses."))
        users_parser.set_defaults(func=process_users)
        devices_parser = subparsers.add_parser(
             'devices', description=("Only captures 'Devices'"),
             help=("Use this command in order to only read and save Devices."))
        devices_parser.set_defaults(func=process_devices)
        groups_parser = subparsers.add_parser(
             'groups', description=("Only captures 'Groups'"),
             help=("Use this command in order to only read and save Groups."))
        groups_parser.set_defaults(func=process_groups)
        all_parser = subparsers.add_parser(
            'all', description=("Capture Users, Devices, and Groups"),
            help=("Use this command in order to process all objects "
                  "from the instance: Sites, Users, Devices, Groups."))
        all_parser.set_defaults(func=process_all)

        # Process arguments
        args = parser.parse_args()

        # Dereference the arguments into the configuration object
        user = None
        password = None
        if args.base_name:
            config.base_name = args.base_name
        if args.instance_type:
            config.instance_type = args.instance_type
        if args.log_filename:
            config.log_filename = args.log_filename
        if args.out_directory:
            config.out_directory = args.out_directory
        if args.noisy > 0:
            config.noisy = args.noisy
        if args.password:
            password = args.password
        if args.user:
            user = args.user
        if args.verbose > 0:
            config.verbosity = args.verbose
        if args.xmod_url:
            config.xmod_url = args.xmod_url

        # Try to read in the defaults from defaults.json
        try:
            with open(args.defaults_filename) as defaults:
                cfg = json.load(defaults)
        except FileNotFoundError:
            raise(_CLIError(
                config.ERR_CLI_MISSING_DEFAULTS_MSG % args.defaults_filename,
                config.ERR_CLI_MISSING_DEFAULTS_CODE))

        # Process the defaults
        if user is None and 'user' in cfg:
            user = cfg['user']
        if password is None and 'password' in cfg:
            password = cfg['password']
        if config.base_name is None and 'baseName' in cfg:
            config.base_name = cfg['baseName']
        if config.dir_sep is None and 'dirSep' in cfg:
            config.dir_sep = cfg['dirSep']
        if config.log_filename is None and 'logFilename' in cfg:
            config.log_filename = cfg['logFilename']
        if config.out_directory is None and 'outDirectory' in cfg:
            config.out_directory = cfg['outDirectory']
        if config.xmod_url is None and 'xmodURL' in cfg:
            config.xmod_url = cfg['xmodURL']
        if config.verbosity == 0 and 'verbosity' in cfg:
            if cfg['verbosity'] in [1, 2, 3]:
                config.verbosity = cfg['verbosity']
        if config.instance_type is None and 'instance' in cfg:
            config.instance_type = cfg['instance']

        # Validate and default instance type to non production
        if config.instance_type is None:
            config.instance_type = 'np'
        if config.instance_type not in ['prod', 'np']:
            raise ValueError('Input must be either "prod" or "np"')

        config.non_prod = True if config.instance_type == 'np' else False
        config.command_name = args.command_name

        # Fix file names
        if config.log_filename:
            config.log_filename = (
                config.out_directory + config.dir_sep + config.base_name +
                '.' + config.instance_type + '.' + config.log_filename +
                '.' + config.time_str + '.log')
        config.sites_filename = (
            config.out_directory + config.dir_sep + config.base_name + '.' +
            config.instance_type + '.sites.' + config.time_str + '.json')
        config.users_filename = (
            config.out_directory + config.dir_sep + config.base_name + '.' +
            config.instance_type + '.users.' + config.time_str + '.json')
        config.devices_filename = (
            config.out_directory + config.dir_sep + config.base_name + '.' +
            config.instance_type + '.devices.' + config.time_str + '.json')
        config.groups_filename = (
            config.out_directory + config.dir_sep + config.base_name + '.' +
            config.instance_type + '.groups.' + config.time_str + '.json')
        config.admin_filename = (
            config.out_directory + config.dir_sep + config.base_name + '.' +
            config.instance_type + '.admin.' + config.time_str + '.json')

        # Initialize logging
        llogger = common_logger.get_logger()
        llogger.info("Capture Instance Data Processor Started.")
        llogger.debug("After parser.parse_args(), command_name=%s",
                    args.command_name)

        # Final verification of arguments
        if config.xmod_url:
            llogger.info("xmatters Instance URL is: %s", config.xmod_url)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_XMOD_URL_MSG,
                            config.ERR_CLI_MISSING_XMOD_URL_CODE))
        if user:
            llogger.info("User is: %s", user)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_USER_MSG,
                            config.ERR_CLI_MISSING_USER_CODE))
        if password:
            llogger.info("Password was provided.")
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_PASSWORD_MSG,
                            config.ERR_CLI_MISSING_PASSWORD_CODE))
        if config.base_name:
            llogger.info("Base name is %s", config.base_name)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_BASENAME_MSG,
                            config.ERR_CLI_MISSING_BASENAME_CODE))
        if args.command_name:
            llogger.info("About to begin processing command(s): %s",
                        config.command_name)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_COMMAND_MSG,
                            config.ERR_CLI_MISSING_COMMAND_CODE))
        if config.out_directory:
            llogger.info("Output directory is: %s", config.out_directory)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_OUTPUT_DIR_MSG,
                            config.ERR_CLI_MISSING_OUTPUT_DIR_CODE))

        # Setup the basic auth object for subsequent REST calls
        config.basic_auth = auth.HTTPBasicAuth(user, password)

        # Make sure we have a func None == all
        if args.func is None:
            args.func = process_all

        return args

    except KeyboardInterrupt:
        ### handle keyboard interrupt ###
        sys.exit(0)

    except _CLIError as cli_except:
        if config.DEBUG or config.TESTRUN:
            raise cli_except # pylint: disable=raising-bad-type
        msg = config.program_name + ": Command Line Error - " + cli_except.msg + " (" + str(cli_except.result_code) + ")"
        if llogger:
            llogger.error(msg)
        else:
            sys.stderr.write(msg+"\n")
        indent = (len(config.program_name) + 30) * " "
        sys.stderr.write(indent + "  for help use --help\n")
        sys.exit(cli_except.result_code)

    except Exception as exc: # pylint: disable=broad-except
        if config.DEBUG or config.TESTRUN:
            raise exc # pylint: disable=raising-bad-type
        sys.stderr.write(
            config.program_name + ": Unexpected exception " + repr(exc) + "\n")
        indent = (len(config.program_name) + 30) * " "
        sys.stderr.write(indent + "  For assistance use --help\n")
        sys.exit(config.ERR_CLI_EXCEPTION)

def main():
    """ By convention and for completeness """
    pass

if __name__ == '__main__':
    main()
