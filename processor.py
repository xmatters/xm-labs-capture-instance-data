"""Queries for and processes xmatters instance data

.. _Google Python Style Guide:
   http://google.github.io/styleguide/pyguide.html

"""

import json
import sys
import pprint
from io import TextIOBase
import urllib.parse

import requests
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

import config
import common_logger

_logger = None
_users = None

def _log_xm_error(url, response):
    """Captures and logs errors
        
        Logs the error caused by attempting to call url.
        
        Args:
        url (str): The location being requested that caused the error
        response (object): JSON object that holds the error response
        """
    body = response.json()
    if response.status_code == 404:
        _logger.warn(config.ERR_INITIAL_REQUEST_FAILED_MSG,
                     response.status_code, url)
    else:
        _logger.error(config.ERR_INITIAL_REQUEST_FAILED_MSG,
                      response.status_code, url)
        _logger.error('Response - code: %s, reason: %s, message: %s',
                    str(body['code']) if 'code' in body else "none",
                    str(body['reason']) if 'reason' in body else "none",
                    str(body['message']) if 'message' in body else "none")

def _create_out_file(filename: str) -> TextIOBase:
    """Creates and opens results file

    Args:
        filename (str): Name of file to hold output

    Returns:
        file: outFile
    """
    outFile = open(filename, 'w')
    return outFile

def _process_sites():
    """Capture and save the instances Site objects

    Retrieves the Site object records from xMatters and saves them in
    JSON payload format to the output file.

    Args:
        None

    Return:
        None
    """
    sites_file = _create_out_file(config.sites_filename)
    sites_file.write('[\n')

    # Initialize conditions
    site_objects = []
    total_sites = 0
    cnt = 0
    url = config.xmod_url + '/api/xm/1/sites?offset=0&limit=' + str(config.page_size)
    _logger.debug('Gathering Sites, url=%s', url)

    while True:

        # Get the audit records
        response = requests.get(url, auth=config.basic_auth)
        if response.status_code not in [200, 404]:
            _log_xm_error(url, response)
            break

        # Process the responses
        bodys = response.json()
        total_sites = bodys['total']
        if bodys['count'] > 0:
            _logger.debug("%d Count of %d Total Sites found via url=%s", bodys['count'], bodys['total'], url)
            for body in bodys['data']:
                cnt += 1
                json.dump(body, sites_file)
                sites_file.write(',\n') if cnt < total_sites else sites_file.write('\n')
            site_objects += bodys['data']

        # See if there are any more to get
        if 'links' in bodys and 'next' in bodys['links']:
            url = config.xmod_url + bodys['links']['next']
        else:
            break
            
    _logger.debug("Collected %d of a possible %d Sites.", len(site_objects), total_sites)

    sites_file.write(']')
    sites_file.close()

def _process_users():
    """Capture and save the instances User objects

    Retrieves the User object records from xMatters and saves them in
    JSON payload format to the output file.

    Args:
        None

    Return:
        None
    """
    pass

def _process_devices():
    """Capture and save the instances User's Device objects

    Retrieves the User's Device object records from xMatters and saves them in
    JSON payload format to the output file.

    Args:
        None

    Return:
        None
    """
    pass

def _process_groups():
    """Capture and save the instances Group objects

    Retrieves the User's Group object records from xMatters and saves them in
    JSON payload format to the output file.

    Args:
        None

    Return:
        None
    """
    pass

def process(objects_to_process: list):
    """Capture objects for this instance.

    If requeste contains 'sites', then capture and save Sites.
    If requeste contains 'users', then capture and save Users.
    If requeste contains 'devices', then capture and save Devices.
    If requeste contains 'groups', then capture and save Groups.

    Args:
        objects_to_process (list): The list of object types to capture
    """
    global _logger # pylint: disable=global-statement

    ### Get the current logger
    _logger = common_logger.get_logger()

    # Capture and save the Site objects
    if 'sites' in objects_to_process:
        _process_sites()

    # Capture and save the User objects
    if 'users' in objects_to_process:
        _process_users()

    # Capture and save the Device objects
    if 'devices' in objects_to_process:
        _process_devices()

    # Capture and save the Device objects
    if 'groups' in objects_to_process:
        _process_groups()

def main():
    """In case we need to execute the module directly"""
    pass

if __name__ == '__main__':
    main()
