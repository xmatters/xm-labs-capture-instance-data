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

import config
import common_logger

_logger = None
_users = None
_sites_cache = {}
# Holds admin info: company_admins, roles, timezones, country, language
_admin_objects = None

def _update_admin(a_type: str, a_value: str):
    """Updates the admin objects set
        
    Puts the value into the appropriate set.
        
    Args:
        a_type (str): The name of a set in the admin dict
        a_value (str): The value to add to the dict
    """
    _admin_objects[a_type].add(a_value)

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

def _lookup_site_name(site_id: str):
    """Retrieves a Site name by ID

    Attempts to find the Site by it's ID in the Site cache,
    if not found, then retrieve the Site object from xMatters.

    Args:
        site_id (str): The ID of the site to find

    Return:
        None
    """
    if site_id in _sites_cache:
        return _sites_cache[site_id]
    
    # Site was not in the Cache, so get it from xMatters

    # Initialize conditions
    url = config.xmod_url + '/api/xm/1/sites/' + site_id
    _logger.debug('Retrieving Site, url=%s', url)

    # Get the site records
    response = requests.get(url, auth=config.basic_auth)
    if response.status_code not in [200, 404]:
        _log_xm_error(url, response)
        _sites_cache[site_id] = None
        return None

    # Process the responses
    site = response.json()
    _sites_cache[site['id']] = site['name']
    return site['name']

def _process_sites():
    """Capture and save the instances Site objects

    Retrieves the Site object records from xMatters and saves them in
    JSON payload format to the output file.

    Args:
        None

    Return:
        None
    """
    _logger.info('Begin Gathering Sites.')
    sites_file = _create_out_file(config.sites_filename)
    sites_file.write('[\n')

    # Initialize conditions
    site_objects = []
    total_sites = 0
    cnt = 0
    url = config.xmod_url + '/api/xm/1/sites?offset=0&limit=' + str(config.page_size)
    _logger.debug('Gathering Sites via url=%s', url)

    while True:

        # Get the site records
        response = requests.get(url, auth=config.basic_auth)
        if response.status_code not in [200]:
            _log_xm_error(url, response)
            break

        # Process the responses
        bodys = response.json()
        total_sites = bodys['total']
        if bodys['count'] > 0:
            _logger.debug("%d Count of %d Total Sites found via url=%s", bodys['count'], bodys['total'], url)
            for body in bodys['data']:
                cnt += 1
                _logger.info(f'Capturing Site "{body["name"]}"')
                json.dump(body, sites_file)
                sites_file.write(',\n') if cnt < total_sites else sites_file.write('\n')
                _sites_cache[body['id']] = body['name']
                # Update admin sets
                if 'language' in body: _update_admin('languages', body['language'])
                if 'timezone' in body: _update_admin('timezones', body['timezone'])
                if 'country' in body: _update_admin('countries', body['country'])
            site_objects += bodys['data']

        # See if there are any more to get
        if 'links' in bodys and 'next' in bodys['links']:
            url = config.xmod_url + bodys['links']['next']
        else:
            break
            
    _logger.info("Collected %d of a possible %d Sites.", len(site_objects), total_sites)

    sites_file.write(']')
    sites_file.close()

def _get_user_devices(user_id: str, target_name: str):
    """Return a User's Devices

    Retrieves the Device records from xMatters for the specified User and
    returns them as a list.

    Args:
        user_id (str): The User's UUID
        target_name (str): The User's targetName field

    Return:
        device_list (list): List of dictionaries of the User's devices.
    """
    # Initialize conditions
    device_list = []
    total_devices = 0
    url = config.xmod_url + '/api/xm/1/people/' + user_id + '/devices/?embed=timeframes&offset=0&limit=' + str(config.page_size)
    _logger.debug('Gathering Devices for user "%s", url=%s', target_name, url)

    while True:

        # Get the site records
        response = requests.get(url, auth=config.basic_auth)
        if response.status_code not in [200, 404]:
            _log_xm_error(url, response)
            break

        # Process the responses
        bodys = response.json()
        total_devices = bodys['total']
        if bodys['count'] > 0:
            _logger.debug('%d Count of %d Total Devices found for User "%s" via url=%s', bodys['count'], bodys['total'], target_name, url)
            device_list += bodys['data']
            # Use Timeframe to update the timezones admin set
            for body in bodys['data']:
                if 'timeframes' in body:
                    for timeframe in body['timeframes']:
                        if 'timezone' in timeframe: _update_admin('timezones', timeframe['timezone'])
                _update_admin('devices', body['deviceType'] + '|' + body['name'])
                if 'provider' in body: _update_admin('usps', body['provider']['id'])

        # See if there are any more to get
        if 'links' in bodys and 'next' in bodys['links']:
            url = config.xmod_url + bodys['links']['next']
        else:
            break

    _logger.debug('Collected %d of a possible %d Devices for User "%s".', len(device_list), total_devices, target_name)

    return device_list

def _get_user(user_id: str, target_name: str):
    """Attempst to retrieve User by id.
        
        If the User exists, retrieve and return the object.
        If not, return None
        
        Args:
        user_id (str): UUID of User to retrieve
        target_name (str): targetName field value for the specified user.
        """
    _logger.info(f'Capturing User: "{target_name}".')
    
    # Set our resource URLs
    url = config.xmod_url + '/api/xm/1/people/' + urllib.parse.quote(user_id) + '?embed=roles,supervisors'
    _logger.debug('Attempting to retrieve User "%s" via url: %s', target_name, url)
    
    # Make the request
    try:
        response = requests.get(url, auth=config.basic_auth)
    except requests.exceptions.RequestException as e:
        _logger.error(config.ERR_REQUEST_EXCEPTION_CODE, url, repr(e))
        return None
    
    # If the initial response fails, log and return null
    if response.status_code != 200:
        _log_xm_error(url, response)
        return None
    
    # Process the response
    user_obj = response.json()
    # _logger.debug('Found User "%s" - json body: %s', user_obj['firstName'] + ' ' + user_obj['lastName'], pprint.pformat(user_obj))
    _logger.debug('Found User "%s" - json body.id: %s', user_obj['firstName'] + ' ' + user_obj['lastName'], user_obj['id'])

    # Update the Admin object
    if 'language' in user_obj: _update_admin('languages', user_obj['language'])
    if 'timezone' in user_obj: _update_admin('timezones', user_obj['timezone'])
    if 'roles' in user_obj and user_obj['roles']['total'] > 0:
        for role in user_obj['roles']['data']:
            _update_admin('roles', role['name'])
            if role['name'] == config.company_admin_role:
                _update_admin('admins', user_obj['targetName'])

    return user_obj

def _process_users(include_devices: bool):
    """Capture and save the instances User objects

    Retrieves the User object records from xMatters and saves them in
    JSON payload format to the output file.

    Args:
        include_devices (bool): If True, get the User's devices too

    Return:
        None
    """
    _logger.info('Begin gathering Users.')
    users_file = _create_out_file(config.users_filename)
    users_file.write('[\n')

    # Initialize conditions
    user_objects = []
    total_users = 0
    cnt = 0
    url = config.xmod_url + '/api/xm/1/people?offset=0&limit=' + str(config.page_size)
    _logger.debug('Gathering Users via url=%s', url)

    while True:

        # Get the user records
        response = requests.get(url, auth=config.basic_auth)
        if response.status_code not in [200, 404]:
            _log_xm_error(url, response)
            break

        # Process the responses
        bodys = response.json()
        total_users = bodys['total']
        if bodys['count'] > 0:
            _logger.debug("%d Count of %d Total Users found via url=%s", bodys['count'], bodys['total'], url)
            for body in bodys['data']:

                # Get the full user object, including Roles and Supervisors
                user_obj = {}
                a_user = _get_user(body['id'], body['targetName'])
                if a_user is not None:
                    user_obj['user'] = a_user

                    # Get the devices, if requested
                    if include_devices:
                        user_obj['devices'] = _get_user_devices(a_user['id'], a_user['targetName'])

                    # Save the User
                    cnt += 1
                    json.dump(user_obj, users_file)
                    users_file.write(',\n') if cnt < total_users else users_file.write('\n')

            user_objects += bodys['data']

        # See if there are any more to get
        if 'links' in bodys and 'next' in bodys['links']:
            url = config.xmod_url + bodys['links']['next']
        else:
            break
            
    _logger.info("Collected %d of a possible %d Users.", len(user_objects), total_users)

    users_file.write(']')
    users_file.close()

def _get_group(group_id: str, target_name: str):
    """Attempst to retrieve Group by id.
        
        If the Group exists, retrieve and return the object.
        If not, return None
        
        Args:
        group_id (str): UUID of Group to retrieve
        target_name (str): targetName field value for the specified Group.
        """
    _logger.info(f"Retrieving Group: {target_name}")
    
    # Set our resource URLs
    url = config.xmod_url + '/api/xm/1/groups/' + urllib.parse.quote(group_id) + '?embed=supervisors'
    _logger.debug(f'Attempting to retrieve Group "{target_name}" via url: {url}')
    
    # Make the request
    try:
        response = requests.get(url, auth=config.basic_auth)
    except requests.exceptions.RequestException as e:
        _logger.error(config.ERR_REQUEST_EXCEPTION_CODE, url, repr(e))
        return None
    
    # If the initial response fails, log and return null
    if response.status_code != 200:
        _log_xm_error(url, response)
        return None
    
    # Process the response
    group_obj = response.json()
    # If present, translate the Site from an ID to a name
    if 'site' in group_obj:
        site_name = _lookup_site_name(group_obj['site']['id'])
        del group_obj['site']
        group_obj['site'] = site_name
    # _logger.debug(f'Found Group "{group_obj["targetName"]}" - json body: {pprint.pformat(group_obj)}')
    _logger.debug(f'Found Group "{group_obj["targetName"]}" - json body.id: {group_obj["id"]}')
    return group_obj

def _get_group_shifts(group_id: str, target_name: str):
    """Return a Group's Shifts

    Retrieves the Shift records from xMatters for the specified Group and
    returns them as a list.

    Args:
        group_id (str): The Group's UUID
        target_name (str): The Group's targetName field

    Return:
        shift_list (list): List of dictionaries of the Group's Shifts.
    """
    # Initialize conditions
    shift_list = []
    total_shifts = 0
    url = config.xmod_url + '/api/xm/1/groups/' + group_id + '/shifts/?embed=members&offset=0&limit=' + str(config.page_size)
    _logger.debug(f'Gathering Shifts for Group "{target_name}", url={url}')

    while True:

        # Get the site records
        response = requests.get(url, auth=config.basic_auth)
        if response.status_code not in [200, 404]:
            _log_xm_error(url, response)
            break

        # Process the responses
        bodys = response.json()
        total_shifts = bodys['total']
        if bodys['count'] > 0:
            _logger.debug('%d Count of %d Total Shifts found for Group "%s" via url=%s', bodys['count'], bodys['total'], target_name, url)
            shift_list += bodys['data']

        # See if there are any more to get
        if 'links' in bodys and 'next' in bodys['links']:
            url = config.xmod_url + bodys['links']['next']
        else:
            break

    _logger.debug(f'Collected {len(shift_list)} of a possible {total_shifts} Shifts for Group "{target_name}".')

    return shift_list

def _process_groups():
    """Capture and save the instances Group objects

    Retrieves the User's Group object records from xMatters and saves them in
    JSON payload format to the output file.

    Args:
        None

    Return:
        None
    """
    _logger.info('Begin capturing Groups.')
    groups_file = _create_out_file(config.groups_filename)
    groups_file.write('[\n')

    # Initialize conditions
    group_objects = []
    total_groups = 0
    cnt = 0
    url = config.xmod_url + '/api/xm/1/groups?offset=0&limit=' + str(config.page_size)
    _logger.debug('Gathering Groups via url=%s', url)

    while True:

        # Get the group records
        response = requests.get(url, auth=config.basic_auth)
        if response.status_code not in [200, 404]:
            _log_xm_error(url, response)
            break

        # Process the responses
        bodys = response.json()
        total_groups = bodys['total']
        if bodys['count'] > 0:
            _logger.debug("%d Count of %d Total Groups found via url=%s", bodys['count'], bodys['total'], url)
            for body in bodys['data']:

                # Get the full Group object, including Roles and Supervisors
                group_obj = {}
                a_group = _get_group(body['id'], body['targetName'])
                if a_group is not None:
                    group_obj['group'] = a_group

                    # Get the shifts,
                    group_obj['shifts'] = _get_group_shifts(a_group['id'], a_group['targetName'])

                    # Save the Group
                    cnt += 1
                    json.dump(group_obj, groups_file)
                    groups_file.write(',\n') if cnt < total_groups else groups_file.write('\n')

            group_objects += bodys['data']

        # See if there are any more to get
        if 'links' in bodys and 'next' in bodys['links']:
            url = config.xmod_url + bodys['links']['next']
        else:
            break
            
    _logger.info(f"Collected {len(group_objects)} of a possible {total_groups} Groups.")

    groups_file.write(']')
    groups_file.close()

def _save_admin_data():
    """Saves the collected admin sets

    Preservs the collected admin sets to the file system.

    Args:
        None

    Return:
        None
    """
    admin_dict = {
        'admins': list(_admin_objects['admins']),
        'roles': list(_admin_objects['roles']),
        'timezones': list(_admin_objects['timezones']),
        'countries': list(_admin_objects['countries']),
        'languages': list(_admin_objects['languages']),
        'devices': list(_admin_objects['devices']),
        'usps': list(_admin_objects['usps'])
    }
    admin_file = _create_out_file(config.admin_filename)
    json.dump(admin_dict, admin_file, indent=2)
    admin_file.close()

def process(objects_to_process: list):
    """Capture objects for this instance.

    If requeste contains 'sites', then capture and save Sites.
    If requeste contains 'users', then capture and save Users.
    If requeste contains 'devices', then capture and save Devices.
    If requeste contains 'groups', then capture and save Groups.

    Args:
        objects_to_process (list): The list of object types to capture
    """
    global _logger, _admin_objects # pylint: disable=global-statement

    ### Get the current logger
    _logger = common_logger.get_logger()

    # Initialize the Admin object
    _admin_objects = {
        'admins': set(),
        'roles': set(),
        'timezones': set(),
        'countries': set(),
        'languages': set(),
        'devices': set(),
        'usps': set()
    }

    # Capture and save the Site objects
    if 'sites' in objects_to_process:
        _process_sites()

    # Capture and save the User objects, and possibly devices
    if 'users' in objects_to_process:
        _process_users('devices' in objects_to_process)
    elif 'devices' in objects_to_process:
        _process_users(True)

    # Capture and save the Device objects
    if 'groups' in objects_to_process:
        _process_groups()

    # Preserve the collected admin data
    _save_admin_data()

def main():
    """In case we need to execute the module directly"""
    pass

if __name__ == '__main__':
    main()
