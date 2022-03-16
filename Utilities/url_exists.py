import requests
from requests.exceptions import ConnectionError

def URL_exists(url_string):
    """
    This funtions checks if a URL exists.
    The url must be passed to the function as a string.
    It returns True if it does exist and False if it does not exist.
    It prints a helpful explanetary message to the console in both instances.
    """
    try:
        request = requests.get(url_string)
        if request.status_code == 200:
            #print('Web site exists')
            return True
        else:
            #print('Web site does not exist')
            return False
    except ConnectionError as e:
        #print('Web site does not exist')
        return False