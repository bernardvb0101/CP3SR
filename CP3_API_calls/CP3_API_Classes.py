# ********************  STEP 1 ********************
# 1. Scaffolding to enable API calls (Part 1 - Create API Classes)
# a. Create an API calling class to call basics such as tokens or help files¶
# b. Create a secondary calling class to enable API calls which include parameters

import requests
from cachetools import cached, TTLCache


# a) Create an API calling class to call basics such as tokens or help files
class CP3Client:
    """
    This is an API calling client built specifically for calling CP3 APIs.
    This specific client is used to call APIs not requiring any parameters.
    The client allows you to get an API token with the get_API_token(self)method
    and to get the API help file with the get_API_help(self) method.
    """

    def __init__(self, username, password, grant_type, BASE_URL):
        self.username = username
        self.password = password
        self.gant_type = grant_type
        self.base_url = BASE_URL
        self.bearer_token = ""

    # This method calls the token required for the help API and all further API calls.
    @property
    @cached(cache=TTLCache(maxsize=2, ttl=86400))  # held in cache for 86400 seconds (24 hours)
    def get_API_token(self):
        self.body = {'username': self.username,
                     'password': self.password,
                     'grant_type': 'password'}
        self.headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        try:
            self.bearer_token = (requests.post(f"{self.base_url}/token", data=self.body, headers=self.headers)).json()["access_token"]
        except KeyError:
            self.bearer_token = "ERROR: Unsuccessful token retrieval on API call. Please check username and password or API permissions on the CP3 system."
        return self.bearer_token

    # This method is used to call the help file API - the content gets used to call all further APIs
    @property
    @cached(cache=TTLCache(maxsize=2, ttl=864400))  # held in cache for 86400 seconds (24 hours)
    def get_API_help(self):
        self.headers = {'Authorization': 'Bearer ' + self.bearer_token}
        return (requests.get(f"{self.base_url}/api/help", headers=self.headers)).text


# b) Create a secondary calling class to enable API calls which include parameters
class CP3Client_API(CP3Client):
    """
    This is a secondary API calling client inheriting from the 'CP3Client' main client.
    It only has one method allowing a user to also pass parameters.
    All methods available in the main client is also callable using this secondary client.
    """

    def __init__(self, username, password, grant_type, BASE_URL, API_Call, API_return_type, API_parameters):
        super().__init__(username, password, grant_type, BASE_URL)
        self.API_Call = API_Call
        self.API_return_type = API_return_type
        self.API_parameters = API_parameters

    @property
    @cached(cache=TTLCache(maxsize=2, ttl=864400))  # held in cache for 86400 seconds (24 hours)
    def get_any_API(self):
        self.headers = {'Authorization': 'Bearer ' + self.bearer_token}
        if self.API_parameters == 'No parameters':  # No parameter dictionary needs to be passed.
            self.API_parameters = {}  # Change parameters into an empty dictionary
        if self.API_return_type == '(application/JSON)':
            return (requests.get(f"{self.base_url}/api/{self.API_Call}", headers=self.headers,
                                 params=self.API_parameters)).json()
        else:
            return (requests.get(f"{self.base_url}/api/{self.API_Call}", headers=self.headers,
                                 params=self.API_parameters)).text