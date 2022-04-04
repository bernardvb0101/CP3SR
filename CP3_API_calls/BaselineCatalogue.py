# ********************  STEP 3 ********************
# 3. Call the BaselineCatalogue API, to show which baseline applies to this query
# From this process we get 2 more variables required for further calls:
# "baselineAPIAccessTag_var" and "baselineStatus_var"

from CP3_API_calls.CP3_API_Classes import CP3Client_API

# 3. Call the BaselineCatalogue API, to show which baseline applies to this query
def baseline_catalogue(username, password, grant_type, url_choice, API_call_dict):
    # Assign the API_Call variable
    API_Call = 'BaselineCatalogue'
    # Assign the API_return_type variable
    API_return_type = API_call_dict[API_Call][0]
    # Assign the API_parameters variable
    API_parameters = API_call_dict[API_Call][1]
    # Call the inherited class
    client2 = CP3Client_API(username, password, grant_type, url_choice, API_Call, API_return_type, API_parameters)
    # You need to call the token again for this inherited class. The 1st token call was for the help API. This one is for the API call
    token2 = client2.get_API_token
    # Assign the returned API to a variable
    BaselineCatalogue = client2.get_any_API
    # So now we have 3 more variables required for further calls:
    """
    BaselineCatalogue looks like this:
    [{'APIAccessTag': 'DEMAND2022', 'Name': '2021/22 Planning + Rollover (20200901)', 
    'Description': 'Duplicated from the 2020/21 Planning + Final Annexure A (20200625) Baseline', 
    'FinancialYear': 2022, 'IsActive': False}]
    To extract something:
    BaselineCatalogue[0]['APIAccessTag'] -> 'DEMAND2022'
    """
    return BaselineCatalogue[0]
