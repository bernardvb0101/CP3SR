# ********************  STEP 5 ********************
# 5. Call the CapExDemandCatalogue API, and create a CapexDemand Catalogue data frame from this call
# Create a data frame for the Capex Demand Catalogue: The name is "df_CapexBudgetDemandCatalogue"
from CP3_API_calls.CP3_API_Classes import CP3Client_API

def CapexDemandCatalogue(username, password, grant_type, full_url, API_call_dict, baselineAPIAccessTag_var):
    # 5. Call the CapExDemandCatalogue API, and create a CapexDemand Catalogue data frame from this call
    # Assign the API_Call variable
    API_Call = 'CapexBudgetDemandCatalogue'
    # Assign the API_return_type variable
    API_return_type = API_call_dict[API_Call][0]
    # Assign the API_parameters variable
    API_parameters = API_call_dict[API_Call][1]
    # Pass it in like so:
    API_parameters = {'baselineAPIAccessTag': baselineAPIAccessTag_var}
    # Call the inherited class
    client2 = CP3Client_API(username, password, grant_type, full_url, API_Call, API_return_type, API_parameters)
    # You need to call the token again for this inherited class. The 1st token call was for the help API. This one is for the API call
    token2 = client2.get_API_token
    # Assign the returned API to a variable
    CapexBudgetDemandCatalogue = client2.get_any_API
    """
    Creates a dataframe that looks like this:
    	ProjectId	BaselineAPIAccessTag	BaselineName	BaselineStatus	ProjectStatus	Interval	FundingSourceName	FundingGUID	LifecyclePhaseName	LifecyclePhaseActiveInactive	LifecycleSubPhaseName	LifecycleSubPhaseActiveInactive	Amount
0	712975302	DEMAND2022	2021/22 Planning + Rollover (20200901)	IsCommitted	Active	2022	001 Council Funding	c1f5b8f4-9b3f-41ad-a76d-9536639895b5	Goods + services: Purchasing	Active	Invoices	Active	2648700.0
    Etc.....
    """
    # Pass the JSON into a pandas dataframe
    # df_CapexBudgetDemandCatalogue = pd.DataFrame(CapexBudgetDemandCatalogue)
    return CapexBudgetDemandCatalogue
