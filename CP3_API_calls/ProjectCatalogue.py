# ********************  STEP 4 ********************
# 4. Call the ProjectCatalogue API, and create a Project Catalogue data frame from this call
# Create a data frame for the Project Catalogue: The name is "df_ProjectCatalogue"

from CP3_API_calls.CP3_API_Classes import CP3Client_API

def ProjectCatalogue(username, password, grant_type, url_choice, API_call_dict):
    # 4. Call the ProjectCatalogue API, and create a Project Catalogue data frame from this call
    # Assign the API_Call variable
    API_Call = 'ProjectCatalogue'
    # Assign the API_return_type variable
    API_return_type = API_call_dict[API_Call][0]
    # Assign the API_parameters variable
    API_parameters = API_call_dict[API_Call][1]
    # BUT we dont want to pass parameters so change to: API_parameters = 'No parameters'
    API_parameters = 'No parameters'
    # Call the inherited class
    client2 = CP3Client_API(username, password, grant_type, url_choice, API_Call, API_return_type, API_parameters)
    # You need to call the token again for this inherited class. The 1st token call was for the help API. This one is for the API call
    token2 = client2.get_API_token
    # Assign the returned API to a variable
    ProjectCatalogue = client2.get_any_API
    # Pass the JSON into a pandas dataframe
    """
    Creates a dataframe that looks like this:
    ProjectId	Name	                                            Description	ParentId	ParentName	                        ProjectStatus	UnitName	                    DepartmentName
0	712974517	Tshwane Leadership & Management Academy: Carports	Carports	712973924	(712953B) Renovation of Facility	Active	        Group Human Capital Management	Tshwane Leadership and Management Academ
    Etc......
    """
    # The dataframe gets create where the call is made to this function
    # df_ProjectCatalogue = pd.DataFrame(ProjectCatalogue)
    return ProjectCatalogue

