# ********************  STEP 6 ********************
# 6. Call the MapServiceLayerCatalogue API, and create a dictionary of available layers
# Create a dictionary from the available Map Service Layers: "layer_dict"
from CP3_API_calls.CP3_API_Classes import CP3Client_API

def MapServiceLayersCatalogue(username, password, grant_type, url_choice, API_call_dict):
    # 6. Call the MapServiceLayerCatalogue API, and create a dictionary of available layers
    # Assign the API_Call variable
    API_Call = 'MapServiceLayerCatalogue'
    # Assign the API_return_type variable
    API_return_type = API_call_dict[API_Call][0]
    # Assign the API_parameters variable
    API_parameters = API_call_dict[API_Call][1]
    # Call the inherited class
    client2 = CP3Client_API(username, password, grant_type, url_choice, API_Call, API_return_type, API_parameters)
    # You need to call the token again for this inherited class. The 1st token call was for the help API. This one is for the API call
    token2 = client2.get_API_token
    # Assign the returned API to a variable
    MapServiceLayerCatalogue = client2.get_any_API
    # *******************************************************************************************************
    # This gets used to choose which features to show in the report by making a choice in the
    # "MapServiceIntersectionCatalogue.py"" script
    # Create a list and dictionary with the layers that are available (pulled from the API)
    layer_dict = {}
    layer_list = []
    for layer in MapServiceLayerCatalogue:
        layer_list.append(layer['Description'])
        layer_dict[layer['Id']] = layer['Description']
    # This results in:
    # layer_list = ['City of Tshwane mSCOA Regional Segment', .....]
    # layer_dict = {174: 'City of Tshwane mSCOA Regional Segment', 175: 'City of Tshwane Regions', 176: 'City of Tshwane Wards'}

    # Put this into a dataframe
    # df_MapServiceLayerCatalogue = pd.DataFrame(MapServiceLayerCatalogue)
    # del df_MapServiceLayerCatalogue['ForIntersection']
    # del df_MapServiceLayerCatalogue['Grouping']
    """
    This creates a df like this:
    	Id	Description	                            ForIntersection	Grouping
0	    174	City of Tshwane mSCOA Regional Segment	Administrative Boundaries
1	    175	City of Tshwane Regions	    	        Administrative Boundaries
2	    176	City of Tshwane Wards	    	        Administrative Boundaries
    """
    return_list = []
    return_list.append(MapServiceLayerCatalogue)
    return_list.append(layer_dict)
    return_list.append(layer_list)
    return return_list