# ********************  STEP 7 ********************
# 7. Call the MapServiceIntersection API, and create a data frame from it
# Create a dataframe for the MapServiceIntersection data: "df_MapServiceIntersections"
# *******************************************************************************************************
# This is where the variable "chosen_feature" gets created and the type of layer you wish to display is entered
import requests

from CP3_API_calls.CP3_API_Classes import CP3Client_API

def MapServiceIntersectionCatalogue(username, password, grant_type, url_choice, API_call_dict, layer_dict, SpatialFeatureChoice):
    global chosen_feature_code
    # 7. Call the MapServiceIntersection API, and create a data frame from it
    # Assign the API_Call variable
    API_Call = 'MapServiceIntersections'
    # Assign the API_return_type variable
    API_return_type = API_call_dict[API_Call][0]
    # Assign the API_parameters variable
    API_parameters = API_call_dict[API_Call][1]
    # ******************************************
    # A choice is made here and hard coded!!!!
    # This is where the specific spatial feature is selected! (layer_list[2] = 176 (wards) in this instance)
    # *******************************************************************************************************
    # layer_list gets constructed in the "MapServiceLayers.py"" module
    # layer_list = [174, 175, 176] therefore layer_list[2] = 176
    # layer_dict = {174: 'City of Tshwane mSCOA Regional Segment', 175: 'City of Tshwane Regions', 176: 'City of Tshwane Wards'}
    # chosen_feature = layer_dict[layer_list[2]] # = 'City of Tshwane Wards'
    # Assign a value to chosen_feature. If SpatialFeatureChoice == "Wards" then chosen_feature = 'City of Tshwane Wards', chosen_feature_code = 176
    for key, item in layer_dict.items():
        if SpatialFeatureChoice in item:
            chosen_feature = item
            chosen_feature_code = key

    API_parameters = {"projectStatus": "Active", "locationType": "WorksLocation", "serviceLayerId": chosen_feature_code}
    # = {'projectStatus': 'Active', 'locationType': 'WorksLocation', 'serviceLayerId': 176}
    # *******************************************************************************************************
    # Call the inherited class
    client2 = CP3Client_API(username, password, grant_type, url_choice, API_Call, API_return_type, API_parameters)
    # You need to call the token again for this inherited class. The 1st token call was for the help API. This one is for the API call
    token2 = client2.get_API_token
    # Assign the returned API to a variable
    try:
        MapServiceIntersections = client2.get_any_API
        return MapServiceIntersections
    except requests.exceptions.JSONDecodeError:  # Picked up this error at Lekwa...
        MapServiceIntersections = {}
        return MapServiceIntersections



