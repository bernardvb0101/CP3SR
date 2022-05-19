from flask import Flask, render_template, request, flash, send_file
import pandas as pd
import numpy as np
import json
import os
from werkzeug.utils import secure_filename
from waitress import serve
from Utilities.file_exists import allowed_file
from Utilities.url_exists import URL_exists
from Utilities.control_growth import control_growth_of_docx, control_growth_of_xlsx
from Utilities.get_url_vars import vars_from_json_file, read_source_file
from Utilities.Convert_to_Excel import write_to_xls
from Utilities.Convert_to_JSON import write_to_JSON_file
from Utilities.Read_from_Excel import read_from_excel
from Utilities.Read_from_JSON import read_json_file
from Utilities.get_API_vars import API_vars_json_file
from CP3_API_calls.Create_API_Variables import create_vars
from CP3_API_calls.BaselineCatalogue import baseline_catalogue
from CP3_API_calls.ProjectCatalogue import ProjectCatalogue
from CP3_API_calls.CapexDemandCatalogue import CapexDemandCatalogue
from CP3_API_calls.MapServiceLayersCatalogue import MapServiceLayersCatalogue
from CP3_API_calls.MapServiceIntersectionCatalogue import MapServiceIntersectionCatalogue
from ReportFactory.create_reports import create_worddoc

# global variables
# Get them from the variables json file
json_file_name = "CP3_deployments.json"
url_vars_file_path = f"./Variables/{json_file_name}"
returned_combined_list = vars_from_json_file(url_vars_file_path)

org_list = returned_combined_list[1]
entityname_list = returned_combined_list[2]
url_list = returned_combined_list[3]
json_file_ok = returned_combined_list[0][0]

# Get API creds from the variables json file
API_file_name = "API_Profile.json"
API_vars_file_path = f"./Variables/{API_file_name}"
API_vars = API_vars_json_file(API_vars_file_path)

username = API_vars[0]
password = API_vars[1]
grant_type = API_vars[2]

nav_stage = 1
all_well = 0
SpatialFeatureChoice = ""

if json_file_ok:
    org_choice = org_list[0]  # Make the 1st one in the list the default
    url_choice = url_list[0]  # Make the 1st one in the list the default
    entity_choice = entityname_list[0]  # Make the 1st one in the list the default
else:
    org_choice = []
    url_choice = []
    entity_choice = []

spatial_var = []
# sys_username = "Bernard"
# This part of the credentials for the API call (to my understanding) is default allways "password"


app = Flask(__name__)  # to make the app run without any

#app.config['SECRET_KEY'] = os.urandom(24)
# Configure a secret key for flask app
app.config['SECRET_KEY'] = '/sdasd!@#CVVWRER12_'
DOWNLOAD_FOLDER = "/DOWNLOAD_FOLDER"
excel_file_path = f".{DOWNLOAD_FOLDER}/CP3_sites.xlsx"
# Configure an download folder for flask app
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
UPLOAD_FOLDER = "./Variables"
# Configure an upload folder for flask app
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# Set 1 MB file upload limit
app.config["MAX_CONTENT_LENGTH"] = 1 * 1024 * 1024


#  **********************************************************************************
@app.route('/cp3report', methods=['GET'])
def cp3report():
    """
    This function eliminates the user interface and does the same as the user interface version but now through
    a callable API.
    Call it like this:
    http://localhost:8080/cp3report?url_key=tshwane&feature_key=Regions
    """
    global all_well, url_list, spatial_var, username, password, grant_type, org_choice
    global API_call_dict, layer_dict, SpecificFeature, spatial_var, entityname_list
    global layer_list, number_of_plots, entity_choice, url_message, profile_name, profile_pw
    global baseline_cat_dict, df_ProjectCatalogue, df_CapexBudgetDemandCatalogue, df_MapServiceLayerCatalogue
    global df_MapServiceIntersections, no_intersects, total_datapoints, intersecting, df_Intersects2
    global url_choice, SpatialFeatureChoice, master_dict

    master_dict = {}
    SpatialFeatureChoice = ''
    # Get the url_key parameter
    url_key = request.args.get('url_key')
    # Get the feature_key parameter
    feature_key = request.args.get('feature_key')
    profile_name = request.args.get('profile_name')
    profile_pw = request.args.get('profile_pw')

    if profile_name == None or profile_pw == None:  # Username and/or password was not provided
        # If nothing was given, use the defaults
        profile_name = username
        profile_pw = password


    # If both of the keys are not None, it means a spatial report is called. Unless invalid params were used
    if url_key != None and feature_key != None:
        url_key = url_key.lower()
        feature_key = feature_key.replace('%20', ' ').strip()  # Maybe R is passing the %20's - who knows.

        if url_key == 'johannesburg':
            url_key = 'joburg'
        if url_key == 'johanesburg':
            url_key = 'joburg'
        if url_key == 'johannisburg':
            url_key = 'joburg'
        if url_key == 'joberg':
            url_key = 'joburg'
        if url_key == 'johannesberg':
            url_key = 'joburg'
        if url_key == 'pretoria':
            url_key = 'tshwane'
        if url_key == 'tshane':
            url_key = 'tshwane'
        if url_key == 'tswane':
            url_key = 'tshwane'

        # url_choice = 'https://www.tshwane.cp3.co.za/'
        url_choice = f"https://www.{url_key}.cp3.co.za/"

        # The all_well variable is zero if no APIs were returned successfully yet
        all_well = 0
        # The spatial_var is a list with the spatial features that the user cna select from later on
        spatial_var = []

        # Populate the entity name for use in the report
        keep_track = 0  # temp variable to enable 'org_choice' variable to match 'entity_choice'
        for url_option in url_list:
            if url_option == url_choice:
                entity_choice = entityname_list[keep_track]
                org_choice = org_list[keep_track]
                break
            keep_track += 1

        # Call the Help API to create variables
        API_call_dict = create_vars(profile_name, profile_pw, grant_type, url_choice)
        if not API_call_dict:  # If the dictionary is returned empty, tell the user
            url_message = "Something went wrong with the API calls. This may be a wrong username/password OR the " \
                          "API credentials are not set up correctly/at all."
        else:  # Returned values from API variables call seems ok so process can continue
            all_well += 1
            # Call the Baseline Catalogue and get info about the baseline
            baseline_cat_dict = baseline_catalogue(profile_name, profile_pw, grant_type, url_choice, API_call_dict)
            if not baseline_cat_dict:
                url_message = "Something went wrong with the BaselineCatalogue API call."
            else:
                all_well += 1
            # Call the ProjectCatalogue and create a dataframe
            df_ProjectCatalogue = pd.DataFrame(ProjectCatalogue(profile_name, profile_pw, grant_type, url_choice,
                                                                API_call_dict))
            if df_ProjectCatalogue.empty:
                url_message = "Something went wrong with the ProjectCatalogue API call."
            else:
                all_well += 1
            # Call the CapexDemandCatalogue and create a dataframe
            df_CapexBudgetDemandCatalogue = pd.DataFrame(CapexDemandCatalogue(profile_name, profile_pw, grant_type,
                                                                              url_choice, API_call_dict,
                                                                              baseline_cat_dict['APIAccessTag']))
            if df_CapexBudgetDemandCatalogue.empty:
                url_message = "Something went wrong with the CapexBudgetDemandCatalogue API call."
            else:
                all_well += 1
            # Call the MapServiceLayersCatalogue
            return_list = MapServiceLayersCatalogue(profile_name, profile_pw, grant_type, url_choice, API_call_dict)

            df_MapServiceLayerCatalogue = pd.DataFrame(return_list[0])
            layer_dict = return_list[1]
            layer_list = return_list[2]
            SpatialFeatureChoice = ''

            # Get the proper layer name from the dictionary
            for content in layer_dict.values():
                # if feature_key in content.lower():
                if feature_key == content:
                    # print(feature_key)
                    SpatialFeatureChoice = content

            del df_MapServiceLayerCatalogue['ForIntersection']
            del df_MapServiceLayerCatalogue['Grouping']
            if df_MapServiceLayerCatalogue.empty:
                url_message = "Something went wrong with the MapServicesLayersCatalogue API call."
            else:
                all_well += 1
                spatial_var.append(layer_list)

            # If all the API's were called successfully
            if all_well != 5 or SpatialFeatureChoice == '':
                if all_well != 5:
                    url_message = f"The call for data from the {org_choice} system was not successful. This may be because the" \
                                  f"API profile for {org_choice} relating to the username and password that you used, may not " \
                                  f"be correctly set up or one or more of the data catalogues returned an 'empty' response. " \
                                  f"Try using a different site and query to see if the problem persists or whether it is " \
                                  f"specific to this site and your user profile replated to this site."
                if SpatialFeatureChoice == '':
                    url_message = f"The specified 'feature_key' parameter you have given was {feature_key}. There is " \
                                  f"no match for that key and a report could not be generated."
                master_dict["message"] = url_message
                return json.dumps(master_dict["message"], indent=4)
            else:
                url_message = f"Successfull API call on {org_choice} CP3 system.\n2nd API call initiated on" \
                              f" {SpatialFeatureChoice}"

                df_MapServiceIntersections = pd.DataFrame(
                            MapServiceIntersectionCatalogue(profile_name, profile_pw, grant_type,
                                                            url_choice, API_call_dict, layer_dict,
                                                            SpatialFeatureChoice))

                # All APIs have now been called so the report building can proceed
                # Change the FeatureClassName column to "category" type
                df_MapServiceIntersections['FeatureClassName'].astype("category")
                # See how many 'no intersects' are there
                no_intersects = df_MapServiceIntersections[
                    df_MapServiceIntersections['FeatureClassName'] == 'No Intersect'].count()[0]
                # Out of...
                total_datapoints = len(df_MapServiceIntersections.index)
                if no_intersects < total_datapoints:  # Check if there are spatial intersects
                    # Continue as normal if there are spatial intersects
                    all_well += 1  # This really is just for problem/error trapping at this stage, this var does not get used anymore
                    intersecting = total_datapoints - no_intersects
                    # Create a small dataframe (this if for plotly)
                    data = {'NoIntersects': no_intersects, 'Intersecting': intersecting}
                    df_Intersects = pd.DataFrame(data, index=[0])  # the `index` argument is important
                    df_Intersects2 = df_Intersects.T
                    df_Intersects2['Projects'] = df_Intersects2[0]
                    del df_Intersects2[0]
                    df_Intersects2['Intersects'] = df_Intersects2.index
                    # Create dictionary with spatial features as key and projects in that feature with their % intersect as values
                    # Create a list containing the unique Spatial entities available in the dataset. This will enable an iteration through them later
                    list_of_features = list(df_MapServiceIntersections['FeatureClassName'].unique())
                    # Remove 'no intersect' because they have no spatial property and the overwhelm ito numbers in many datasets
                    list_of_features.remove('No Intersect')
                    # Create a list for all the fin years in the system
                    list_of_years = list(df_CapexBudgetDemandCatalogue['Interval'].unique())
                    list_of_years = sorted(list_of_years)
                    # This variable stores how many of the chosen feature there is
                    chosen_feature_qty = len(list_of_features)

                    # Initialise the master dictionary
                    feature_intersect_dict = {}
                    # Create a master dictionary with each spatial feature as a key
                    # Each key (e.g. ward) contains another dictionary with project number as key and percentage intersect as value
                    sub_frame = {}
                    for feature in list_of_features:
                        sub_frame[feature] = df_MapServiceIntersections[
                            df_MapServiceIntersections["FeatureClassName"] == feature]
                        project_dict = {}
                        for row in sub_frame[feature].itertuples():
                            project_dict[row.ProjectId] = row.PercentageIntersect
                        feature_intersect_dict[feature] = project_dict
                    # Now this dictionary can be used to query spatial feature to get to the projects and their intersects.
                    # print(feature_intersect_dict['Ward 100']) -> All projects in a ward is given with their % intersect

                    # Put this entire dictionary in a dataframe. Projects with their %Intersect per ChosenSpatialFeature
                    df_FeatureIntersectPer = pd.DataFrame(feature_intersect_dict)
                    # Replace all the NaN's with zeros
                    df_FeatureIntersectPer.replace(np.nan, 0, inplace=True)

                    # ********************************************************************************************
                    # Build df_EntireSet with columns for each fin year
                    # 1) Chosen Feature 2) No of Projs in Chosen Feature 3) Total Capital Demand in Chosen Feature 4) Capital Demand in Chosen Feature per Year
                    list_nr = []  # List with number of projects in a ward
                    list_cost = {}  # list_cost[2022], etc
                    list_cost['Total'] = []
                    list_cost['MTREF'] = []
                    for year in list_of_years:
                        list_cost[year] = []
                    mask3_dict = {}
                    # There will bae a dataframe for each mask for each year
                    df_mask3_dict = {}  # The will be a dataframe for each year so they will be store in a dictionary
                    df_perward = {}
                    project_year_total = {}

                    for feature in list_of_features:
                        # The number of project per chosen spatial feature goes into the 'list_nr' list
                        list_nr.append(len(
                            df_MapServiceIntersections[df_MapServiceIntersections['FeatureClassName'] == feature][
                                'ProjectId']))
                        # Now get the list of projects within the chosen feature (per feature)
                        df_perward[feature] = df_FeatureIntersectPer[df_FeatureIntersectPer[feature] > 0][feature]
                        # Set the variables that will keep track of the budget per spatial feature to zero
                        project_total = 0
                        mtref_total = 0

                        for year in list_of_years:
                            project_year_total[year] = 0

                        # Now iterate through the project numbers
                        for index in df_perward[feature].index:  # 'index' is the project number
                            # The 'df_perward[feature]' ensures the project exist in that ward
                            # Create a mask for the project number represented by 'index'
                            mask1 = df_CapexBudgetDemandCatalogue['ProjectId'] == index  # Mask per project
                            df_mask1 = df_CapexBudgetDemandCatalogue[
                                mask1]  # A dataframe for that mask for just that project ('index')

                            # mask3 is for a specific year
                            for year in list_of_years:
                                mask3_dict[year] = df_mask1['Interval'] == year  # There is a mask for each year
                                df_mask3_dict[year] = df_mask1[mask3_dict[
                                    year]]  # Create a dataframe for each year 1) Project -> 2) Year
                                try:
                                    project_year_total[year] += float(df_mask3_dict[year]['Amount'].sum()) * \
                                                                df_perward[feature][index]
                                    project_total += float(df_mask3_dict[year]['Amount'].sum()) * df_perward[feature][
                                        index]
                                    if year in list_of_years[0:3]:
                                        mtref_total += float(df_mask3_dict[year]['Amount'].sum()) * \
                                                       df_perward[feature][index]
                                except IndexError:  # If there is no year, an index error is returned. The populate the list with a zero.
                                    project_year_total[year] += 0
                                    project_total += 0
                                    if year in list_of_years[0:3]:
                                        mtref_total += 0

                        list_cost['Total'].append(project_total)
                        list_cost['MTREF'].append(mtref_total)
                        for year in list_of_years:
                            list_cost[year].append(project_year_total[year])

                    # With these lists a new dataframe can be created to plot
                    # Thus, create a dataframe/dataframes containing all the spatial feautures selected, each containing
                    # the number of projects in that feature and the capital demand per that feature
                    # Decide on the number of data sets depending on the size of the data

                    # 1st Create Dataset of all the data to split up for graphing purposes
                    df_EntireSet = pd.DataFrame(
                        {SpatialFeatureChoice: list_of_features, f'Projects per {SpatialFeatureChoice}': list_nr,
                         'Capital All Years': list_cost['Total'], 'Capital MTREF': list_cost['MTREF']})
                    column_name_list = []
                    for year in list_of_years:
                        column_name_list.append(f'Capital {year}')
                        df_EntireSet[f'Capital {year}'] = list_cost[year]


                    # Now sort the dataset in order of number of projects from largest to smallest
                    df_EntireSet.sort_values(f'Projects per {SpatialFeatureChoice}', inplace=True, ascending=True)
                    # Modify df_EntireSet by adding columns to rank
                    df_EntireSet["CapAllRank"] = df_EntireSet['Capital All Years'].rank(ascending=False)
                    df_EntireSet["CapMTREFRank"] = df_EntireSet['Capital MTREF'].rank(ascending=False)
                    df_EntireSet["NoProjectsRank"] = df_EntireSet[f'Projects per {SpatialFeatureChoice}'].rank(
                        ascending=False)

                    # Wrap all the loose variables in a dictionary for use in the report
                    var_dict = {}
                    # var_dict['username'] = sys_username
                    var_dict['url_choice'] = url_choice
                    var_dict['org_choice'] = org_choice
                    var_dict['entity_choice'] = entity_choice
                    var_dict['SpatialFeatureChoice'] = SpatialFeatureChoice
                    var_dict['Layer_List'] = layer_list
                    var_dict['total_datapoints'] = total_datapoints
                    var_dict['intersecting'] = intersecting
                    var_dict['no_intersects'] = no_intersects
                    var_dict['chosen_feature_qty'] = chosen_feature_qty
                    var_dict['column_name_list'] = column_name_list
                    var_dict['list_of_years'] = list_of_years
                    var_dict['list_of_features'] = list_of_features


                    # Check the growth of files and reduce it
                    control_growth_of_docx()
                    # Now create the spatial feature report
                    path = create_worddoc(var_dict=var_dict, baseline_dict=baseline_cat_dict,
                                          df_intersects2=df_Intersects2,
                                          df_EntireSet=df_EntireSet, df_perward=df_perward,
                                          df_MapServiceIntersections=df_MapServiceIntersections,
                                          df_CapexBudgetDemandCatalogue=df_CapexBudgetDemandCatalogue,
                                          df_ProjectCatalogue=df_ProjectCatalogue)
                    return send_file(path, as_attachment=True)
                else:
                    url_message = "There are no spatial intersects"
                    master_dict["message"] = url_message
                    return json.dumps(master_dict["message"], indent=4)
    else: # url_key or featue_key is empty
        # If the url_key and feature_key vars are empty, get the site and layer parameters
        site = request.args.get('site')
        layer = request.args.get('layer')

        if site != None:
            site = site.lower()

            return_list = read_source_file(url_vars_file_path)
            file_status = return_list[0]
            jsn_var = return_list[1]

            master_dict = {}
            if file_status:  # If nothing is wrong with the file
                for entity in jsn_var:
                    entity_dict = {}
                    entity_dict["url_key"] = entity['org'].lower()
                    master_dict[entity_dict["url_key"].lower()] = entity_dict
            else:  # Something is wrong with the source json file
                master_dict["message"] = jsn_var  # Returns an error message saying the file is buggered

            try:
                if master_dict["message"]:  # This just triggers a KeyError if this entry does not exist.
                    return json.dumps(master_dict["message"], indent=4)  # Tell the user went wrong.
            except KeyError:  # There was no problem the process can go ahead
                if site == 'help':
                    master_dict["message"] = "<b>CP3 Spatial Report API Parameter Options:<br>Note: Passing the " \
                                             "'profile_name' and 'profile_pw' parameters are optional.<br>If " \
                                             "no 'profile_name' and 'profile_pw' parameters are provided, the defaul" \
                                             "ts that can be " \
                                             "managed on the app management screen (https://connect.cp3.co.za/docx_s" \
                                             "patial_report) under the 'Maintenance' menu will be used.<br><br>" \
                                             "1. To find out what the 'url_key' parameter is for the '/cp3report' " \
                                             "API use the following parameter:<br>'https://connect.cp3.co.za/docx_sp" \
                                             "atial_report/cp3report?site=all'<br>or<br>'https://connect.cp3.co.za/d" \
                                             "ocx_spatial_report/cp3report?site=name' (where name is for example 'ts" \
                                             "hwane').<br><br>2. To find out what the spatial layer options are use " \
                                             "the 'site' AND 'layer' parameters:<br>Full example 1:<br>'https://conne" \
                                             "ct.cp3.co.za/docx_spatial_report/cp3report?site=tshwane&layer=list'" \
                                             "<br>Full example 2:<br>'https://connect.cp3.co.za/docx_spatial_report/" \
                                             "cp3report?site=all'<br>Full example 3:<br>'https://connect.cp3.co.za/d" \
                                             "ocx_spatial_report/cp3report?profile_name=....&profile_pw=.....&site=a" \
                                             "ll'<br><br>3. To download a spatial report:<br>Full example 1:" \
                                             "<br>'https://connect.cp3.co.za/docx_spatial_report/cp3report?url_key=" \
                                             "midvaal&feature_key=Midvaal 5 EDCs'<br>Full example 2:<br>'https://co" \
                                             "nnect.cp3.co.za/docx_spatial_report/cp3report?url_key=midvaal&feature" \
                                             "_key=Midvaal 5 EDCs&profile_name=....&profile_pw=....'"

                    json_object = json.dumps(master_dict["message"], indent=4)
                    return json_object
                elif (site != 'help' and site != None) or (site == 'all'):  # A site's parameter was requested
                    if site == 'all':  # User wants a list of all available url_key's
                        return json.dumps(master_dict, indent=4)
                    else:  # user entered a specific site
                        # Try and fix some common input errors
                        if site == 'joburg':
                            site = 'johannesburg'
                        if site == 'johanesburg':
                            site = 'johannesburg'
                        if site == 'johannisburg':
                            site = 'johannesburg'
                        if site == 'joberg':
                            site = 'johannesburg'
                        if site == 'johannesberg':
                            site = 'johannesburg'
                        if site == 'pretoria':
                            site = 'tshwane'
                        if site == 'tshane':
                            site = 'tshwane'
                        if site == 'tswane':
                            site = 'tshwane'
                        if site.lower() in master_dict.keys():
                            if layer == None:  # No layer was asked to you can return the url_key
                                return json.dumps(master_dict[site.lower()], indent=4)
                            else:  # A layer was asked
                                if layer == 'list':
                                    url_choice = f"https://www.{site}.cp3.co.za/"
                                    API_call_dict = create_vars(profile_name, profile_pw, grant_type, url_choice)
                                    if not API_call_dict:  # If the dictionary is returned empty, tell the user
                                        master_dict[
                                            "message"] = "Something went wrong with the APi call to the CP3 site."
                                        return json.dumps(master_dict["message"], indent=4)
                                    else:  # API calling to CP3 right works
                                        # Call the MapServiceLayersCatalogue
                                        return_list = MapServiceLayersCatalogue(profile_name, profile_pw, grant_type,
                                                                                url_choice,
                                                                                API_call_dict)
                                        layer_list = return_list[2]
                                        return_dict = {}
                                        temp_count = 1
                                        for item in layer_list:
                                            return_dict[f"'feature_key' parameter option {temp_count}"] = item
                                            temp_count += 1
                                        return json.dumps(return_dict, indent=4)
                                else:
                                    master_dict["message"] = "Layer query for a specific site must be 'list'."
                                    return json.dumps(master_dict["message"], indent=4)
                        else:  # A valid site could not be found
                            master_dict["message"] = "No match for your specified site was found. The CP3 site may " \
                                                     "exist but if you get this message this particular site has not " \
                                                     "yet been included for API Spatial Reports."
                            return json.dumps(master_dict["message"], indent=4)
        else:  # Site returned None
            master_dict = {}
            master_dict["message"] = "Use:<br>'https://connect.cp3.co.za/docx_spatial_report/cp3report?site=help'<br>" \
                                     "to get guidance."
            return json.dumps(master_dict["message"], indent=4)


# This route is the "home" route that redirects immediately to "home_in.html"
@app.route('/', methods=['GET', 'POST'])
def home():
    """
    This function allows the user to select from a list of CP3 sites and then from a list of spatial layers from
    that CP3 site. The selectable Spatial Layers are the ones that are associated with the user's API profile.
    A MSWord document is returned containing an analysis of the spatial feature that was selected.
    """
    global all_well, url_choice, url_list, spatial_var, username, password, grant_type, org_choice
    global returned_combined_list, org_list, json_file_ok
    global API_call_dict, layer_dict, SpatialFeatureChoice, SpecificFeature, spatial_var, entityname_list
    global layer_list, number_of_plots, nav_stage, entity_choice, excel_file_path, API_vars_file_path
    global baseline_cat_dict, df_ProjectCatalogue, df_CapexBudgetDemandCatalogue, df_MapServiceLayerCatalogue
    global df_MapServiceIntersections, no_intersects, total_datapoints, intersecting, df_Intersects2

    if json_file_ok:
        if request.method == 'POST':  # Step 1: Pressed the submit button with username and pw, no dropdown yet

            # Get the status of the buttons that were pressed
            button_1stAPI = request.form.get("call_1st_APIs")
            button_2ndAPI = request.form.get("call_2nd_APIs")
            button_dlreport = request.form.get("download_report")
            button_url_upload = request.form.get("Upload_Sites")
            button__url_download = request.form.get("Download_Sites")
            button_api_upload = request.form.get("Upload_Credentials")
            button__api_download = request.form.get("Download_Credentials")

            if button_1stAPI is not None and nav_stage == 1:  # Pressed the button to call an API
                # The all_well variable is zero if no APIs were returned successfully yet
                all_well = 0
                # The spatial_var is a list with the spatial feautures that the user cna select from later on
                spatial_var = []

                # Read the url chosen from the radio button choice that was made (it defaults on the 1st one)
                url_choice = request.form['flexRadioDefault']

                # Populate the entity name for use in the report
                keep_track = 0  # temp variable to enable 'org_choice' variable to match 'entity_choice'
                for url_option in url_list:
                    if url_option == url_choice:
                        entity_choice = entityname_list[keep_track]
                        org_choice = org_list[keep_track]
                        break
                    keep_track += 1

                if username != "" and password != "":  # Credentials for API calling can not be empty
                    # Test the URL that was selected to see if it is active
                    if URL_exists(url_choice):
                        url_ok = True
                    else:
                        url_message = f"{url_choice} was tested and did not return an active response as expected."
                        url_ok = False
                        flash(url_message)
                    if url_ok:
                        # Call the Help API to create variables
                        API_call_dict = create_vars(username, password,grant_type, url_choice)
                        if not API_call_dict:  # If the dictionary is returned empty, tell the user
                            flash("Something went wrong with the API calls.")
                            flash("This may be a wrong username/password OR the API credentials are not set up correctly/at all.")
                        else:  # Returned values from API variables call seems ok so process can continue
                            all_well += 1
                            # Call the Baseline Catalogue and get info about the baseline
                            baseline_cat_dict = baseline_catalogue(username, password, grant_type, url_choice, API_call_dict)
                            if not baseline_cat_dict:
                                flash("Something went wrong with the BaselineCatalogue API call.")
                            else:
                                all_well += 1
                            # Call the ProjectCatalogue and create a dataframe
                            df_ProjectCatalogue = pd.DataFrame(ProjectCatalogue(username,password, grant_type, url_choice, API_call_dict))
                            if df_ProjectCatalogue.empty:
                                flash("Something went wrong with the ProjectCatalogue API call.")
                            else:
                                all_well += 1
                            # Call the CapexDemandCatalogue and create a dataframe
                            df_CapexBudgetDemandCatalogue = pd.DataFrame(CapexDemandCatalogue(username, password, grant_type, url_choice, API_call_dict, baseline_cat_dict['APIAccessTag']))
                            if df_CapexBudgetDemandCatalogue.empty:
                                flash("Something went wrong with the CapexBudgetDemandCatalogue API call.")
                            else:
                                all_well += 1
                            # Call the MapServiceLayersCatalogue
                            return_list = MapServiceLayersCatalogue(username, password, grant_type, url_choice, API_call_dict)

                            df_MapServiceLayerCatalogue = pd.DataFrame(return_list[0])
                            layer_dict = return_list[1]
                            layer_list = return_list[2]

                            del df_MapServiceLayerCatalogue['ForIntersection']
                            del df_MapServiceLayerCatalogue['Grouping']
                            if df_MapServiceLayerCatalogue.empty:
                                flash("Something went wrong with the MapServicesLayersCatalogue API call.")
                            else:
                                all_well += 1
                                spatial_var.append(layer_list)

                else: # Pressed the submit button with empty username and/or password field
                    flash("Username and/or Password field is empty.")

                # If all the API's were called successfully, show the DropDown
                if all_well == 5:
                    flash(f"Successfull API call on {org_choice} CP3 system.\n"
                          f"Select spatial feature and initiate 2nd API call.")
                    nav_stage = 2
                else:
                    flash(f"The call for data from the {org_choice} system was not successful. This may be because the"
                          f" API profile for {org_choice} relating to the username and password that you used, may not"
                          f"be correctly set up or one or more of the data catalogues returned an 'empty' response."
                          f" Try using a different site and query to see if the problem persists or whether it is "
                          f"specific to this site and your user profile replated to this site.")
                # If nav_stage is 2, the user would be able to call the 2nd APIs, otherwise he will just get a fault
                # message and stay on the page with the url choices and username and password
                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var,
                                       SpatialFeatureChoice=SpatialFeatureChoice)

            elif button_1stAPI is not None and (nav_stage == 2 or nav_stage ==3):  # Pressed the button to select a another site
                nav_stage = 1
                SpatialFeatureChoice = ""
                flash("Select another site.")
                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var,
                                       SpatialFeatureChoice=SpatialFeatureChoice)

            elif button_2ndAPI is not None and (nav_stage == 2 or nav_stage ==3): # Calling the 2nd set of APIs
                SpatialFeatureChoice = request.form['inputGroupSelect01']
                if SpatialFeatureChoice != "Choose...":  # User selected a spatial feature
                    # Call the MapServiceIntersectionCatalogue - it can only be called now that the preferred spatial
                    # feature is selected by the user (SpatialFeatureChoice).
                    df_MapServiceIntersections = pd.DataFrame(
                        MapServiceIntersectionCatalogue(username, password, grant_type,
                                                        url_choice, API_call_dict, layer_dict,
                                                        SpatialFeatureChoice))
                    if df_MapServiceIntersections.empty:
                        flash("Something went wrong with the MapServiceIntersectionsCatalogue API call.")
                    else:
                        flash(f"Successfull API call on the {SpatialFeatureChoice} from the {org_choice} CP3 system.\n"
                              f"You may now download a spatial feature report (MS Word) on {SpatialFeatureChoice}.")
                        nav_stage = 3
                else:
                    SpatialFeatureChoice = ""
                    flash("You need to select spatial feature.")

                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var,
                                       SpatialFeatureChoice=SpatialFeatureChoice)

            elif button_dlreport is not None and nav_stage == 3:  # Calling the report
                SpatialFeatureChoice = request.form['inputGroupSelect01']
                # All APIs have now been called so the report building can proceed
                # Change the FeatureClassName column to "category" type
                df_MapServiceIntersections['FeatureClassName'].astype("category")
                # See how many 'no intersects' are there
                no_intersects = df_MapServiceIntersections[
                    df_MapServiceIntersections['FeatureClassName'] == 'No Intersect'].count()[0]
                # Out of...
                total_datapoints = len(df_MapServiceIntersections.index)
                if no_intersects < total_datapoints:  # Check if there are spatial intersects
                    # Continue as normal if there are spatial intersects
                    all_well += 1  # This really is just for problem/error trapping at this stage, this var does not get used anymore
                    intersecting = total_datapoints - no_intersects
                    # Create a small dataframe (this if for plotly)
                    data = {'NoIntersects': no_intersects, 'Intersecting': intersecting}
                    df_Intersects = pd.DataFrame(data, index=[0])  # the `index` argument is important
                    df_Intersects2 = df_Intersects.T
                    df_Intersects2['Projects'] = df_Intersects2[0]
                    del df_Intersects2[0]
                    df_Intersects2['Intersects'] = df_Intersects2.index
                    # Create dictionary with spatial features as key and projects in that feature with their % intersect as values
                    # Create a list containing the unique Spatial entities available in the dataset. This will enable an iteration through them later
                    list_of_features = list(df_MapServiceIntersections['FeatureClassName'].unique())
                    # Remove 'no intersect' because they have no spatial property and the overwhelm ito numbers in many datasets
                    list_of_features.remove('No Intersect')
                    # Create a list for all the fin years in the system
                    list_of_years = list(df_CapexBudgetDemandCatalogue['Interval'].unique())
                    list_of_years = sorted(list_of_years)
                    # This variable stores how many of the chosen feature there is
                    chosen_feature_qty = len(list_of_features)

                    # Initialise the master dictionary
                    feature_intersect_dict = {}
                    # Create a master dictionary with each spatial feature as a key
                    # Each key (e.g. ward) contains another dictionary with project number as key and percentage intersect as value
                    sub_frame = {}
                    for feature in list_of_features:
                        sub_frame[feature] = df_MapServiceIntersections[
                            df_MapServiceIntersections["FeatureClassName"] == feature]
                        project_dict = {}
                        for row in sub_frame[feature].itertuples():
                            project_dict[row.ProjectId] = row.PercentageIntersect
                        feature_intersect_dict[feature] = project_dict
                    # Now this dictionary can be used to query spatial feature to get to the projects and their intersects.
                    # print(feature_intersect_dict['Ward 100']) -> All projects in a ward is given with their % intersect

                    # Put this entire dictionary in a dataframe. Projects with their %Intersect per ChosenSpatialFeature
                    df_FeatureIntersectPer = pd.DataFrame(feature_intersect_dict)
                    # Replace all the NaN's with zeros
                    df_FeatureIntersectPer.replace(np.nan, 0, inplace=True)

                    # ********************************************************************************************
                    # Build df_EntireSet with columns for each fin year
                    # 1) Chosen Feature 2) No of Projs in Chosen Feature 3) Total Capital Demand in Chosen Feature 4) Capital Demand in Chosen Feature per Year
                    list_nr = []  # List with number of projects in a ward
                    list_cost = {}  # list_cost[2022], etc
                    list_cost['Total'] = []
                    list_cost['MTREF'] = []
                    for year in list_of_years:
                        list_cost[year] = []
                    mask3_dict = {}
                    # There will bae a dataframe for each mask for each year
                    df_mask3_dict = {}  # The will be a dataframe for each year so they will be store in a dictionary
                    df_perward = {}
                    project_year_total = {}

                    for feature in list_of_features:
                        # The number of project per chosen spatial feature goes into the 'list_nr' list
                        list_nr.append(len(
                            df_MapServiceIntersections[df_MapServiceIntersections['FeatureClassName'] == feature][
                                'ProjectId']))
                        # Now get the list of projects within the chosen feature (per feature)
                        df_perward[feature] = df_FeatureIntersectPer[df_FeatureIntersectPer[feature] > 0][feature]
                        # Set the variables that will keep track of the budget per spatial feature to zero
                        project_total = 0
                        mtref_total = 0

                        for year in list_of_years:
                            project_year_total[year] = 0

                        # Now iterate through the project numbers
                        for index in df_perward[feature].index:  # 'index' is the project number
                            # The 'df_perward[feature]' ensures the project exist in that ward
                            # Create a mask for the project number represented by 'index'
                            mask1 = df_CapexBudgetDemandCatalogue['ProjectId'] == index  # Mask per project
                            df_mask1 = df_CapexBudgetDemandCatalogue[
                                mask1]  # A dataframe for that mask for just that project ('index')

                            # mask3 is for a specific year
                            for year in list_of_years:
                                mask3_dict[year] = df_mask1['Interval'] == year  # There is a mask for each year
                                df_mask3_dict[year] = df_mask1[mask3_dict[
                                    year]]  # Create a dataframe for each year 1) Project -> 2) Year
                                try:
                                    project_year_total[year] += float(df_mask3_dict[year]['Amount'].sum()) * \
                                                                df_perward[feature][index]
                                    project_total += float(df_mask3_dict[year]['Amount'].sum()) * df_perward[feature][
                                        index]
                                    if year in list_of_years[0:3]:
                                        mtref_total += float(df_mask3_dict[year]['Amount'].sum()) * \
                                                       df_perward[feature][index]
                                except IndexError:  # If there is no year, an index error is returned. The populate the list with a zero.
                                    project_year_total[year] += 0
                                    project_total += 0
                                    if year in list_of_years[0:3]:
                                        mtref_total += 0

                        list_cost['Total'].append(project_total)
                        list_cost['MTREF'].append(mtref_total)
                        for year in list_of_years:
                            list_cost[year].append(project_year_total[year])

                    # With these lists a new dataframe can be created to plot
                    # Thus, create a dataframe/dataframes containing all the spatial feautures selected, each containing
                    # the number of projects in that feature and the capital demand per that feature
                    # Decide on the number of data sets depending on the size of the data

                    # 1st Create Dataset of all the data to split up for graphing purposes
                    df_EntireSet = pd.DataFrame(
                        {SpatialFeatureChoice: list_of_features, f'Projects per {SpatialFeatureChoice}': list_nr,
                         'Capital All Years': list_cost['Total'], 'Capital MTREF': list_cost['MTREF']})
                    column_name_list = []
                    for year in list_of_years:
                        column_name_list.append(f'Capital {year}')
                        df_EntireSet[f'Capital {year}'] = list_cost[year]

                    # Now sort the dataset in order of number of projects from largest to smallest
                    df_EntireSet.sort_values(f'Projects per {SpatialFeatureChoice}', inplace=True, ascending=True)
                    # Modify df_EntireSet by adding columns to rank
                    df_EntireSet["CapAllRank"] = df_EntireSet['Capital All Years'].rank(ascending=False)
                    df_EntireSet["CapMTREFRank"] = df_EntireSet['Capital MTREF'].rank(ascending=False)
                    df_EntireSet["NoProjectsRank"] = df_EntireSet[f'Projects per {SpatialFeatureChoice}'].rank(
                        ascending=False)

                    # Wrap all the loose variables in a dictionary for use in the report
                    var_dict = {}
                    # var_dict['username'] = sys_username
                    var_dict['url_choice'] = url_choice
                    var_dict['org_choice'] = org_choice
                    var_dict['entity_choice'] = entity_choice
                    var_dict['SpatialFeatureChoice'] = SpatialFeatureChoice
                    var_dict['Layer_List'] = layer_list
                    var_dict['total_datapoints'] = total_datapoints
                    var_dict['intersecting'] = intersecting
                    var_dict['no_intersects'] = no_intersects
                    var_dict['chosen_feature_qty'] = chosen_feature_qty
                    var_dict['column_name_list'] = column_name_list
                    var_dict['list_of_years'] = list_of_years
                    var_dict['list_of_features'] = list_of_features

                    nav_stage = 2  # To allow the user to once more select a field
                    # Check the growth of files and reduce it
                    control_growth_of_docx()
                    # Now create the spatial feature report
                    path = create_worddoc(var_dict=var_dict, baseline_dict=baseline_cat_dict,
                                          df_intersects2=df_Intersects2,
                                          df_EntireSet=df_EntireSet, df_perward=df_perward,
                                          df_MapServiceIntersections=df_MapServiceIntersections,
                                          df_CapexBudgetDemandCatalogue=df_CapexBudgetDemandCatalogue,
                                          df_ProjectCatalogue=df_ProjectCatalogue)
                    return send_file(path, as_attachment=True)
                else:
                    nav_stage = 2
                    flash(f"There are no spatial intersects on {SpatialFeatureChoice}. A spatial feature report can "
                          f"therefore not be generated.")
                    return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                           url_list=url_list, spatial_var=spatial_var,
                                           SpatialFeatureChoice=SpatialFeatureChoice)
            elif button_dlreport is not None and nav_stage == 2:  # User pressed the report button again but did not load the spatial features again
                nav_stage = 2
                flash("You selected to download a report again without having loaded the variables required for your "
                      "newly selected spatial feature selection. Please press the '2. Call selected spatial feature "
                      "variables' before attempting to download a spatial report again.")
                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var,
                                       SpatialFeatureChoice=SpatialFeatureChoice)
            elif button_url_upload is not None:  # User wishes to upload a json file
                file = request.files['file']
                if file.filename == '':  # File was not selected
                    flash('No selected file to upload.')
                elif file and allowed_file(file.filename):  # File extension checked and is ok ('.xlsx')
                    # Step 1: Upload the .xlsx file and save it under the upload folder
                    filename = secure_filename(file.filename)
                    upload_path_var = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(upload_path_var)
                    # Step 2: Read the file into a variable
                    returned_data = read_from_excel(upload_path_var)
                    if 'org' in returned_data[0].keys():  # The right kind of file was uploaded
                        # Step 3: Write the variable into the json file used by the system 'CP3_deployments.json'
                        write_to_JSON_file(returned_data, url_vars_file_path)
                        # Step 4: Re-populate the variables
                        returned_combined_list = vars_from_json_file(url_vars_file_path)
                        org_list = returned_combined_list[1]
                        entityname_list = returned_combined_list[2]
                        url_list = returned_combined_list[3]
                        json_file_ok = returned_combined_list[0][0]
                        flash('New list uploaded!')
                    else:  # The user attempted to upload the wrong file
                        flash('Caution! It appears that you have uploaded an incorrect file. Process abandoned.')
                else:  # The only remaining option is that the wrong file extension was chosen.
                    flash("Wrong extension type for upload. Must be '.xlsx'")

                # Check the growth of files in the folder and reduce it
                control_growth_of_xlsx(f'.{UPLOAD_FOLDER}/*.xlsx')

                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var,
                                       SpatialFeatureChoice=SpatialFeatureChoice)

            elif button__url_download is not None:  # User wishes to download the MSExcel file with all the sites
                # Check the growth of files in the folder and reduce it
                control_growth_of_xlsx(f'.{DOWNLOAD_FOLDER}/*.xlsx')
                # Step 1: Read file at 'url_vars_file_path'
                json_var_from_file = read_json_file(url_vars_file_path)
                # Step 2: Convert to Excel
                excel_file_path = write_to_xls(json_var_from_file, 'url')
                # Step 3: Download to browser
                flash('Template downloaded.')
                return send_file(excel_file_path, as_attachment=True)

            elif button_api_upload is not None:  # User wishes to upload new API credentials
                file = request.files['file']
                if file.filename == '':  # File was not selected
                    flash('No selected file to upload.')
                elif file and allowed_file(file.filename):  # File extension checked and is ok ('.xlsx')
                    # Step 1: Upload the .xlsx file and save it under the upload folder
                    filename = secure_filename(file.filename)
                    upload_path_var = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(upload_path_var)
                    # Step 2: Read the file into a variable
                    returned_data = read_from_excel(upload_path_var)
                    if 'username' in returned_data[0].keys():  # The right kind of file was uploaded
                        # Step 3: Write the variable into the json file used by the system 'CP3_deployments.json'
                        write_to_JSON_file(returned_data, API_vars_file_path)
                        # Step 4: Re-populate the variables
                        # Get API creds from the variables json file
                        API_file_name = "API_Profile.json"
                        API_vars_file_path = f"./Variables/{API_file_name}"
                        API_vars = API_vars_json_file(API_vars_file_path)
                        username = API_vars[0]
                        password = API_vars[1]
                        grant_type = API_vars[2]
                        flash('New API credentials uploaded!')
                    else:
                        flash('Caution! It appears that you have uploaded an incorrect file. Process abandoned.')
                else:  # The only remaining option is that the wrong file extension was chosen.
                    flash("Wrong extension type for upload. Must be '.xlsx'")

                # Check the growth of files in the folder and reduce it
                control_growth_of_xlsx(f'.{UPLOAD_FOLDER}/*.xlsx')

                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var,
                                       SpatialFeatureChoice=SpatialFeatureChoice)

            elif button__api_download is not None:  # User wishes to donwload the existing API credentials
                control_growth_of_xlsx(f'.{DOWNLOAD_FOLDER}/*.xlsx')
                # Step 1: Read file at 'url_vars_file_path'
                json_var_from_file = read_json_file(API_vars_file_path)
                # Step 2: Convert to Excel
                excel_file_path = write_to_xls(json_var_from_file, 'api')
                # Step 3: Download to browser
                flash('API Credentials downloaded.')
                return send_file(excel_file_path, as_attachment=True)

        else:  # Get not Post, in other words when it lands
            return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                   url_list=url_list, spatial_var=spatial_var,
                                   SpatialFeatureChoice=SpatialFeatureChoice)

    else:  # json_file_ok == False
        if request.method == 'GET':
            flash(returned_combined_list[0][1])
        return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                               url_list=url_list, spatial_var=spatial_var,
                               SpatialFeatureChoice=SpatialFeatureChoice)


# This is required for the programme to run
if __name__ == '__main__':  # This runs the app and starts the server that allows it to receive connections
    # app.run(host="localhost", port=8080, debug=True)
    serve(app, host='0.0.0.0', port=8080)
