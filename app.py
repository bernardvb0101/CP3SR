from flask import Flask, render_template, request, flash, send_file
import os
import pandas as pd
import numpy as np
from Utilities.url_exists import URL_exists
from Utilities.control_growth import control_growth_of_docx
from Utilities.get_url_vars import vars_from_json_file
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

show_drop_down = False
nav_stage = 1
all_well = 0
if json_file_ok:
    org_choice = org_list[0]  # Make the 1st one in the list the default
    url_choice = url_list[0]  # Make the 1st one in the list the default
    entity_choice = entityname_list[0]  # Make the 1st one in the list the default
else:
    org_choice = []
    url_choice = []
    entity_choice = []
spatial_var = []
sys_username = "Bernard"
# This part of the credentials for the API call (to my understanding) is default allways "password"
grant_type = "password"

app = Flask(__name__)  # to make the app run without any
app.config['SECRET_KEY'] = os.urandom(24)
app.config['DOWNLOAD_FOLDER'] = "/DOWNLOAD_FOLDER"


# This route is the "home" route that redirects immediately to "home_in.html"
@app.route('/', methods=['GET', 'POST'])
def home():
    global show_drop_down, all_well, url_choice, url_list, spatial_var, username, password, grant_type, org_choice
    global API_call_dict, layer_dict, SpatialFeatureChoice, SpecificFeature, spatial_var, entityname_list, entity_choice
    global layer_list, number_of_plots, nav_stage
    global baseline_cat_dict, df_ProjectCatalogue, df_CapexBudgetDemandCatalogue, df_MapServiceLayerCatalogue
    global df_MapServiceIntersections, no_intersects, total_datapoints, intersecting, df_Intersects2, sys_username

    if json_file_ok:
        if request.method == 'POST' and not show_drop_down:  # Step 1: Pressed the submit button with username and pw, no dropdown yet

            # Get the status of the buttons that were pressed
            button_1stAPI = request.form.get("call_1st_APIs")
            button_2ndAPI = request.form.get("call_2nd_APIs")
            button_dlreport = request.form.get("download_report")

            if button_1stAPI is not None and nav_stage == 1:  # Pressed the button to call an API
                # The all_well variable is zero if no APIs were returned successfully yet
                all_well = 0
                # The spatial_var is a list with the spatial feautures that the user cna select from later on
                spatial_var = []

                # Get the credentials for the API call from the user
                username = request.form.get('username')
                password = request.form.get('password')

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
                    show_drop_down = True
                    flash(f"Successfull API call on {org_choice} CP3 system.\n"
                          f"Select spatial feature and initiate 2nd API call.")
                    nav_stage = 2
                else:
                    flash(f"The call for data from the {org_choice} system was not successful. This may be because the"
                          f"API profile for {org_choice} relating to the username and password that you used, may not"
                          f"be correctly set up or one or more of the data catalogues returned an 'empty' response."
                          f" Try using a different site and query to see if the problem persists or whether it is "
                          f"specific to this site and your user profile replated to this site.")
                # If nav_stage is 2, the user would be able to call the 2nd APIs, otherwise he will just get a fault
                # message and stay on the page with the url choices and username and password
                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var)

            elif button_1stAPI is not None and nav_stage == 2:  # Pressed the button to select a another site
                nav_stage = 1
                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var)

            elif button_2ndAPI is not None and nav_stage == 2: # Calling the 2nd set of APIs
                SpatialFeatureChoice = request.form['inputGroupSelect01']

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

                return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                       url_list=url_list, spatial_var=spatial_var)

            elif button_dlreport is not None and nav_stage == 3:  # Calling the 2nd set of APIs
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
                    # This variable stores how many of the chosen feature there is
                    # *******************************************************************************************************
                    # This variable gets created to understand how many features will be discplayed
                    chosen_feature_qty = len(list_of_features)
                    # Initialise the master dictionary
                    feature_intersect_dict = {}
                    # Create a master dictionary with each spatial feature as a key
                    # Each key (e.g. ward) contains another dictionary with project number as key and percentage intersect as value
                    for feature in list_of_features:
                        sub_frame = df_MapServiceIntersections[
                            df_MapServiceIntersections["FeatureClassName"] == feature]
                        project_dict = {}
                        for row in sub_frame.itertuples():
                            project_dict[row.ProjectId] = row.PercentageIntersect
                        feature_intersect_dict[feature] = project_dict
                    # Now this dictionary can be used to query spatial feature to get to the projects and their intersects.
                    # print(feature_intersect_dict['Ward 100'])

                    # Put this entire dictionary in a dataframe
                    # this dataaframe contains
                    df_FeatureIntersectPer = pd.DataFrame(feature_intersect_dict)
                    # Replace all the NaN's with zeros
                    df_FeatureIntersectPer.replace(np.nan, 0, inplace=True)

                    # Also, Create a 2 lists with 1) the number of projects per spatial feature and 2) the budget per spatial feature
                    list_nr = []
                    list_cost = []
                    for feature in list_of_features:
                        # Create the list containing how many project per wards are there
                        list_nr.append(len(
                            df_MapServiceIntersections[df_MapServiceIntersections['FeatureClassName'] == feature][
                                'ProjectId']))
                        # Create a list of the total budget per ward asked

                        # Before you determine the subtotat for that spatial feature, reset the sub_total variable to 0
                        sub_total = 0
                        # iterate over the sub-dictionaries
                        for key, value in feature_intersect_dict[feature].items():
                            # Get the budget of each project and multiply with the % overlap
                            try:
                                sub_total += df_CapexBudgetDemandCatalogue.loc[
                                                 df_CapexBudgetDemandCatalogue['ProjectId'] == key, 'Amount'].iloc[
                                                 0] * value
                            except IndexError:  # There is no budget demand for this project
                                pass
                        # Now this total can be added to the list
                        list_cost.append(sub_total)
                    # With these lists a new dataframe can be created to plot
                    # Thus, create a dataframe/dataframes containing all the spatial feautures selected, each containing the
                    # number of projects in that feature and the capital demand per that feature
                    # Decide on the number of data sets depending on the size of the data

                    # 1st Create Dataset of all the data to split up for grpahing purposes
                    df_EntireSet = pd.DataFrame(
                        {SpatialFeatureChoice: list_of_features, f'Projects per {SpatialFeatureChoice}': list_nr,
                         f'Capital Demand': list_cost})
                    # Now sort the dataset in order of number of projects from largest to smallest
                    df_EntireSet.sort_values(f'Projects per {SpatialFeatureChoice}', inplace=True, ascending=True)
                    """
                    df_EntireSet looks like this: (df_subsets also!)
                        City of Tshwane Wards	Projects per City of Tshwane Wards	Capital Demand per City of Tshwane Wards
                    0	Ward 58	                85	                                5.605180e+08
                    1	Ward 66	                8	                                1.557022e+0
                    ...
                    """
                    # The splitting of the df_EntireSet dataframe will happen below, depending on 'number_of_plots' variable
                    if chosen_feature_qty <= 40:  # Only one dataset
                        number_of_plots = 1
                        df_subset1 = df_EntireSet
                        df_subset2 = {}
                        df_subset3 = {}
                        df_subset4 = {}
                    elif chosen_feature_qty > 40 and chosen_feature_qty <= 80:  # Create 2 data sets
                        number_of_plots = 2
                        dfs = np.array_split(df_EntireSet, number_of_plots)
                        df_subset1 = dfs[0]
                        df_subset2 = dfs[1]
                        df_subset3 = {}
                        df_subset4 = {}
                    elif chosen_feature_qty > 80 and chosen_feature_qty <= 120:  # Create 3 data sets
                        number_of_plots = 3
                        dfs = np.array_split(df_EntireSet, number_of_plots)
                        df_subset1 = dfs[0]
                        df_subset2 = dfs[1]
                        df_subset3 = dfs[2]
                        df_subset4 = {}
                    else:  # Create 4 plots
                        number_of_plots = 4
                        dfs = np.array_split(df_EntireSet, number_of_plots)
                        df_subset1 = dfs[0]
                        df_subset2 = dfs[1]
                        df_subset3 = dfs[2]
                        df_subset4 = dfs[3]

                    # Wrap all the loose variables in a dictionary for use in the report
                    var_dict = {}
                    var_dict['username'] = sys_username
                    var_dict['url_choice'] = url_choice
                    var_dict['entity_choice'] = entity_choice
                    var_dict['SpatialFeatureChoice'] = SpatialFeatureChoice
                    var_dict['Layer_List'] = layer_list
                    var_dict['total_datapoints'] = total_datapoints
                    var_dict['intersecting'] = intersecting
                    var_dict['no_intersects'] = no_intersects
                    var_dict['chosen_feature_qty'] = chosen_feature_qty

                    # Check the growth of files and reduce it
                    control_growth_of_docx()
                    # Now create the spatial feature report
                    path = create_worddoc(var_dict=var_dict, baseline_dict=baseline_cat_dict,
                                          df_project_cat=df_ProjectCatalogue, df_intersects2=df_Intersects2,
                                          df_subset1=df_subset1, df_subset2=df_subset2, df_subset3=df_subset3,
                                          df_subset4=df_subset4, number_of_plots=number_of_plots,
                                          df_EntireSet=df_EntireSet)
                    return send_file(path, as_attachment=True)
                else:
                    nav_stage = 2
                    flash(f"There is no spatial intersect on {SpatialFeatureChoice}. A spatial feature report can "
                          f"therefore not be generated.")
                    return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                           url_list=url_list, spatial_var=spatial_var)

    else:  # json_file_ok == False
        if request.method == 'GET':
            flash(returned_combined_list[0][1])
            # If all the API's were called successfully, show the DropDown
        return render_template('home.html', nav_stage=nav_stage, url_choice=url_choice,
                                   url_list=url_list, spatial_var=spatial_var)


# This is required for the programme to run
if __name__ == '__main__':  # This runs the app and starts the server that allows it to receive connections
    app.run(host="localhost", port=8080, debug=True)
    #serve(app, host='0.0.0.0', port=5000)
