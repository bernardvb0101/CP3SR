from flask import Flask, render_template, request, url_for, redirect, flash, session, send_file
import os
import pandas as pd
from Utilities.url_exists import URL_exists
# from CP3_API_calls.CP3_API_scaffolding import get_token
from CP3_API_calls.Create_API_Variables import create_vars
from CP3_API_calls.BaselineCatalogue import baseline_catalogue
from CP3_API_calls.ProjectCatalogue import ProjectCatalogue
from CP3_API_calls.CapexDemandCatalogue import CapexDemandCatalogue
from CP3_API_calls.MapServiceLayersCatalogue import MapServiceLayersCatalogue
from CP3_API_calls.MapServiceIntersectionCatalogue import MapServiceIntersectionCatalogue
from Variables.available_urls import url_list

# global variables
show_drop_down = False
all_well = 0
url_choice = url_list[0]  # Make the 1st one in the list the default
spatial_var = []

app = Flask(__name__)  # to make the app run without any
app.config['SECRET_KEY'] = os.urandom(24)


# This route is the "home" route that redirects immediately to "home_in.html"
@app.route('/', methods=['GET', 'POST'])
def home():
    global show_drop_down, all_well, url_choice, url_list, spatial_var, username, password, full_url, grant_type
    global API_call_dict, layer_dict
    if request.method == 'POST' and not show_drop_down:  # Pressed the submit button with username and pw
        # The all_well variable is zero if no APIs were returned successfully yet
        all_well = 0
        spatial_var = []

        # Get the crednetials for the API call from the user
        username = request.form.get('username')
        password = request.form.get('password')
        full_url = request.form['flexRadioDefault']
        grant_type = "password"

        # Now assign a value to 'url_choice' based on the url that was chosen
        url_choice = full_url

        if username != "" and password != "" :
            # flash(f"Username: {username}")
            # flash(f"Password: {password}")
            # Test the URL that was selected to see if it is active
            if URL_exists(full_url):
                url_ok = True
            else:
                url_message = f"{full_url} was tested and did not return an active response as expected."
                url_ok = False
                flash(url_message)
            if url_ok:
                # Call the Help API to create variables
                API_call_dict = create_vars(username, password,grant_type, full_url)
                if not API_call_dict:  # If the dictionary is returned empty, tell the user
                    flash("Something went wrong with the API calls.")
                    flash("This may be a wrong username/password OR the API credentials are not set up correctly/at all.")
                else:  # Returned values from API variables call seems ok so process can continue
                    all_well += 1
                    # Call the Baseline Catalogue and get info about the baseline
                    baseline_cat_dict = baseline_catalogue(username, password, grant_type, full_url, API_call_dict)
                    if not baseline_cat_dict:
                        flash("Something went wrong with the BaselineCatalogue API call.")
                    else:
                        all_well += 1
                    # Call the ProjectCatalogue and create a dataframe
                    df_ProjectCatalogue = pd.DataFrame(ProjectCatalogue(username,password, grant_type, full_url, API_call_dict))
                    if df_ProjectCatalogue.empty:
                        flash("Something went wrong with the ProjectCatalogue API call.")
                    else:
                        all_well += 1
                    # Call the CapexDemandCatalogue and create a dataframe
                    df_CapexBudgetDemandCatalogue = pd.DataFrame(CapexDemandCatalogue(username, password, grant_type, full_url, API_call_dict, baseline_cat_dict['APIAccessTag']))
                    if df_CapexBudgetDemandCatalogue.empty:
                        flash("Something went wrong with the CapexBudgetDemandCatalogue API call.")
                    else:
                        all_well += 1
                    # Call the MapServiceLayersCatalogue
                    return_list = MapServiceLayersCatalogue(username, password, grant_type, full_url, API_call_dict)

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
        return render_template('home.html', show_drop_down = show_drop_down, url_choice = url_choice,
                               url_list = url_list, spatial_var = spatial_var)
    elif request.method == 'POST' and show_drop_down: # Pressed the submit button with spatial features selected
        SpatialFeatureChoice = request.form['inputGroupSelect01']

        print(SpatialFeatureChoice)
        # Call the MapServiceIntersectionCatalogue
        df_MapServiceIntersections = pd.DataFrame(MapServiceIntersectionCatalogue(username, password, grant_type,
                                                                                  full_url, API_call_dict, layer_dict,
                                                                                  SpatialFeatureChoice))
        if df_MapServiceIntersections.empty:
            flash("Something went wrong with the MapServiceIntersectionsCatalogue API call.")
        else:
            all_well += 1
            flash(f"all_well = {all_well},  SpatialFeatureChoice = {SpatialFeatureChoice}")
        return render_template('home.html', show_drop_down=show_drop_down, url_choice=url_choice,
                               url_list=url_list, spatial_var=spatial_var)
    else:
        # If all the API's were called successfully, show the DropDown
        if all_well == 5:
            show_drop_down = True
        return render_template('home.html', show_drop_down = show_drop_down, url_choice = url_choice,
                               url_list = url_list, spatial_var = spatial_var)


# This is required for the programme to run
if __name__ == '__main__':  # This runs the app and starts the server that allows it to receive connections
    app.run(host="localhost", port=8080, debug=True)
    #serve(app, host='0.0.0.0', port=5000)
