import json
from Utilities.file_exists import file_exists


# Function reads the source file provided and return file content as a list
def read_source_file(jsn_file_name):
    """
    This function reads a JSON file into a list of dictionaries. It first checks whether the file it needs to
    read exists.
    """
    if file_exists(jsn_file_name):
        try: #even if the file exists - it may not be in the right format
            with open(jsn_file_name, "r") as data_file:
                file_content = json.load(data_file)
            # returns a list of dictionaries : [{'Date': '2018/11/30', 'Category': 'Fuel', 'Cost': '1665,9',.....
            return_list =[True, file_content]
            return return_list
        except UnicodeDecodeError:
            problem_message = "The json file with the API credentials is corrupted. This program will " \
                              "not be able to execute. Please email the developer at info@novus3.co.za."
            return_list = [False, problem_message]
            return return_list
    else:
        problem_message = "The json file with the API credentials is corrupted. This program will not be able to" \
                          " execute. Please email the developer at info@novus3.co.za."
        return_list = [False, problem_message]
        return return_list

def API_vars_json_file(json_source_file):
    """
    This function gets the json variables from the json reader and reads them into separate variables
    """
    return_list = read_source_file(json_source_file)
    file_status = return_list[0]
    API_creds = return_list[1][0]
    tup = ()

    def un_dict(username, password, grant_type):
        return (username, password, grant_type)

    if file_status:  # If nothing is wrong with the file
        tup = un_dict(**API_creds)

    return tup
