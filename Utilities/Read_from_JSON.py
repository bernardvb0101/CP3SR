import json
from Utilities.Fix_Data_Types import fix_data
from Utilities.file_exists import file_exists


# Function reads the source file provided and return file content as a list
def read_json_file(jsn_file_name):
    """
    This function reads a JSON file into a list of dictionaries. It first checks whether the file it needs to
    read exists.
    """
    if file_exists(jsn_file_name):
        try: #even if the file exists - it may not be in the right format
            with open(jsn_file_name, "r") as data_file:
                file_content = json.load(data_file)
            # returns a list of dictionaries : [{'Date': '2018/11/30', 'Category': 'Fuel', 'Cost': '1665,9',.....
            # which is then fixed with fix_data to be the true data types: [{'Date': '2018/11/30', 'Category': 'Fuel', 'Cost': 1665.9,.....
            return fix_data(file_content)
        except UnicodeDecodeError:
            print("--------------------------------------------------------------")
            print("This file format appears not to be a json file. Import failed.")
            print("--------------------------------------------------------------")
    else:
        print("-------------------------------------------------------------")
        print(f"Data file '{jsn_file_name}' not found.")
        print("-------------------------------------------------------------")

