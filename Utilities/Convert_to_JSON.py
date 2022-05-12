import json


def write_to_JSON_file(output,filename):
    """
    This function asks the name of the JSON file it needs to write to. I then writes a list of dictionaries to JSON.
    """
    new_json_file_name = filename
    file = open(new_json_file_name, 'w')
    json.dump(output, file)
    file.close()
