import pandas
from Utilities.Fix_Data_Types import fix_data
from Utilities.file_exists import file_exists

def read_from_excel(file_name):
    """
    Function that read the content from an excel file into a list of dictionaries
    """
    if file_exists(file_name):
        try:
            raw_data = pandas.read_excel(file_name)
            data = raw_data.fillna('').to_dict('records') #.fillna gets rid of the "nan" in the mepty fields
            data = fix_data(data)
            return data
        except ValueError:
            print("-------------------------------------------------------------------------------------------")
            print("File is not a recognized excel file. The ligitimate Excel file must have a .xlxs extension.")
            print("-------------------------------------------------------------------------------------------")
    else:
        print(f"Data file '{file_name}' not found.")
