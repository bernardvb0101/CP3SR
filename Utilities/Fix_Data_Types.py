import ast

def create_dict_and_zip_headings(headings, raw_content):
    """
    This function require that a list of headings and a list of content gets passed to it.
    It then zips this together and creates a list of dictionaries.
    """
    new_content=[]
    for row in raw_content:
        new_line = row
        new_content_dict = dict(zip(headings, new_line))
        new_content.append(new_content_dict)
    return new_content


def evaluate(x):
    """
    This function evaluates the data that gets passed to it and returns it in its correct format.
    """
    try:
        return ast.literal_eval(x)
    except (ValueError,SyntaxError):
        return x


def fix_data(input_data):
    """
    This function goes through the entire list of dictionaries passed to it and translates the data to its correct
    data types. E.g. '123,4' becomes 123,4
    """
    transformed_data_line = []
    transformed_data = []
    header = input_data[0].keys()

    for outp in input_data:
        # 1) Convert list into a dictionary
        a_dictionary = dict(outp)
        # 2) Extract only the values (without the keys)
        a_view = a_dictionary.values()
        # 3) Convert this into a tuple
        a_tuple = tuple(a_view)

        i = 0
        while i < len(a_tuple):
            item_str = a_tuple[i]
            try: # When the feal is read from SQlite3, the are no ","'s so this line will return an error. So "try"
                item_str = item_str.replace(',','.')
            except: # When there is an error, do not return it, just continue, because it means the data is ok
                pass
            finally: # "finally" always gets done.
                transformed_data_line.append(evaluate(item_str))
                i = i + 1
        else:
            transformed_data.append(transformed_data_line)
            transformed_data_line = []

    transformed_data = create_dict_and_zip_headings(header, transformed_data)
    return transformed_data






