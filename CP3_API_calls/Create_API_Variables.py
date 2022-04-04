# ********************  STEP 2 ********************
# 2. Scaffolding to enable API calls (Part 2 - Get further variables from the API Help call)
# This will create a list and a dictionary variable
# Dictionary variable is: "API_calls_dict"
# List variable is: "API_calls_list"

from CP3_API_calls.CP3_API_Classes import CP3Client


def create_vars(username, password, grant_type, url_choice):
    # Get the token
    client = CP3Client(username, password, grant_type, url_choice)
    token = client.get_API_token
    # The 'get_API_help' method already extracts the return value as text using the ".text" attribute
    help = client.get_API_help
    # Convert the string into a list of seperate items
    help_API_list = help.splitlines()
    # Create lists for the i) Main API call (e.g. BaselineCatalogue) ii) Output Type (e.g. JSON or Text) iii) Parameter
    # 3. Scaffolding to enable API calls (Part 2 - Get further variables from the API Help call)
    # Create the empty lists
    call1 = []
    output_type = []
    parameter1 = []
    API_calls_dict = {}
    # Copy the content of the help_API_list that was created above into each variable and manipulate and reduce further.
    for call in help_API_list[8:]:
        call1.append(call.split(" ")[1].split(":")[0].split("/")[1].split("?")[0])
        if call.split(" ")[1].split(":")[1]:
            output_type.append(call.split(" ")[1].split(":")[1])
        else:
            output_type.append("text")
        try:
            parameter1.append(call.split(" ")[1].split("?")[1].split(":")[0])
        except IndexError:
            parameter1.append("No parameters")
    # Create a combined list comprising of Output Type and Parameters in preparation of creating a dictionary using
    # the Main API call as key value to each dictionary entry
    # Build a dictionary
    combined = []
    key_value = ""
    for number in range(len(call1)):
        # print(f"{number}.1" + " " + call1[number].split(" ")[1].split(":")[0].split("/")[1].split("?")[0])
        key_value = call1[number]
        combined.append(output_type[number])
        combined.append(parameter1[number])
        API_calls_dict[key_value] = combined
        combined = []

    return API_calls_dict
